import streamlit as st
import pandas as pd
import json, os, base64, requests, re, threading
from io import BytesIO
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill
import matplotlib.pyplot as plt

# === CONFIG ===
st.set_page_config(page_title="Studio ISA - Alcyon Italia", layout="wide")
GITHUB_FILE = os.getenv("GITHUB_FILE", "keywords_memory.json")
GITHUB_REPO = os.getenv("GITHUB_REPO")
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")

# === CACHE LETTURA EXCEL ===
@st.cache_data(show_spinner=False, ttl=600)
def load_excel(f):
    return pd.read_excel(f)

# === UTILS ===
def norm(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip().lower())

def any_kw_in(t, kws):
    return any(k in t for k in kws)

def coerce_numeric(series: pd.Series) -> pd.Series:
    """Converte in numerico gestendo virgole, simboli e stringhe."""
    # replace euro/spazi/punti mille e virgole decimali
    s = series.astype(str).str.replace(r"[€\s]", "", regex=True) \
                          .str.replace(".", "", regex=False) \
                          .str.replace(",", ".", regex=False)
    return pd.to_numeric(s, errors="coerce").fillna(0)

def drop_columns_robust(df: pd.DataFrame, names: list[str]) -> pd.DataFrame:
    """Drop colonne ignorando maiuscole/spazi/varianti."""
    to_drop = []
    norm_names = {norm(n) for n in names}
    for c in df.columns:
        if norm(c) in norm_names:
            to_drop.append(c)
    if to_drop:
        df = df.drop(columns=to_drop)
    return df

def round_pct_series(values: pd.Series) -> pd.Series:
    """Arrotonda come nel desktop: somma esattamente 100."""
    if values.sum() == 0:
        return values
    # calcolo percentuali precise con Decimal
    total = Decimal(str(values.sum()))
    raw = [Decimal(str(v)) * Decimal("100") / total for v in values]
    rounded = [x.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP) for x in raw]
    diff = Decimal("100.00") - sum(rounded)
    if len(rounded) > 0:
        rounded[-1] = (rounded[-1] + diff).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return pd.Series([float(x) for x in rounded], index=values.index)

# === REGOLE ===
RULES_A = {
    "LABORATORIO": ["analisi","emocromo","test","esame","coprolog","feci","giardia","leishmania","citolog","istolog","urinocolt","urine"],
    "VISITE": ["visita","controllo","consulto","dermatologic"],
    "FAR": ["meloxidyl","konclav","enrox","profenacarp","apoquel","osurnia","cylan","mometa","aristos","cytopoint","milbemax","stomorgyl","previcox","royal","stronghold","nexgard","procox"],
    "CHIRURGIA": ["intervento","chirurg","castraz","sterilizz","ovariect","detartrasi","estraz","biopsia","orchiettomia","odontostomat"],
    "DIAGNOSTICA PER IMMAGINI": ["rx","radiograf","eco","ecografia","tac"],
    "MEDICINA": ["terapia","terapie","flebo","day hospital","trattamento","emedog","cerenia","endovena","pressione"],
    "VACCINI": ["vacc","letifend","rabbia","trivalente","felv"],
    "CHIP": ["microchip","chip"],
    "ALTRE PRESTAZIONI": ["trasporto","eutanasia","unghie","cremazion","otoematoma","pet corner","ricette","medicazione","manualità"]
}

RULES_B = {
    "Visite domiciliari o presso allevamenti": ["visite domiciliari","allevamenti","domicilio"],
    "Visite ambulatoriali": ["visite ambulatoriali","terapia","trattamenti","vaccinazioni","ambulatorio","manualità","pet corner","visite","ricette","medicazione","microchip","controllo"],
    "Esami diagnostici per immagine": ["esami diagnostici per immagine","radiologia","eco","ecografia","tac","rx","raggi"],
    "Altri esami diagnostici": ["altri esami diagnostici","esami biochimici","laboratorio","malattie infettive","emocromo","prelievo"],
    "Interventi chirurgici": ["interventi chirurgici","avulsione","endoscopia","eutanasia","sedazione","anestesia","chirurgia","odontostomat","orchiettomia","asportazione","biopsia","ovariectomia"],
    "Assistenza al parto/ostetricia": ["assistenza al parto","ostetricia","parto"],
    "Attività di consulenza, perizia e collaborazione": ["attività di consulenza","perizia","collaborazione","telemedicina","consulto"],
    "Prestazioni di inseminazione artificiale": ["inseminazione artificiale"],
    "Altre attività": ["acconto"]
}

# === CLASSIFICAZIONE ===
def classify_A(desc, fam_val, mem):
    if pd.notna(fam_val) and str(fam_val).strip():
        return str(fam_val).strip()
    d = norm(desc)
    for k, v in mem.items():
        if norm(k) in d:
            return v
    for cat, keys in RULES_A.items():
        if any_kw_in(d, keys):
            return cat
    return "ALTRE PRESTAZIONI"

def classify_B(prest, cat_val, mem):
    if pd.notna(cat_val) and str(cat_val).strip():
        return str(cat_val).strip()
    d = norm(prest)
    for k, v in mem.items():
        if norm(k) in d:
            return v
    for cat, keys in RULES_B.items():
        if any_kw_in(d, keys):
            return cat
    return "Altre attività"

# === GITHUB HANDLERS ===
def github_load_json():
    try:
        if not (GITHUB_REPO and GITHUB_FILE):
            return {}
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_FILE}"
        headers = {"Authorization": f"token {GITHUB_TOKEN}"} if GITHUB_TOKEN else {}
        r = requests.get(url, headers=headers, timeout=15)
        if r.status_code == 200 and "content" in r.json():
            return json.loads(base64.b64decode(r.json()["content"]).decode("utf-8"))
    except Exception:
        pass
    return {}

def github_save_json_sync(data: dict) -> bool:
    """Salvataggio sincrono (affidabile) con esito True/False."""
    try:
        if not (GITHUB_REPO and GITHUB_FILE and GITHUB_TOKEN):
            st.warning("⚠️ Variabili GitHub mancanti: salvataggio cloud disattivato.")
            return False
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_FILE}"
        headers = {"Authorization": f"token {GITHUB_TOKEN}"}
        get_res = requests.get(url, headers=headers, timeout=15)
        sha = get_res.json().get("sha") if get_res.status_code == 200 else None
        encoded = base64.b64encode(json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")).decode("utf-8")
        payload = {"message": "Aggiornamento dizionario Studio ISA", "content": encoded, "branch": "main"}
        if sha:
            payload["sha"] = sha
        put_res = requests.put(url, headers=headers, data=json.dumps(payload), timeout=20)
        return put_res.status_code in (200, 201)
    except Exception:
        return False

def github_save_json_async(data: dict):
    def worker():
        ok = github_save_json_sync(data)
        if ok:
            st.toast("💾 Dizionario salvato sul cloud!")
        else:
            st.toast("⚠️ Salvataggio sul cloud non riuscito.")
    threading.Thread(target=worker, daemon=True).start()

# === MAIN ===
def main():
    st.title("📊 Studio ISA - Alcyon Italia")

    up = st.file_uploader("📁 Seleziona file Excel", type=["xlsx", "xls"])
    if not up:
        st.info("Carica un file per iniziare.")
        return

    if "df" not in st.session_state or up.name != st.session_state.get("last_file"):
        st.session_state.df = load_excel(up)
        st.session_state.user_memory = github_load_json()
        st.session_state.local_updates = {}
        st.session_state.idx = 0
        st.session_state.last_file = up.name

    # Copia df e rimuovi colonne extra (privato/professionista) in modo robusto
    df = st.session_state.df.copy()
    df = drop_columns_robust(df, ["Privato", "PROFESSIONISTA", "Professionista"])

    mem = st.session_state.user_memory
    updates = st.session_state.local_updates

    cols = [c.lower().strip() for c in df.columns]
    ftype = "B" if any("prestazione" in c for c in cols) and any("totaleimpon" in c for c in cols) else "A"
    st.caption(f"🔍 Tipo rilevato: {'A – DrVeto' if ftype == 'A' else 'B – VetsGo'}")

    # Classificazione automatica + individuazione colonne numeriche con coercizione
    if ftype == "A":
        col_desc = next(c for c in df.columns if "descrizione" in c.lower())
        col_fam = next(c for c in df.columns if "famiglia" in c.lower())
        # Netto (dopo sconto)
        col_netto = next(c for c in df.columns if "netto" in c.lower() and "dopo" in c.lower())
        # Quantità: priorità a "Quantità" o "Quantità / %" altrimenti usa "%" (come nello script desktop)
        qty_candidates = [c for c in df.columns if norm(c) in {"quantita", "quantita/%", "quantita/%", "quantità", "quantità/%"}]
        if qty_candidates:
            col_perc = qty_candidates[0]
        else:
            # fallback storico: la colonna "%" nel DrVeto è quella sommata come Qtà
            col_perc = next(c for c in df.columns if c.strip() == "%")

        # coerce numerico
        df[col_netto] = coerce_numeric(df[col_netto])
        df[col_perc]  = coerce_numeric(df[col_perc])

        df["FamigliaCategoria"] = df.apply(lambda r: classify_A(r[col_desc], r[col_fam], mem | updates), axis=1)
        base_col, cat_col = col_desc, "FamigliaCategoria"
    else:
        col_prest = next(c for c in df.columns if "prestazioneprodotto" in c.replace(" ", "").lower())
        col_cat = next(c for c in df.columns if "categoria" in c.lower())
        col_imp = next(c for c in df.columns if "totaleimpon" in c.lower())
        col_iva = next(c for c in df.columns if "totaleconiva" in c.replace(" ", "").lower())
        col_tot_list = [c for c in df.columns if c.lower().strip() == "totale"]
        col_tot = col_tot_list[0] if col_tot_list else next(c for c in df.columns if "totale" in c.lower())

        # coerce numerico
        df[col_imp] = coerce_numeric(df[col_imp])
        df[col_iva] = coerce_numeric(df[col_iva])
        df[col_tot] = coerce_numeric(df[col_tot])

        df["Categoria"] = df.apply(lambda r: classify_B(r[col_prest], r[col_cat], mem | updates), axis=1)
        base_col, cat_col = col_prest, "Categoria"

    # Costruisci lista termini da apprendere
    all_terms = sorted({str(v).strip() for v in df[base_col].dropna().unique()}, key=lambda s: s.casefold())
    pending = [t for t in all_terms if not any(norm(k) in norm(t) for k in (mem | updates).keys())]

    # === BLOCCO APPRENDIMENTO ===
    if pending and st.session_state.idx < len(pending):
        term = pending[st.session_state.idx]
        total = len(pending)
        progress = (st.session_state.idx + 1) / total
        st.info(f"🧠 Da classificare: {st.session_state.idx + 1} di {total} termini ({progress:.0%} completato)")

        if "last_category" not in st.session_state:
            st.session_state.last_category = list(RULES_A.keys())[0] if ftype == "A" else list(RULES_B.keys())[0]

        opts = list(RULES_A.keys()) if ftype == "A" else list(RULES_B.keys())
        cat = st.selectbox(
            f"Categoria per “{term}”:",
            opts,
            index=opts.index(st.session_state.last_category) if st.session_state.last_category in opts else 0,
            key=f"cat_{term}"
        )

        c1, c2 = st.columns([1, 1])
        with c1:
            if st.button("✅ Salva e prossimo", key=f"save_{term}"):
                updates[term] = cat
                st.session_state.local_updates = updates
                st.session_state.last_category = cat
                if st.session_state.idx + 1 < len(pending):
                    st.session_state.idx += 1
                    st.rerun()
                else:
                    # 🧠 Tutto finito → salvataggio automatico su GitHub (SINCRONO per affidabilità)
                    mem.update(updates)
                    ok = github_save_json_sync(mem)
                    st.session_state.user_memory = mem
                    st.session_state.local_updates = {}
                    st.session_state.idx = 0
                    if ok:
                        st.success("🎉 Tutti classificati e salvati automaticamente sul cloud!")
                    else:
                        st.warning("⚠️ Tutti classificati, ma NON sono riuscito a salvare sul cloud. Contatta l'amministratore.")
                    st.rerun()
        with c2:
            if st.button("💾 Salva tutto sul cloud", key=f"save_all_{term}"):
                mem.update(updates)
                ok = github_save_json_sync(mem)
                st.session_state.user_memory = mem
                st.session_state.local_updates = {}
                st.session_state.idx = 0
                if ok:
                    st.success("✅ Dizionario aggiornato sul cloud.")
                else:
                    st.warning("⚠️ Salvataggio sul cloud non riuscito. Contatta l'amministratore.")
                st.rerun()

        st.progress(progress)
        st.stop()

    # === REPORT ===
    st.success("✅ Tutti classificati. Genero Studio ISA…")

    if ftype == "A":
        # usa le stesse variabili scelte sopra
        studio = df.groupby(cat_col, dropna=False).agg({col_perc: "sum", col_netto: "sum"}).reset_index()
        studio.columns = ["FamigliaCategoria", "Qtà", "Netto"]

        # percentuali robuste come desktop
        studio["% Qtà"] = round_pct_series(studio["Qtà"])
        studio["% Netto"] = round_pct_series(studio["Netto"])

        # Totali
        tot_row = pd.DataFrame([{
            "FamigliaCategoria": "Totale",
            "Qtà": float(studio["Qtà"].sum()),
            "Netto": float(studio["Netto"].sum()),
            "% Qtà": 100.00,
            "% Netto": 100.00
        }])
        studio = pd.concat([studio, tot_row], ignore_index=True)

        # grafico
        xlab = "FamigliaCategoria"; ylab = "Netto"; title = "Somma Netto per FamigliaCategoria"

    else:
        studio = df.groupby(cat_col, dropna=False).agg({col_imp: "sum", col_iva: "sum", col_tot: "sum"}).reset_index()
        studio.columns = ["Categoria", "TotaleImponibile", "TotaleConIVA", "Totale"]

        # % Totale robusta
        studio["% Totale"] = round_pct_series(studio["Totale"])

        tot_row = pd.DataFrame([{
            "Categoria": "Totale",
            "TotaleImponibile": float(studio["TotaleImponibile"].sum()),
            "TotaleConIVA": float(studio["TotaleConIVA"].sum()),
            "Totale": float(studio["Totale"].sum()),
            "% Totale": 100.00
        }])
        studio = pd.concat([studio, tot_row], ignore_index=True)

        xlab = "Categoria"; ylab = "Totale"; title = "Somma Totale per Categoria"

    # === GRAFICO ===
    st.dataframe(studio)
    fig, ax = plt.subplots(figsize=(8, 5))
    ax.bar(studio[xlab].astype(str), studio[ylab])
    ax.set_title(title)
    plt.xticks(rotation=45, ha="right")
    buf = BytesIO()
    plt.tight_layout()
    plt.savefig(buf, format="png")
    buf.seek(0)
    st.image(buf)

    # === EXCEL ===
    wb = Workbook()
    ws = wb.active
    ws.title = "Report"
    start_row, start_col = 3, 2
    total_fill = PatternFill(start_color="FFF4B084", end_color="FFF4B084", fill_type="solid")

    for j, h in enumerate(studio.columns, start=start_col):
        ws.cell(row=start_row, column=j, value=h).font = Font(bold=True)
    for i, row in enumerate(dataframe_to_rows(studio, index=False, header=False), start=start_row + 1):
        for j, v in enumerate(row, start=start_col):
            ws.cell(row=i, column=j, value=v)
    tot_row_idx = start_row + len(studio)
    for j in range(start_col, start_col + len(studio.columns)):
        c = ws.cell(row=tot_row_idx, column=j)
        c.font = Font(bold=True)
        c.fill = total_fill

    img = XLImage(buf)
    img.anchor = f"A{tot_row_idx + 3}"
    ws.add_image(img)

    out = BytesIO()
    wb.save(out)
    st.download_button("⬇️ Scarica Excel", data=out.getvalue(), file_name=f"StudioISA_{ftype}_{datetime.now().year}.xlsx")


if __name__ == "__main__":
    main()
