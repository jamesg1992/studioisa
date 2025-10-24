import streamlit as st
import pandas as pd
import json, os, base64, requests, re
from io import BytesIO
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill
import matplotlib.pyplot as plt

# === CONFIG ===
st.set_page_config(page_title="Studio ISA - DrVeto + VetsGo", layout="wide")
GITHUB_FILE = os.getenv("GITHUB_FILE", "keywords_memory.json")
GITHUB_REPO = os.getenv("GITHUB_REPO")
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")

# === CACHE EXCEL ===
@st.cache_data(show_spinner=False, ttl=600)
def load_excel(f):
    return pd.read_excel(f)

# === UTILS ===
def norm(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip().lower())

def any_kw_in(t, keys):
    return any(k in t for k in keys)

def coerce_numeric(series: pd.Series) -> pd.Series:
    s = (series.astype(str)
         .str.replace(r"[€\s]", "", regex=True)
         .str.replace(".", "", regex=False)
         .str.replace(",", ".", regex=False))
    return pd.to_numeric(s, errors="coerce").fillna(0)

def round_pct_series(values: pd.Series) -> pd.Series:
    if values.sum() == 0:
        return values
    total = Decimal(str(values.sum()))
    raw = [Decimal(str(v)) * Decimal("100") / total for v in values]
    rounded = [x.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP) for x in raw]
    diff = Decimal("100.00") - sum(rounded)
    rounded[-1] = (rounded[-1] + diff).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return pd.Series([float(x) for x in rounded], index=values.index)

def drop_columns_robust(df: pd.DataFrame, names) -> pd.DataFrame:
    remove = {norm(n) for n in names}
    keep = [c for c in df.columns if norm(c) not in remove]
    return df[keep]

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
    # usa fam_val SOLO se proviene dalla famiglia prestazione/categoria (non famiglia cliente)
    if fam_val:
        fv = str(fam_val).strip()
        if fv and fv.lower() not in {"privato","professionista"}:
            return fv
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

# === GITHUB ===
def github_load_json():
    try:
        if not (GITHUB_REPO and GITHUB_FILE):
            return {}
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_FILE}"
        headers = {"Authorization": f"token {GITHUB_TOKEN}"} if GITHUB_TOKEN else {}
        r = requests.get(url, headers=headers, timeout=12)
        if r.status_code == 200 and "content" in r.json():
            return json.loads(base64.b64decode(r.json()["content"]).decode("utf-8"))
    except:
        pass
    return {}

def github_save_json_sync(data: dict) -> bool:
    try:
        if not (GITHUB_REPO and GITHUB_FILE and GITHUB_TOKEN):
            return False
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_FILE}"
        headers = {"Authorization": f"token {GITHUB_TOKEN}"}
        get = requests.get(url, headers=headers, timeout=12)
        sha = get.json().get("sha") if get.status_code == 200 else None
        encoded = base64.b64encode(json.dumps(data, ensure_ascii=False, indent=2).encode()).decode()
        payload = {"message": "Studio ISA update", "content": encoded, "branch": "main"}
        if sha: payload["sha"] = sha
        put = requests.put(url, headers=headers, data=json.dumps(payload), timeout=20)
        return put.status_code in (200, 201)
    except:
        return False

# === MAIN ===
def main():
    st.title("📊 Studio ISA – DrVeto + VetsGo (fix categorie & conteggi)")

    up = st.file_uploader("📁 Carica file Excel", type=["xlsx","xls"])
    if not up:
        st.info("Carica un file per iniziare.")
        return

    if "df" not in st.session_state or up.name != st.session_state.get("last_file"):
        st.session_state.df = load_excel(up)
        st.session_state.mem = github_load_json()
        st.session_state.upd = {}
        st.session_state.idx = 0
        st.session_state.last_file = up.name

    # Rimuovi eventuali colonne inutili
    df = drop_columns_robust(st.session_state.df.copy(), ["Privato","PROFESSIONISTA","Professionista"])
    mem = st.session_state.mem
    upd = st.session_state.upd

    # Rilevazione tipo file
    low_cols = [c.lower().strip() for c in df.columns]
    is_type_b = any("prestazioneprodotto" in c.replace(" ", "") for c in low_cols) and any("totaleimpon" in c for c in low_cols)
    ftype = "B" if is_type_b else "A"
    st.caption(f"🔍 Tipo rilevato: {'B – VetsGo' if ftype=='B' else 'A – DrVeto'}")

    if ftype == "A":
        # --- TROVA colonne giuste (FIX: preferisci *famiglia prestazione/categoria prodotto* e NON 'famiglia cliente')
        col_desc = next(c for c in df.columns if "descrizione" in c.lower())

        # famiglie candidate
        fam_candidates = [c for c in df.columns if "famiglia" in c.lower()]
        # priorità alle colonne che parlano di prestazione/categoria prodotto o che contengono "/" (quella classica)
        fam_pref = [c for c in fam_candidates if ("prestazione" in c.lower()) or ("categoria" in c.lower()) or ("/" in c)]
        if fam_pref:
            col_fam = fam_pref[0]
        else:
            # se non trova, non usare famiglia (evitiamo Famiglia cliente)
            col_fam = None

        col_netto = next(c for c in df.columns if "netto" in c.lower() and "dopo" in c.lower())

        # quantità: preferisci Quantità / Quantità / % al posto di %
        q_candidates = [c for c in df.columns if norm(c) in {"quantita","quantità","quantita/%","quantità/%"}]
        col_qta = q_candidates[0] if q_candidates else next(c for c in df.columns if c.strip() == "%")

        # numerici
        df[col_netto] = coerce_numeric(df[col_netto])
        df[col_qta]   = coerce_numeric(df[col_qta])

        # classificazione
        df["CategoriaFinale"] = df.apply(lambda r: classify_A(r[col_desc], r[col_fam] if col_fam else None, mem | upd), axis=1)

        # escludi eventuali valori indesiderati
        df = df[~df["CategoriaFinale"].str.lower().isin(["privato","professionista"])]

        base_col = col_desc
        group_vals = {col_qta: "sum", col_netto: "sum"}
        out_columns = ["Categoria", "Qtà", "Netto"]
        chart_y = "Netto"
        chart_title = "Somma Netto per Categoria"

    else:
        # Tipo B (VetsGo)
        col_prest = next(c for c in df.columns if "prestazioneprodotto" in c.replace(" ", "").lower())
        col_cat   = next(c for c in df.columns if "categoria" in c.lower())
        col_imp   = next(c for c in df.columns if "totaleimpon" in c.lower())
        col_iva   = next(c for c in df.columns if "totaleconiva" in c.replace(" ", "").lower())
        # 'Totale' può chiamarsi proprio 'Totale' oppure con altri suffissi: cerchiamo prima quello esatto, poi un generico
        tot_exact = [c for c in df.columns if c.lower().strip() == "totale"]
        col_tot = tot_exact[0] if tot_exact else next(c for c in df.columns if "totale" in c.lower())

        # numerici
        df[col_imp] = coerce_numeric(df[col_imp])
        df[col_iva] = coerce_numeric(df[col_iva])
        df[col_tot] = coerce_numeric(df[col_tot])

        # classifica
        df["CategoriaFinale"] = df.apply(lambda r: classify_B(r[col_prest], r[col_cat], mem | upd), axis=1)

        # sicurezza: rimuovi eventuali 'privato/professionista' se mai capitassero
        df = df[~df["CategoriaFinale"].str.lower().isin(["privato","professionista"])]

        base_col = col_prest
        group_vals = {col_imp: "sum", col_iva: "sum", col_tot: "sum"}
        out_columns = ["Categoria", "TotaleImponibile", "TotaleConIVA", "Totale"]
        chart_y = "Totale"
        chart_title = "Somma Totale per Categoria"

    # === APPRENDIMENTO ===
    all_terms = sorted({str(v).strip() for v in df[base_col].dropna().unique()}, key=str.casefold)
    pending = [t for t in all_terms if not any(norm(k) in norm(t) for k in (mem | upd).keys())]

    if pending and st.session_state.idx < len(pending):
        term = pending[st.session_state.idx]
        st.info(f"🧠 Da classificare: {st.session_state.idx+1}/{len(pending)} → {term}")

        opts = list(RULES_A.keys()) if ftype == "A" else list(RULES_B.keys())
        if "last_cat" not in st.session_state:
            st.session_state.last_cat = opts[0]
        cat = st.selectbox("Categoria:", opts, index=opts.index(st.session_state.last_cat))

        if st.button("✅ Salva & prossimo"):
            upd[term] = cat
            st.session_state.last_cat = cat
            st.session_state.idx += 1
            if st.session_state.idx >= len(pending):
                mem.update(upd)
                ok = github_save_json_sync(mem)
                if ok:
                    st.success("🎉 Tutto classificato e salvato su GitHub!")
                else:
                    st.warning("⚠️ Classificato, ma non sono riuscito a salvare su GitHub.")
            st.rerun()
        st.stop()

    # === REPORT ===
    st.success("✅ Tutti classificati. Genero Studio ISA…")

    studio = df.groupby("CategoriaFinale", dropna=False).agg(group_vals).reset_index()
    studio = studio.rename(columns={"CategoriaFinale": "Categoria"})
    # percentuali
    if ftype == "A":
        studio["% Qtà"] = round_pct_series(studio[out_columns[1]])
        studio["% Netto"] = round_pct_series(studio[out_columns[2]])
        # totale
        studio.loc[len(studio)] = [
            "Totale",
            float(studio[out_columns[1]].sum()),
            float(studio[out_columns[2]].sum()),
            100.00,
            100.00
        ]
    else:
        studio["% Totale"] = round_pct_series(studio[out_columns[-1]])
        studio.loc[len(studio)] = [
            "Totale",
            float(studio[out_columns[1]].sum()),
            float(studio[out_columns[2]].sum()),
            float(studio[out_columns[3]].sum()),
            100.00
        ]

    st.dataframe(studio)

    # grafico
    fig, ax = plt.subplots(figsize=(8, 5))
    ax.bar(studio["Categoria"].astype(str), studio[chart_y])
    ax.set_title(chart_title)
    plt.xticks(rotation=45, ha="right")
    buf = BytesIO(); plt.tight_layout(); plt.savefig(buf, format="png"); buf.seek(0)
    st.image(buf)

    # excel
    wb = Workbook(); ws = wb.active; ws.title = "Report"
    start_row, start_col = 3, 2
    total_fill = PatternFill(start_color="FFF4B084", end_color="FFF4B084", fill_type="solid")

    for j, h in enumerate(studio.columns, start=start_col):
        ws.cell(row=start_row, column=j, value=h).font = Font(bold=True)
    for i, row in enumerate(dataframe_to_rows(studio, index=False, header=False), start=start_row + 1):
        for j, v in enumerate(row, start=start_col):
            ws.cell(row=i, column=j, value=v)
    last = start_row + len(studio)
    for j in range(start_col, start_col + len(studio.columns)):
        c = ws.cell(row=last, column=j); c.font = Font(bold=True); c.fill = total_fill

    ws.add_image(XLImage(buf), f"A{last+3}")
    out = BytesIO(); wb.save(out)
    st.download_button("⬇️ Scarica Excel", data=out.getvalue(), file_name=f"StudioISA_{ftype}_{datetime.now().year}.xlsx")

if __name__ == "__main__":
    main()
