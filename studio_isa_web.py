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


st.set_page_config(page_title="Studio ISA - Alcyon Italia", layout="wide")

GITHUB_FILE = os.getenv("GITHUB_FILE", "keywords_memory.json")
GITHUB_REPO = os.getenv("GITHUB_REPO")
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")


@st.cache_data(show_spinner=False, ttl=600)
def load_excel(f):
    return pd.read_excel(f)


def norm(s):  # normalizza testo
    return re.sub(r"\s+", " ", str(s).strip().lower())


def any_kw_in(t, keys):
    return any(k in t for k in keys)


def coerce_numeric(series):
    s = (series.astype(str)
         .str.replace(r"[€\s]", "", regex=True)
         .str.replace(".", "", regex=False)
         .str.replace(",", ".", regex=False))
    return pd.to_numeric(s, errors="coerce").fillna(0)


def round_pct_series(values):
    if values.sum() == 0:
        return values
    total = Decimal(str(values.sum()))
    raw = [Decimal(str(v)) * Decimal("100") / total for v in values]
    rounded = [x.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP) for x in raw]
    diff = Decimal("100.00") - sum(rounded)
    rounded[-1] = (rounded[-1] + diff).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return pd.Series([float(x) for x in rounded], index=values.index)


def drop_columns(df):
    remove = {"privato", "professionista"}
    return df[[c for c in df.columns if norm(c) not in remove]]


# === REGOLE TIPO A (DrVeto) ===
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

# === REGOLE TIPO B (VetsGo) ===
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


def classify_A(desc, fam, mem):
    if pd.notna(fam) and str(fam).strip():
        return fam
    d = norm(desc)
    for k,v in mem.items():
        if norm(k) in d: return v
    for cat, keys in RULES_A.items():
        if any_kw_in(d, keys): return cat
    return "ALTRE PRESTAZIONI"


def classify_B(prest, cat, mem):
    if pd.notna(cat) and str(cat).strip():
        return cat
    d = norm(prest)
    for k,v in mem.items():
        if norm(k) in d: return v
    for cat, keys in RULES_B.items():
        if any_kw_in(d, keys): return cat
    return "Altre attività"


def github_load():
    try:
        if not (GITHUB_REPO and GITHUB_TOKEN):
            return {}
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_FILE}"
        headers = {"Authorization": f"token {GITHUB_TOKEN}"}
        r = requests.get(url, headers=headers, timeout=12)
        if r.status_code == 200:
            return json.loads(base64.b64decode(r.json()["content"]).decode())
    except:
        pass
    return {}


def github_save(data):
    try:
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_FILE}"
        headers = {"Authorization": f"token {GITHUB_TOKEN}"}
        get = requests.get(url, headers=headers).json()
        sha = get.get("sha") if "sha" in get else None
        encoded = base64.b64encode(json.dumps(data, ensure_ascii=False, indent=2).encode()).decode()
        payload = {"message": "Studio ISA update", "content": encoded, "branch": "main"}
        if sha: payload["sha"] = sha
        put = requests.put(url, headers=headers, data=json.dumps(payload))
        return put.status_code in (200,201)
    except:
        return False


def main():
    st.title("📊 Studio ISA - DrVeto + VetsGo)")

    up = st.file_uploader("📁 Carica file", type=["xlsx","xls"])
    if not up:
        st.stop()

    if "df" not in st.session_state:
        st.session_state.df = load_excel(up)
        st.session_state.mem = github_load()
        st.session_state.upd = {}
        st.session_state.idx = 0

    df = drop_columns(st.session_state.df.copy())
    mem = st.session_state.mem
    upd = st.session_state.upd

    cols = [norm(c) for c in df.columns]
    ftype = "B" if ("prestazioneprodotto" in "".join(cols) and "totaleimpon" in "".join(cols)) else "A"
    st.caption(f"🔍 Tipo rilevato: { 'A (DrVeto)' if ftype=='A' else 'B (VetsGo)' }")

    if ftype=="A":
        col_desc = next(c for c in df.columns if "descrizione" in c.lower())
        col_fam  = next(c for c in df.columns if "famiglia" in c.lower())
        col_netto= next(c for c in df.columns if "netto" in c.lower() and "dopo" in c.lower())
        q_cols = [c for c in df.columns if norm(c) in {"quantita","quantità","quantita/%","quantità/%"}]
        col_qta = q_cols[0] if q_cols else next(c for c in df.columns if c.strip()=="%")
        df[col_netto] = coerce_numeric(df[col_netto])
        df[col_qta]   = coerce_numeric(df[col_qta])
        df["CategoriaFinale"] = df.apply(lambda r: classify_A(r[col_desc], r[col_fam], mem|upd), axis=1)
        base = col_desc

    else:
        col_prest = next(c for c in df.columns if "prestazioneprodotto" in c.replace(" ","").lower())
        col_cat   = next(c for c in df.columns if "categoria" in c.lower())
        col_imp   = next(c for c in df.columns if "totaleimpon" in c.lower())
        col_iva   = next(c for c in df.columns if "totaleconiva" in c.lower().replace(" ",""))
        col_tot   = next(c for c in df.columns if c.lower().strip()=="totale")
        df[col_imp] = coerce_numeric(df[col_imp])
        df[col_iva] = coerce_numeric(df[col_iva])
        df[col_tot] = coerce_numeric(df[col_tot])
        df["CategoriaFinale"] = df.apply(lambda r: classify_B(r[col_prest], r[col_cat], mem|upd), axis=1)
        base = col_prest

    all_terms = sorted({str(v).strip() for v in df[base].dropna().unique()}, key=str.casefold)
    pending = [t for t in all_terms if not any(norm(k) in norm(t) for k in (mem|upd).keys())]

    if pending and st.session_state.idx < len(pending):
        term = pending[st.session_state.idx]
        st.warning(f"🧠 {st.session_state.idx+1}/{len(pending)} → {term}")

        opts = list(RULES_A.keys()) if ftype=="A" else list(RULES_B.keys())
        if "last_cat" not in st.session_state: st.session_state.last_cat = opts[0]
        cat = st.selectbox("Categoria:", opts, index=opts.index(st.session_state.last_cat))

        if st.button("✅ Salva & prossimo"):
            upd[term] = cat
            st.session_state.last_cat = cat
            st.session_state.idx += 1
            if st.session_state.idx >= len(pending):
                mem.update(upd)
                github_save(mem)
                st.success("🎉 Tutto classificato & salvato sul cloud!")
            st.rerun()
        st.stop()

    st.success("✅ Tutti i termini classificati — Genero Studio ISA…")

    if ftype=="A":
        studio = df.groupby("CategoriaFinale").agg({col_qta:"sum", col_netto:"sum"}).reset_index()
        studio.columns = ["Categoria","Qtà","Netto"]
        studio["% Qtà"] = round_pct_series(studio["Qtà"])
        studio["% Netto"] = round_pct_series(studio["Netto"])
        studio.loc[len(studio)] = ["Totale", studio["Qtà"].sum(), studio["Netto"].sum(), 100, 100]

        ylabel="Netto"; title="Somma Netto per Categoria"

    else:
        studio = df.groupby("CategoriaFinale").agg({col_imp:"sum", col_iva:"sum", col_tot:"sum"}).reset_index()
        studio.columns=["Categoria","TotaleImponibile","TotaleConIVA","Totale"]
        studio["% Totale"] = round_pct_series(studio["Totale"])
        studio.loc[len(studio)] = ["Totale", studio["TotaleImponibile"].sum(), studio["TotaleConIVA"].sum(), studio["Totale"].sum(), 100]

        ylabel="Totale"; title="Somma Totale per Categoria"

    st.dataframe(studio)

    fig, ax = plt.subplots(figsize=(8,5))
    ax.bar(studio["Categoria"], studio[ylabel])
    ax.set_title(title)
    plt.xticks(rotation=45, ha="right")
    buf=BytesIO(); plt.tight_layout(); plt.savefig(buf, format="png"); buf.seek(0)
    st.image(buf)

    wb = Workbook(); ws = wb.active; ws.title="Report"
    start=3; fill=PatternFill(start_color="FFF4B084",end_color="FFF4B084",fill_type="solid")
    for j,h in enumerate(studio.columns,start=2): ws.cell(row=start,column=j,value=h).font=Font(bold=True)
    for i,row in enumerate(dataframe_to_rows(studio,index=False,header=False),start=start+1):
        for j,v in enumerate(row,start=2): ws.cell(row=i,column=j,value=v)
    last=start+len(studio)
    for j in range(2,2+len(studio.columns)): ws.cell(row=last,column=j).font=Font(bold=True); ws.cell(row=last,column=j).fill=fill
    ws.add_image(XLImage(buf),f"A{last+3}")

    out=BytesIO(); wb.save(out)
    st.download_button("⬇️ Scarica Excel", data=out.getvalue(), file_name=f"StudioISA_{datetime.now().year}.xlsx")


if __name__ == "__main__":
    main()
