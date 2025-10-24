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

# === UTILITIES ===
def norm(s): return re.sub(r"\s+", " ", str(s).strip().lower())

def any_kw_in(t, keys): return any(k in t for k in keys)

def coerce_numeric(s):
    s = (s.astype(str)
         .str.replace(r"[€\s]", "", regex=True)
         .str.replace(".", "", regex=False)
         .str.replace(",", ".", regex=False))
    return pd.to_numeric(s, errors="coerce").fillna(0)

def round_pct(values):
    values = pd.to_numeric(values, errors="coerce").fillna(0)
    total = Decimal(str(values.sum()))
    if total == 0: return values * 0
    raw = [Decimal(str(v)) * Decimal("100") / total for v in values]
    rounded = [r.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP) for r in raw]
    diff = Decimal("100.00") - sum(rounded)
    rounded[-1] = (rounded[-1] + diff).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return pd.Series([float(x) for x in rounded], index=values.index)

def load_excel(file):
    return pd.read_excel(file)

# === GITHUB ===
def github_load_json():
    if not (GITHUB_REPO and GITHUB_FILE):
        return {}
    try:
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_FILE}"
        headers = {"Authorization": f"token {GITHUB_TOKEN}"} if GITHUB_TOKEN else {}
        r = requests.get(url, headers=headers)
        if r.status_code == 200:
            content = base64.b64decode(r.json()["content"]).decode("utf-8")
            return json.loads(content)
    except:
        pass
    return {}

def github_save_json(data):
    if not (GITHUB_REPO and GITHUB_FILE and GITHUB_TOKEN):
        return
    url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_FILE}"
    headers = {"Authorization": f"token {GITHUB_TOKEN}"}
    get = requests.get(url, headers=headers)
    sha = get.json().get("sha") if get.status_code == 200 else None
    encoded = base64.b64encode(json.dumps(data, ensure_ascii=False, indent=2).encode()).decode()
    payload = {"message": "Update keywords", "content": encoded, "branch": "main"}
    if sha: payload["sha"] = sha
    requests.put(url, headers=headers, data=json.dumps(payload))

# === REGOLE ===
RULES_A = {
    "LABORATORIO": ["analisi","emocromo","test","esame","coprolog","feci","giardia","leishmania","citolog","istolog","urinocolt","urine"],
    "VISITE": ["visita","controllo","consulto","dermatologic"],
    "FAR": ["meloxidyl","konclav","enrox","profenacarp","apoquel","osurnia","cylan","mometa","aristos","cytopoint","milbemax","stomorgyl","previcox"],
    "CHIRURGIA": ["intervento","chirurg","castraz","sterilizz","ovariect","detartrasi","estraz","biopsia","orchiettomia","odontostomat"],
    "DIAGNOSTICA PER IMMAGINI": ["rx","radiograf","eco","ecografia","tac"],
    "MEDICINA": ["terapia","terapie","flebo","day hospital","trattamento","emedog","cerenia","endovena","pressione"],
    "VACCINI": ["vacc","letifend","rabbia","trivalente","felv"],
    "CHIP": ["microchip","chip"],
    "ALTRE PRESTAZIONI": ["trasporto","eutanasia","unghie","cremazion","otoematoma","pet corner","ricette","medicazione","manualità"]
}

RULES_B = {
    "Visite domiciliari o presso allevamenti": ["visite domiciliari","allevamenti","domicilio"],
    "Visite ambulatoriali": ["terapia","trattamenti","vaccinazioni","ambulatorio","manualità","pet corner","visite","ricette","medicazione","microchip","controllo"],
    "Esami diagnostici per immagine": ["radiologia","eco","ecografia","tac","rx","raggi"],
    "Altri esami diagnostici": ["laboratorio","emocromo","prelievo"],
    "Interventi chirurgici": ["chirurg","castraz","ovariect","detartrasi","estraz","eutanasia","anestesia"],
    "Altre attività": ["acconto"]
}

# === CLASSIFICAZIONE ===
def classify_A(desc, fam, mem):
    if fam and fam.lower() not in {"privato","professionista"}:
        return fam.strip()
    d = norm(desc)
    for k,v in mem.items():
        if norm(k) in d: return v
    for cat,keys in RULES_A.items():
        if any_kw_in(d, keys): return cat
    return "ALTRE PRESTAZIONI"

def classify_B(prest, cat, mem):
    if cat and cat.strip(): return cat.strip()
    d = norm(prest)
    for k,v in mem.items():
        if norm(k) in d: return v
    for cat,keys in RULES_B.items():
        if any_kw_in(d, keys): return cat
    return "Altre attività"

# === MAIN ===
def main():
    st.title("📊 Studio ISA – DrVeto + VetsGo")

    file = st.file_uploader("Seleziona Excel", type=["xlsx","xls"])
    if not file: 
        st.stop()

    if "df" not in st.session_state:
        st.session_state.df = load_excel(file)
        st.session_state.mem = github_load_json()
        st.session_state.new = {}
        st.session_state.idx = 0

    df = st.session_state.df.copy()
    mem = st.session_state.mem
    new = st.session_state.new

    # Identifica tipo file
    cols = [c.lower() for c in df.columns]
    typeB = any("prestazioneprodotto" in c for c in cols)
    mode = "B" if typeB else "A"

    # === TIPO A ===
    if mode == "A":
        desc = next(c for c in df.columns if "descrizione" in c.lower())
        fam = next(c for c in df.columns if "prestazione" in c.lower() or "/" in c, None)
        q_candidates = [c for c in df.columns if "quant" in c.lower()]
        qta = q_candidates[0] if q_candidates else next(c for c in df.columns if c.strip()=="%")
        netto = next(c for c in df.columns if "netto" in c.lower() and "dopo" in c.lower())

        df[qta] = coerce_numeric(df[qta])
        df[netto] = coerce_numeric(df[netto])
        df["CategoriaFinale"] = df.apply(lambda r: classify_A(r[desc], r[fam] if fam else None, mem|new), axis=1)
        df = df[~df["CategoriaFinale"].str.lower().isin(["privato","professionista"])]

        base = desc

    else:
        prest = next(c for c in df.columns if "prestazioneprodotto" in c.replace(" ","").lower())
        cat = next(c for c in df.columns if "categoria" in c.lower())
        imp = next(c for c in df.columns if "totaleimpon" in c.lower())
        iva = next(c for c in df.columns if "totaleconiva" in c.replace(" ","").lower())
        tot = next(c for c in df.columns if c.lower()=="totale" or "totale" in c.lower())

        df[imp] = coerce_numeric(df[imp])
        df[iva] = coerce_numeric(df[iva])
        df[tot] = coerce_numeric(df[tot])
        df["CategoriaFinale"] = df.apply(lambda r: classify_B(r[prest], r[cat], mem|new), axis=1)
        df = df[~df["CategoriaFinale"].str.lower().isin(["privato","professionista"])]
        base = prest

    # === APPRENDIMENTO ===
    terms = sorted({str(v).strip() for v in df[base].dropna().unique()}, key=str.casefold)
    pending = [t for t in terms if not any(norm(k) in norm(t) for k in (mem|new).keys())]

    if pending:
        term = pending[st.session_state.idx]
        st.warning(f"🧠 Da classificare: {st.session_state.idx+1}/{len(pending)} → “{term}”")

        opts = list(RULES_A.keys()) if mode=="A" else list(RULES_B.keys())
        cat = st.selectbox("Categoria:", opts)

        if st.button("✅ Salva e prossimo"):
            new[term] = cat
            st.session_state.new = new
            st.session_state.idx += 1

            if st.session_state.idx >= len(pending):
                mem.update(new)
                github_save_json(mem)
                st.success("🎉 Tutti salvati su GitHub!")
            st.rerun()

        st.stop()

    # === REPORT ===
    if mode=="A":
        studio = df.groupby("CategoriaFinale", dropna=False).agg({qta:"sum", netto:"sum"}).reset_index()
        studio = studio.rename(columns={"CategoriaFinale":"Categoria", qta:"Qtà", netto:"Netto"})
        studio["% Qtà"] = round_pct(studio["Qtà"])
        studio["% Netto"] = round_pct(studio["Netto"])
        studio.loc[len(studio)] = ["Totale", studio["Qtà"].sum(), studio["Netto"].sum(), 100, 100]

    else:
        studio = df.groupby("CategoriaFinale", dropna=False).agg({imp:"sum", iva:"sum", tot:"sum"}).reset_index()
        studio = studio.rename(columns={"CategoriaFinale":"Categoria", imp:"TotaleImponibile", iva:"TotaleConIVA", tot:"Totale"})
        studio["% Totale"] = round_pct(studio["Totale"])
        studio.loc[len(studio)] = ["Totale", studio["TotaleImponibile"].sum(), studio["TotaleConIVA"].sum(), studio["Totale"].sum(), 100]

    st.dataframe(studio)

    # Grafico
    fig, ax = plt.subplots(figsize=(8,5))
    ycol = "Netto" if mode=="A" else "Totale"
    ax.bar(studio["Categoria"], studio[ycol])
    plt.xticks(rotation=45, ha="right")
    buf = BytesIO(); plt.tight_layout(); plt.savefig(buf, format="png"); buf.seek(0)
    st.image(buf)

    # Excel
    wb = Workbook(); ws = wb.active; ws.title = "Report"
    for r in dataframe_to_rows(studio, index=False, header=True):
        ws.append(r)
    ws.add_image(XLImage(buf), f"A{len(studio)+4}")
    out = BytesIO(); wb.save(out)

    st.download_button("⬇️ Scarica Excel", out.getvalue(), file_name=f"StudioISA_{datetime.now().year}.xlsx")

if __name__ == "__main__":
    main()
