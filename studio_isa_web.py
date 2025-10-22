import streamlit as st
import pandas as pd
import json, os, base64, requests, re
from io import BytesIO
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill
import matplotlib.pyplot as plt

# ───────────── CONFIG ─────────────
st.set_page_config(page_title="Studio ISA - Alcyon Italia", layout="wide")
GITHUB_FILE = os.getenv("GITHUB_FILE", "keywords_memory.json")
GITHUB_REPO = os.getenv("GITHUB_REPO")
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")

# ───────────── RULES (Tipo A e B) ─────────────
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
    "Domicilio": ["visite domiciliari","allevamenti","domicilio"],
    "Terapia": ["visite ambulatoriali","terapia","trattamenti","vaccinazioni","ambulatorio","manualità","pet corner","visite","ricette","medicazione","microchip","controllo"],
    "Radiologia": ["esami diagnostici per immagine","radiologia","eco","ecografia","tac","rx","raggi"],
    "Laboratorio": ["altri esami diagnostici","esami biochimici","laboratorio","malattie infettive","emocromo","prelievo"],
    "Chirurgia": ["interventi chirurgici","avulsione","endoscopia","eutanasia","sedazione","anestesia","chirurgia","odontostomat","orchiettomia","asportazione","biopsia","ovariectomia"],
    "Ostetricia": ["assistenza al parto","ostetricia","parto"],
    "Consulenza": ["attività di consulenza","perizia","collaborazione","telemedicina","consulto"],
    "Inseminazione": ["inseminazione artificiale"],
    "Altre attività": ["acconto"]
}

# ───────────── GITHUB I/O ─────────────
def github_load_json():
    try:
        if not (GITHUB_REPO and GITHUB_FILE):
            return {}
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_FILE}"
        headers = {"Authorization": f"token {GITHUB_TOKEN}"} if GITHUB_TOKEN else {}
        r = requests.get(url, headers=headers)
        if r.status_code == 200 and "content" in r.json():
            return json.loads(base64.b64decode(r.json()["content"]).decode("utf-8"))
    except Exception:
        pass
    return {}

def github_save_json(data: dict):
    try:
        if not (GITHUB_REPO and GITHUB_FILE and GITHUB_TOKEN):
            st.info("ℹ️ GitHub non configurato.")
            return
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_FILE}"
        headers = {"Authorization": f"token {GITHUB_TOKEN}"}
        get_res = requests.get(url, headers=headers)
        sha = get_res.json().get("sha") if get_res.status_code == 200 else None
        encoded = base64.b64encode(json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")).decode("utf-8")
        payload = {"message": "Aggiornamento dizionario Studio ISA", "content": encoded, "branch": "main"}
        if sha: payload["sha"] = sha
        requests.put(url, headers=headers, data=json.dumps(payload))
    except Exception as e:
        st.error(f"❌ Salvataggio GitHub fallito: {e}")

# ───────────── UTILS ─────────────
def norm(s): return re.sub(r"\s+", " ", str(s).strip().lower())
def any_kw_in(t, kws): return any(k in t for k in kws)
@st.cache_data(show_spinner=False)
def load_excel(f): return pd.read_excel(f)

# ───────────── CLASSIFICAZIONE ─────────────
def classify_A(desc, fam_val, mem):
    if pd.notna(fam_val) and str(fam_val).strip(): return str(fam_val).strip()
    d = norm(desc)
    for k,v in mem.items():
        if norm(k) in d: return v
    for cat,keys in RULES_A.items():
        if any_kw_in(d, keys): return cat
    return "ALTRE PRESTAZIONI"

def classify_B(prest, cat_val, mem):
    if pd.notna(cat_val) and str(cat_val).strip(): return str(cat_val).strip()
    d = norm(prest)
    for k,v in mem.items():
        if norm(k) in d: return v
    for cat,keys in RULES_B.items():
        if any_kw_in(d, keys): return cat
    return "Altre attività"

# ───────────── MAIN ─────────────
def main():
    st.title("📊 Studio ISA - DrVeto e VetsGo")
    up = st.file_uploader("📁 Seleziona file Excel", type=["xlsx","xls"])
    if not up:
        st.info("Carica un file per iniziare.")
        return

    # Inizializza sessione
    if "df" not in st.session_state or up.name != st.session_state.get("last_file"):
        st.session_state.df = load_excel(up)
        st.session_state.user_memory = github_load_json()
        st.session_state.local_updates = {}
        st.session_state.pending_terms = []
        st.session_state.idx = 0
        st.session_state.last_file = up.name

    df = st.session_state.df.copy()
    mem = st.session_state.user_memory
    updates = st.session_state.local_updates

    # Rileva tipo file
    ftype = "A" if "descrizione" in " ".join(df.columns).lower() else "B"
    st.caption(f"🔍 Tipo rilevato: {'A – DrVeto' if ftype=='A' else 'B – Gestionale'}")

    # CLASSIFICA
    if ftype == "A":
        col_desc = next(c for c in df.columns if "descrizione" in c.lower())
        col_fam = next(c for c in df.columns if "famiglia" in c.lower())
        col_netto = next(c for c in df.columns if "netto" in c.lower())
        col_perc = next(c for c in df.columns if c.strip() == "%")
        df["FamigliaCategoria"] = df.apply(lambda r: classify_A(r[col_desc], r[col_fam], mem|updates), axis=1)
        base_col, cat_col = col_desc, "FamigliaCategoria"
    else:
        col_prest = next(c for c in df.columns if "prestazione" in c.lower())
        col_cat = next(c for c in df.columns if "categoria" in c.lower())
        col_imp = next(c for c in df.columns if "impon" in c.lower())
        col_iva = next(c for c in df.columns if "coniva" in c.lower())
        col_tot = next(c for c in df.columns if c.lower().startswith("totale"))
        df["Categoria"] = df.apply(lambda r: classify_B(r[col_prest], r[col_cat], mem|updates), axis=1)
        base_col, cat_col = col_prest, "Categoria"

    # Trova nuovi termini
    all_terms = sorted({str(v).strip() for v in df[base_col].dropna().unique()}, key=lambda s: s.casefold())
    pending = [t for t in all_terms if not any(norm(k) in norm(t) for k in (mem|updates).keys())]

    if pending and st.session_state.idx < len(pending):
        term = pending[st.session_state.idx]
        st.warning(f"🧠 Nuovo termine: {term} ({st.session_state.idx+1}/{len(pending)})")
        opts = list(RULES_A.keys()) if ftype=="A" else list(RULES_B.keys())
        cat = st.selectbox("Categoria:", opts, key=f"sel_{term}")
        if st.button("✅ Salva locale e prossimo"):
            updates[term] = cat
            st.session_state.local_updates = updates
            st.session_state.idx += 1
            st.rerun()
        st.stop()

    # Tutti classificati → report
    st.success("✅ Tutti classificati. Genero report…")

    if ftype == "A":
        col_perc = next(c for c in df.columns if c.strip() == "%")
        col_netto = next(c for c in df.columns if "netto" in c.lower())
        studio = df.groupby(cat_col, dropna=False).agg({col_perc:"sum", col_netto:"sum"}).reset_index()
        studio.columns = ["FamigliaCategoria","Qtà","Netto"]
        tot_q, tot_n = studio["Qtà"].sum(), studio["Netto"].sum()
        studio["% Qtà"] = (studio["Qtà"]/tot_q*100).round(2)
        studio["% Netto"] = (studio["Netto"]/tot_n*100).round(2)
    else:
        col_imp = next(c for c in df.columns if "impon" in c.lower())
        col_iva = next(c for c in df.columns if "coniva" in c.lower())
        col_tot = next(c for c in df.columns if c.lower().startswith("totale"))
        studio = df.groupby(cat_col, dropna=False).agg({col_imp:"sum", col_iva:"sum", col_tot:"sum"}).reset_index()
        studio.columns = ["Categoria","TotaleImponibile","TotaleConIVA","Totale"]
        tot_t = studio["Totale"].sum()
        studio["% Totale"] = (studio["Totale"]/tot_t*100).round(2)

    st.dataframe(studio)

    # Salva su GitHub (tutti insieme)
    if st.button("💾 Salva TUTTI su GitHub"):
        mem.update(updates)
        github_save_json(mem)
        st.session_state.user_memory = mem
        st.session_state.local_updates = {}
        st.session_state.idx = 0
        st.success("✅ Dizionario aggiornato su GitHub!")

if __name__ == "__main__":
    main()
