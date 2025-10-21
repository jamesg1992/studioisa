import streamlit as st
import pandas as pd
import json, os, base64, requests
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill
import matplotlib.pyplot as plt

# === CONFIG ===
st.set_page_config(page_title="Studio ISA", layout="wide")
GITHUB_FILE = os.getenv("GITHUB_FILE", "keywords_memory.json")
GITHUB_REPO = os.getenv("GITHUB_REPO")
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")

# === GITHUB I/O ===
def github_load_json():
    try:
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_FILE}"
        headers = {"Authorization": f"token {GITHUB_TOKEN}"} if GITHUB_TOKEN else {}
        r = requests.get(url, headers=headers)
        if r.status_code == 200:
            content = base64.b64decode(r.json()["content"]).decode("utf-8")
            return json.loads(content)
        return {}
    except Exception:
        return {}

def github_save_json(data: dict):
    try:
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_FILE}"
        headers = {"Authorization": f"token {GITHUB_TOKEN}"} if GITHUB_TOKEN else {}
        get_res = requests.get(url, headers=headers)
        sha = get_res.json().get("sha") if get_res.status_code == 200 else None

        encoded = base64.b64encode(json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")).decode("utf-8")
        payload = {"message": "Aggiornamento dizionario Studio ISA", "content": encoded, "branch": "main"}
        if sha:
            payload["sha"] = sha

        put_res = requests.put(url, headers=headers, data=json.dumps(payload))
        if put_res.status_code in (200, 201):
            st.success("✅ Dizionario aggiornato su GitHub!")
        else:
            st.error(f"❌ Errore GitHub: {put_res.status_code}")
    except Exception as e:
        st.error(f"❌ Salvataggio GitHub fallito: {e}")

# === CACHE LETTURA EXCEL ===
@st.cache_data(show_spinner=False)
def load_excel(file):
    return pd.read_excel(file)

# === CLASSIFICAZIONE ===
_RULES = {
    "LABORATORIO": ["analisi","emocromo","test","esame","coprolog","feci","giardia","leishmania","citolog","istolog","urinocolt"],
    "VISITE": ["visita","controllo","consulto","dermatologico"],
    "FAR": ["meloxidyl","konclav","enrox","profenacarp","apoquel","osurnia","cylanic","mometa","aristos","cytopoint","milbemax","stomorgyl","previcox"],
    "CHIRURGIA": ["intervento","chirurg","castraz","sterilizz","ovariect","detartrasi","estraz"],
    "DIAGNOSTICA PER IMMAGINI": ["rx","radiograf","eco","ecografia","tac"],
    "MEDICINA": ["terapia","terapie","flebo","day hospital","trattamento","emedog","cerenia","endovena"],
    "VACCINI": ["vacc","letifend","rabbia","trivalente","felv"],
    "CHIP": ["microchip"],
    "ALTRE PRESTAZIONI": ["trasporto","eutanasia","unghie","cremazion","otoematoma"]
}

def classify(desc, fam_val, memory: dict):
    if pd.notna(fam_val) and str(fam_val).strip():
        return fam_val
    d = str(desc).lower().strip()
    if not d:
        return "ALTRE PRESTAZIONI"
    for key, cat in memory.items():
        if key.lower() in d:
            return cat
    for cat, keys in _RULES.items():
        if any(k in d for k in keys):
            return cat
    return "ALTRE PRESTAZIONI"

# === MAIN ===
def main():
    st.title("📊 Studio ISA – Web App (Fast v5)")

    uploaded = st.file_uploader("📁 Seleziona file Excel", type=["xlsx","xls"])
    if not uploaded:
        st.info("Carica un file per iniziare.")
        return

    # INIT
    if "df" not in st.session_state:
        st.session_state.df = load_excel(uploaded)
        st.session_state.user_memory = github_load_json()
        st.session_state.local_updates = {}  # nuove categorie non ancora salvate
        st.session_state.pending_terms = []
        st.session_state.idx = 0

    df = st.session_state.df
    user_memory = st.session_state.user_memory
    local_updates = st.session_state.local_updates

    col_desc = next((c for c in df.columns if "descrizione" in c.lower()), None)
    col_fam  = next((c for c in df.columns if "famiglia" in c.lower()), None)
    col_netto= next((c for c in df.columns if "netto" in c.lower() and "dopo" in c.lower()), None)
    col_perc = next((c for c in df.columns if c.strip() == "%"), None)
    if not all([col_desc, col_fam, col_netto, col_perc]):
        st.error("❌ Colonne richieste non trovate.")
        return

    df["FamigliaCategoria"] = df.apply(lambda r: classify(r[col_desc], r[col_fam], user_memory), axis=1)

    # --- Nuovi termini ---
    if not st.session_state.pending_terms:
        uniq = sorted({str(v).strip() for v in df[col_desc].dropna().unique()}, key=lambda s: s.casefold())
        for term in uniq:
            t = term.lower()
            if not any(k.lower() in t for k in user_memory.keys()) and not any(k in t for keys in _RULES.values() for k in keys):
                st.session_state.pending_terms.append(term)

    pending = st.session_state.pending_terms
    idx = st.session_state.idx
# Se l'indice è fuori range (es. dopo l'ultimo), resetta a 0
    if idx >= len(pending):
        st.session_state.idx = 0
        idx = 0
    
    if pending:
        term = pending[idx]
        st.warning(f"🧠 Da classificare: {len(pending)} termini | Corrente: {idx+1}/{len(pending)}")
        cat = st.selectbox(f"Categoria per “{term}”:", list(_RULES.keys()), key=f"cat_{idx}")

        c1, c2, c3 = st.columns([1,1,2])
        with c1:
            if st.button("✅ Salva locale", key=f"save_{idx}"):
                local_updates[term] = cat
                st.session_state.local_updates = local_updates
                st.session_state.idx += 1
                if st.session_state.idx >= len(pending):
                    st.success("🎉 Tutti classificati! Ora puoi salvare su GitHub.")
                st.rerun()

        with c2:
            if st.button("⏭️ Salta"):
                st.session_state.idx += 1
                if st.session_state.idx >= len(pending):
                    st.session_state.idx = 0
                st.rerun()

        with c3:
            if st.button("💾 Salva tutto su GitHub", type="primary"):
                user_memory.update(local_updates)
                github_save_json(user_memory)
                st.session_state.user_memory = user_memory
                st.session_state.local_updates = {}
                st.session_state.pending_terms = []
                st.success("✅ Tutti i nuovi termini salvati su GitHub!")
                st.rerun()
        return

    # --- Calcolo report ---
    st.success("✅ Tutti i termini classificati. Genero report…")

    studio_isa = (
        df.groupby("FamigliaCategoria", dropna=False)
        .agg({col_perc: "sum", col_netto: "sum"})
        .reset_index()
        .rename(columns={col_perc: "Qtà", col_netto: "Netto"})
    )
    tot_qta = studio_isa["Qtà"].sum()
    tot_netto = studio_isa["Netto"].sum()
    studio_isa["% Qtà"] = (studio_isa["Qtà"]/tot_qta*100).round(2)
    studio_isa["% Netto"] = (studio_isa["Netto"]/tot_netto*100).round(2)
    studio_isa = pd.concat([studio_isa, pd.DataFrame([["Totale",tot_qta,tot_netto,100,100]], columns=studio_isa.columns)], ignore_index=True)

    # --- Grafico ---
    fig, ax = plt.subplots(figsize=(8,5))
    ax.bar(studio_isa["FamigliaCategoria"], studio_isa["Netto"], color="skyblue")
    ax.set_title("Somma Netto per FamigliaCategoria")
    plt.xticks(rotation=45, ha="right")
    buf = BytesIO(); plt.tight_layout(); plt.savefig(buf, format="png"); buf.seek(0)

    # --- Excel ---
    wb = Workbook()
    ws = wb.active; ws.title = "Report"
    start_row, start_col = 3, 2
    total_fill = PatternFill(start_color="FFF4B084", end_color="FFF4B084", fill_type="solid")

    headers = ["FamigliaCategoria","Qtà","Netto","% Qtà","% Netto"]
    for j,h in enumerate(headers,start=start_col):
        ws.cell(row=start_row,column=j,value=h).font = Font(bold=True)
    for i,row in enumerate(dataframe_to_rows(studio_isa,index=False,header=False),start=start_row+1):
        for j,v in enumerate(row,start=start_col):
            ws.cell(row=i,column=j,value=v)
    tot_row_idx = start_row+len(studio_isa)
    for j in range(start_col, start_col+len(headers)):
        c=ws.cell(row=tot_row_idx,column=j)
        c.font=Font(bold=True); c.fill=total_fill
    img=XLImage(buf); img.anchor=f"A{tot_row_idx+3}"; ws.add_image(img)
    out=BytesIO(); wb.save(out)
    st.download_button("⬇️ Scarica report Excel", data=out.getvalue(), file_name="StudioISA_Report.xlsx")

if __name__ == "__main__":
    main()

