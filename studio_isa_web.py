import streamlit as st
import pandas as pd
import json
import re
import os
import base64
import requests
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

# === GITHUB FUNZIONI ===
def github_load_json():
    try:
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_FILE}"
        headers = {"Authorization": f"token {GITHUB_TOKEN}"}
        res = requests.get(url, headers=headers)
        if res.status_code == 200:
            content = base64.b64decode(res.json()["content"])
            return json.loads(content.decode("utf-8"))
        return {}
    except Exception as e:
        st.warning(f"⚠️ Errore caricamento dizionario GitHub: {e}")
        return {}

def github_save_json(data):
    try:
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_FILE}"
        headers = {"Authorization": f"token {GITHUB_TOKEN}"}
        get_res = requests.get(url, headers=headers)
        sha = get_res.json().get("sha") if get_res.status_code == 200 else None
        encoded = base64.b64encode(json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")).decode("utf-8")
        payload = {
            "message": "Aggiornamento dizionario Studio ISA",
            "content": encoded,
            "branch": "main",
            "sha": sha
        }
        res = requests.put(url, headers=headers, data=json.dumps(payload))
        if res.status_code in (200, 201):
            st.success("✅ Dizionario aggiornato su GitHub!")
        else:
            st.error(f"❌ Errore aggiornando su GitHub: {res.status_code}")
    except Exception as e:
        st.error(f"❌ Errore salvataggio GitHub: {e}")

# === CACHE FILE EXCEL ===
@st.cache_data(show_spinner=False)
def load_excel(uploaded_file):
    return pd.read_excel(uploaded_file)

# === REGOLE DI CLASSIFICAZIONE ===
_RULES = {
    "LABORATORIO": ["analisi", "emocromo", "test", "esame", "coprologico", "feci", "giardia", "leishmania"],
    "VISITE": ["visita", "controllo", "consulto", "dermatologico"],
    "FAR": ["meloxidyl", "konclav", "enrox", "profenacarp", "apoquel", "osurnia", "cylanic", "mometa", "aristos", "cytopoint", "milbemax"],
    "CHIRURGIA": ["intervento", "chirurgico", "castrazione", "sterilizzazione", "ovariectomia", "detartrasi", "estrazione"],
    "DIAGNOSTICA PER IMMAGINI": ["rx", "radiografia", "eco", "ecografia", "tac"],
    "MEDICINA": ["terapia", "flebo", "day hospital", "trattamento", "emedog", "cerenia"],
    "VACCINI": ["vaccino", "letifend", "rabbia", "trivalente"],
    "CHIP": ["microchip"],
    "ALTRE PRESTAZIONI": ["trasporto", "eutanasia", "unghie", "cremazione"]
}

# === CLASSIFICAZIONE ===
def classify(desc, fam_value, memory):
    if pd.notna(fam_value) and str(fam_value).strip():
        return fam_value
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
    st.title("📊 Studio ISA – Web App (Fast v2 🚀)")

    uploaded = st.file_uploader("📁 Seleziona il file Excel", type=["xlsx", "xls"])
    if not uploaded:
        st.info("Carica un file Excel per iniziare.")
        return

    # --- Caricamento cache ---
    if "df" not in st.session_state:
        st.session_state.df = load_excel(uploaded)
        st.session_state.user_memory = github_load_json()
        st.session_state.new_terms = {}
        st.session_state.idx = 0
        st.session_state.refresh_trigger = False

    df = st.session_state.df
    user_memory = st.session_state.user_memory

    # --- Rileva colonne principali ---
    col_desc = next((c for c in df.columns if "descrizione" in c.lower()), None)
    col_fam = next((c for c in df.columns if "famiglia" in c.lower()), None)
    col_netto = next((c for c in df.columns if "netto" in c.lower() and "dopo" in c.lower()), None)
    col_perc = next((c for c in df.columns if c.strip() == "%"), None)
    if not all([col_desc, col_fam, col_netto, col_perc]):
        st.error("❌ Impossibile trovare tutte le colonne richieste nel file.")
        return

    # --- Classifica righe ---
    df["FamigliaCategoria"] = df.apply(
        lambda r: classify(r[col_desc], r[col_fam], user_memory), axis=1
    )

    # --- Scopri nuovi termini ---
    new_terms = {}
    for val in df[col_desc].unique():
        v = str(val).lower()
        if not any(k.lower() in v for k in user_memory.keys()):
            if not any(k in v for keys in _RULES.values() for k in keys):
                new_terms[val] = ""

    if new_terms:
        st.warning(f"Trovati {len(new_terms)} nuovi termini non classificati.")
        terms = list(new_terms.keys())

        if st.session_state.idx < len(terms):
            term = terms[st.session_state.idx]
            st.markdown(f"### 🆕 {st.session_state.idx+1}/{len(terms)} — '{term}'")
            cat = st.selectbox("Categoria corretta:", list(_RULES.keys()), key=f"sel_{term}")

            col1, col2, col3 = st.columns([1, 1, 2])
            with col1:
                if st.button("✅ Salva locale"):
                    user_memory[term] = cat
                    st.session_state.user_memory = user_memory
                    st.session_state.idx += 1
                    st.session_state.refresh_trigger = not st.session_state.refresh_trigger
                    st.success(f"💾 '{term}' salvato come {cat}")
            with col2:
                if st.button("⏹️ Interrompi"):
                    st.stop()
            with col3:
                if st.button("💾 Fine e salva su GitHub"):
                    github_save_json(user_memory)
                    st.session_state.idx = 0
                    st.toast("✅ Tutto salvato su GitHub!")
            return
        else:
            st.success("🎉 Tutti i nuovi termini sono stati classificati.")
            github_save_json(user_memory)
            st.session_state.idx = 0

    # --- Calcolo tabella Studio ISA ---
    studio_isa = df.groupby("FamigliaCategoria", dropna=False).agg({
        col_perc: "sum",
        col_netto: "sum"
    }).reset_index().rename(columns={col_perc: "Qtà", col_netto: "Netto"})

    tot_qta = studio_isa["Qtà"].sum()
    tot_netto = studio_isa["Netto"].sum()
    studio_isa["% Qtà"] = (studio_isa["Qtà"] / tot_qta * 100).round(2)
    studio_isa["% Netto"] = (studio_isa["Netto"] / tot_netto * 100).round(2)
    studio_isa = pd.concat([
        studio_isa,
        pd.DataFrame([["Totale", tot_qta, tot_netto, 100, 100]], columns=studio_isa.columns)
    ])

    st.success("✅ Tabella Studio ISA generata.")

    # --- Grafico ---
    fig, ax = plt.subplots(figsize=(8, 5))
    ax.bar(studio_isa["FamigliaCategoria"], studio_isa["Netto"], color="skyblue")
    ax.set_title("Somma Netto per FamigliaCategoria")
    plt.xticks(rotation=45, ha="right")
    buf = BytesIO()
    plt.tight_layout()
    plt.savefig(buf, format="png")
    buf.seek(0)

    # --- Excel Report ---
    wb = Workbook()
    ws = wb.active
    ws.title = "Report"
    start_row, start_col = 3, 2
    total_fill = PatternFill(start_color="FFF4B084", end_color="FFF4B084", fill_type="solid")

    headers = ["FamigliaCategoria", "Qtà", "Netto", "% Qtà", "% Netto"]
    for j, h in enumerate(headers, start=start_col):
        ws.cell(row=start_row, column=j, value=h).font = Font(bold=True)
    for i, row in enumerate(dataframe_to_rows(studio_isa, index=False, header=False), start=start_row+1):
        for j, v in enumerate(row, start=start_col):
            ws.cell(row=i, column=j, value=v)
    tot_row_idx = start_row + len(studio_isa)
    for j in range(start_col, start_col + len(headers)):
        c = ws.cell(row=tot_row_idx, column=j)
        c.font = Font(bold=True)
        c.fill = total_fill

    img = XLImage(buf)
    img.anchor = f"A{tot_row_idx + 3}"
    ws.add_image(img)

    out = BytesIO()
    wb.save(out)
    st.download_button("⬇️ Scarica report Excel", data=out.getvalue(), file_name="StudioISA_Report.xlsx")

if __name__ == "__main__":
    main()

