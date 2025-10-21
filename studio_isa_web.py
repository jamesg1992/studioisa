import io, json, re, os
from datetime import datetime

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill
from openpyxl.drawing.image import Image as XLImage


# === CONFIG ===
st.set_page_config(page_title="Studio ISA (Beta)", page_icon="🐾", layout="centered")
st.title("🐾 Studio ISA (Beta)")
st.caption("Studio ISA con autocompletamento")


# === UPLOAD FILE ===
uploaded_file = st.file_uploader("📤 Carica file Excel (.xlsx)", type=["xlsx"])
if not uploaded_file:
    st.stop()


# === LETTURA E PULIZIA COLONNE ===
def map_and_clean_columns(cols):
    rename_map = {}
    for c in cols:
        s = str(c).strip()
        if s == "%":
            rename_map[c] = "Perc"
        elif "netto" in s.lower() and "dopo" in s.lower():
            rename_map[c] = "Netto"
        elif ("famiglia" in s.lower() and "categoria" in s.lower()) or ("famiglia" in s.lower() and "/" in s):
            rename_map[c] = "FamigliaCategoria"
        else:
            cleaned = re.sub(r"[^\w]", "", s)
            rename_map[c] = cleaned if cleaned else "Col"
    return rename_map


df = pd.read_excel(uploaded_file)
df = df.rename(columns=map_and_clean_columns(df.columns))

required = ["FamigliaCategoria", "Perc", "Netto"]
missing = [r for r in required if r not in df.columns]
if missing:
    st.error(f"Mancano colonne obbligatorie: {missing}")
    st.stop()

df["FamigliaCategoria"] = df["FamigliaCategoria"].astype(str).str.strip()
desc_col = next((c for c in df.columns if "descrizione" in c.lower()), None)


# === DIZIONARIO BASE ===
_RULES = {
    "LABORATORIO": ["analisi", "emocromo", "urine", "feci", "giardia", "test", "citolog", "istolog", "coprolog"],
    "VISITE": ["visita", "controllo", "check"],
    "VACCINI": ["vacc", "rabbia", "leish", "felv", "letifend"],
    "CHIRURGIA": ["castraz", "sterilizz", "chirurg", "detartrasi", "estraz", "ovariect"],
    "MEDICINA": ["terapia", "terapie", "flebo", "day hospital", "pressione", "ciclo", "endovena"],
    "FAR": ["meloxidyl", "apoquel", "konclav", "mitex", "osurnia", "cylanic", "royal", "mometamax", "milbemax", "stomorgyl", "previcox", "mg", "cpr"],
    "DIAGNOSTICA PER IMMAGINI": ["rx", "ecograf", "radiograf", "lastra", "eco"],
    "CHIP": ["microchip"],
    "ALTRE PRESTAZIONI": ["cremazion", "trasporto", "unghie", "otoematoma", "eutanasia"]
}

_FORCE = {"visita e terapia": "VISITE", "emedog": "MEDICINA", "cerenia": "MEDICINA"}


# === GESTIONE APPRENDIMENTO ===
MEMORY_FILE = "keywords_memory.json"

if os.path.exists(MEMORY_FILE):
    with open(MEMORY_FILE, "r", encoding="utf-8") as f:
        user_memory = json.load(f)
else:
    user_memory = {}

def _norm(s): return str(s).lower().strip()

def _has_any(text, keywords):
    return any(kw in text for kw in keywords)

def rule_based_category(text):
    t = _norm(text)
    if t in user_memory:
        return user_memory[t]
    for phrase, cat in _FORCE.items():
        if phrase in t:
            return cat
    for cat, kwlist in _RULES.items():
        if _has_any(t, kwlist):
            return cat
    return None  # non trovata


# === COMPLETA FAMIGLIA CATEGORIA + AUTOLEARNING ===
if desc_col:
    st.info("🔎 Completamento automatico categorie con apprendimento…")
    new_terms = {}
    for i, row in df.iterrows():
        if not row["FamigliaCategoria"] or row["FamigliaCategoria"].lower() == "nan":
            categoria = rule_based_category(row[desc_col])
            if categoria is None:
                term = _norm(row[desc_col])
                if term not in new_terms:
                    new_terms[term] = None
            else:
                df.at[i, "FamigliaCategoria"] = categoria

    if new_terms:
    st.warning(f"Trovati {len(new_terms)} termini sconosciuti da classificare.")

    # Memorizza la lista di termini nella sessione
    if "pending_terms" not in st.session_state:
        st.session_state.pending_terms = list(new_terms.keys())
        st.session_state.current_idx = 0

    terms = st.session_state.pending_terms
    idx = st.session_state.current_idx

    if idx < len(terms):
        current_term = terms[idx]
        st.markdown(f"### 🆕 {idx+1}/{len(terms)} — '{current_term}'")
        cat = st.selectbox("Seleziona la categoria corretta:", list(_RULES.keys()), key=f"term_{current_term}")

        col1, col2 = st.columns([1, 2])
        with col1:
            if st.button("✅ Salva e passa al prossimo"):
                user_memory[current_term] = cat
                with open(MEMORY_FILE, "w", encoding="utf-8") as f:
                    json.dump(user_memory, f, ensure_ascii=False, indent=2)
                st.session_state.current_idx += 1
                st.experimental_rerun()

        with col2:
            if st.button("⏹️ Interrompi"):
                st.success("Salvataggio interrotto. I progressi sono memorizzati.")
                st.stop()
    else:
        st.success("🎉 Tutti i nuovi termini sono stati classificati e salvati!")
        del st.session_state.pending_terms
        del st.session_state.current_idx

# === CREA TABELLA STUDIO ISA ===
studio_isa = (
    df.groupby("FamigliaCategoria", dropna=False)
      .agg({"Perc": "sum", "Netto": "sum"})
      .reset_index()
      .rename(columns={"Perc": "Qtà"})
)
tot_qta = studio_isa["Qtà"].sum()
tot_netto = studio_isa["Netto"].sum()
studio_isa["% Qtà"] = (studio_isa["Qtà"] / tot_qta * 100).round(2)
studio_isa["% Netto"] = (studio_isa["Netto"] / tot_netto * 100).round(2)
studio_isa.loc[len(studio_isa)] = ["Totale", tot_qta, tot_netto, 100, 100]


# === CREA PIVOT ===
pivot = (
    pd.pivot_table(df, values=["Perc", "Netto"], index=["FamigliaCategoria"], aggfunc="sum", fill_value=0)
      .reset_index()
      .rename(columns={"Perc": "Somma_Qta", "Netto": "Somma_Netto"})
)
pivot.loc[len(pivot)] = ["Totale", pivot["Somma_Qta"].sum(), pivot["Somma_Netto"].sum()]


# === GRAFICO ===
fig, ax = plt.subplots(figsize=(8, 4))
pivot_no_total = pivot[pivot["FamigliaCategoria"] != "Totale"]
ax.bar(pivot_no_total["FamigliaCategoria"], pivot_no_total["Somma_Qta"], label="Somma_Qta")
ax.bar(pivot_no_total["FamigliaCategoria"], pivot_no_total["Somma_Netto"], alpha=0.7, label="Somma_Netto")
ax.legend()
ax.set_title("Pivot - Somma_Qta e Somma_Netto per FamigliaCategoria")
plt.xticks(rotation=45, ha="right")
plt.tight_layout()
buf = io.BytesIO()
plt.savefig(buf, format="png", dpi=150)
plt.close(fig)
buf.seek(0)


# === SALVATAGGIO EXCEL ===
anno = None
for c in df.columns:
    if "data" in c.lower():
        s = pd.to_datetime(df[c], errors="coerce").dropna()
        if not s.empty:
            anno = int(s.iloc[0].year)
            break
anno = anno or datetime.now().year
out_name = f"Studio_ISA_{anno}.xlsx"

wb = Workbook()
ws = wb.active
ws.title = "Report"

start_row, start_col = 3, 2
total_fill = PatternFill(start_color="FFF4B084", end_color="FFF4B084", fill_type="solid")

isa_headers = ["FamigliaCategoria", "Qtà", "Netto", "% Qtà", "% Netto"]
for j, h in enumerate(isa_headers, start=start_col):
    ws.cell(row=start_row, column=j, value=h).font = Font(bold=True)
for i, row in enumerate(dataframe_to_rows(studio_isa, index=False, header=False), start=start_row+1):
    for j, v in enumerate(row, start=start_col):
        ws.cell(row=i, column=j, value=v)
tot_row = start_row + len(studio_isa)
for j in range(start_col, start_col + len(isa_headers)):
    c = ws.cell(row=tot_row, column=j)
    c.font = Font(bold=True)
    c.fill = total_fill

# Pivot accanto
piv_col = start_col + 7
piv_headers = ["FamigliaCategoria", "Somma_Qta", "Somma_Netto"]
for j, h in enumerate(piv_headers, start=piv_col):
    ws.cell(row=start_row, column=j, value=h).font = Font(bold=True)
for i, row in enumerate(dataframe_to_rows(pivot, index=False, header=False), start=start_row+1):
    for j, v in enumerate(row, start=piv_col):
        ws.cell(row=i, column=j, value=v)
tot_row_piv = start_row + len(pivot)
for j in range(piv_col, piv_col + len(piv_headers)):
    c = ws.cell(row=tot_row_piv, column=j)
    c.font = Font(bold=True)
    c.fill = total_fill

img = XLImage(buf)
img.anchor = f"A{start_row + len(pivot) + 4}"
ws.add_image(img)

ws_data = wb.create_sheet("DatiOriginali")
for r in dataframe_to_rows(df, index=False, header=True):
    ws_data.append(r)

for wsx in [ws, ws_data]:
    for col in wsx.columns:
        max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        wsx.column_dimensions[col[0].column_letter].width = min(max_len + 2, 40)

output = io.BytesIO()
wb.save(output)
output.seek(0)

st.success("✅ File generato con successo!")
st.download_button("📥 Scarica Excel", data=output, file_name=out_name,
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

col1, col2 = st.columns(2)
with col1:
    st.markdown("### Tabella Studio ISA")
    st.dataframe(studio_isa)
with col2:
    st.markdown("### Pivot per FamigliaCategoria")
    st.dataframe(pivot)

