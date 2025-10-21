import io
import re
from datetime import datetime

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font
from openpyxl.drawing.image import Image as XLImage


# === Config Streamlit ===
st.set_page_config(page_title="Studio ISA", page_icon="🐾", layout="centered")
st.title("🐾 Studio ISA (Web)")
st.caption("Upload Excel → Analisi → Tabella + Pivot + Grafico → Download Excel")

# === Upload file ===
uploaded_file = st.file_uploader("📤 Carica file Excel (.xlsx)", type=["xlsx"])
if not uploaded_file:
    st.stop()

# === Lettura file ===
df = pd.read_excel(uploaded_file)
st.info(f"File caricato con {df.shape[0]} righe e {df.shape[1]} colonne")

# === Pulizia colonne ===
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
            rename_map[c] = re.sub(r"[^\w]", "", s)
    return rename_map

df = df.rename(columns=map_and_clean_columns(df.columns))

# === Controllo colonne ===
required = ["FamigliaCategoria", "Perc", "Netto"]
missing = [r for r in required if r not in df.columns]
if missing:
    st.error(f"Mancano colonne obbligatorie: {missing}")
    st.stop()

df["FamigliaCategoria"] = df["FamigliaCategoria"].astype(str).str.strip()

# === Calcola Studio ISA ===
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

total_row = pd.DataFrame({
    "FamigliaCategoria": ["Totale"],
    "Qtà": [tot_qta],
    "Netto": [tot_netto],
    "% Qtà": [100.0],
    "% Netto": [100.0]
})
studio_isa = pd.concat([studio_isa, total_row], ignore_index=True)

# === Pivot (uguale a quella Excel) ===
pivot = (
    pd.pivot_table(df, values=["Perc", "Netto"], index=["FamigliaCategoria"], aggfunc="sum", fill_value=0)
      .reset_index()
      .rename(columns={"Perc": "Somma_Qta", "Netto": "Somma_Netto"})
)
pivot.loc["Totale"] = ["Totale", pivot["Somma_Qta"].sum(), pivot["Somma_Netto"].sum()]

# === Grafico (stesso stile Excel) ===
fig, ax = plt.subplots(figsize=(8, 4))
pivot_no_total = pivot[pivot["FamigliaCategoria"] != "Totale"]
ax.bar(pivot_no_total["FamigliaCategoria"], pivot_no_total["Somma_Qta"], label="Somma_Qta")
ax.bar(pivot_no_total["FamigliaCategoria"], pivot_no_total["Somma_Netto"], label="Somma_Netto", alpha=0.7)
ax.set_title("Pivot - Somma_Qta e Somma_Netto per FamigliaCategoria")
ax.legend()
plt.xticks(rotation=45, ha="right")
plt.tight_layout()
buf = io.BytesIO()
plt.savefig(buf, format="png", dpi=150)
plt.close(fig)
buf.seek(0)

# === Anno dal file ===
anno = None
for c in df.columns:
    if "data" in c.lower():
        series = pd.to_datetime(df[c], errors="coerce").dropna()
        if not series.empty:
            anno = int(series.iloc[0].year)
            break
if not anno:
    anno = datetime.now().year
out_name = f"Studio_ISA_{anno}.xlsx"

# === Scrittura Excel ===
wb = Workbook()
ws = wb.active
ws.title = "Report"

# Scrivi intestazione tabella ISA
isa_headers = ["FamigliaCategoria", "Qtà", "Netto", "% Qtà", "% Netto"]
start_row, start_col = 3, 2
for j, h in enumerate(isa_headers, start=start_col):
    ws.cell(row=start_row, column=j, value=h).font = Font(bold=True)
for i, row in enumerate(dataframe_to_rows(studio_isa, index=False, header=False), start=start_row+1):
    for j, v in enumerate(row, start=start_col):
        ws.cell(row=i, column=j, value=v)

# Scrivi Pivot accanto
piv_headers = ["FamigliaCategoria", "Somma_Qta", "Somma_Netto"]
piv_col = start_col + 7
for j, h in enumerate(piv_headers, start=piv_col):
    ws.cell(row=start_row, column=j, value=h).font = Font(bold=True)
for i, row in enumerate(dataframe_to_rows(pivot, index=False, header=False), start=start_row+1):
    for j, v in enumerate(row, start=piv_col):
        ws.cell(row=i, column=j, value=v)

# Inserisci grafico sotto la pivot
img = XLImage(buf)
img.anchor = f"A{start_row + len(pivot) + 4}"
ws.add_image(img)

# Auto-larghezza
for col in ws.columns:
    max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
    ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 40)

# Esporta file Excel
output = io.BytesIO()
wb.save(output)
output.seek(0)

# === Download & Anteprima ===
st.success("✅ File pronto")
st.download_button("📥 Scarica Excel", data=output, file_name=out_name,
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

col1, col2 = st.columns(2)
with col1:
    st.markdown("### Tabella Studio ISA")
    st.dataframe(studio_isa)
with col2:
    st.markdown("### Pivot per FamigliaCategoria")
    st.dataframe(pivot)
