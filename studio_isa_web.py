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


# === CONFIG ===
st.set_page_config(page_title="Studio ISA", page_icon="🐾", layout="centered")
st.title("🐾 Studio ISA")
st.caption("Versione completa con riconoscimento automatico categorie e pivot Excel-like")


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

# Trova la colonna descrizione
desc_col = next((c for c in df.columns if "descrizione" in c.lower()), None)


# === REGOLE DI CATEGORIZZAZIONE ===
_RULES = {
    "LABORATORIO": [
        "analisi", "emocromo", "urine", "pu/cu", "urinocoltur",
        "coprolog", "feci", "giardia", "citolog", "citologia", "auricolare",
        "istolog", "test", "titolazione", "4 dx"
    ],
    "VISITE": [
        "visita", "controllo", "check", "visita con citologia",
        "visita dermatologica", "controllo dermatologico"
    ],
    "VACCINI": [
        "vacc", "vaccino", "vaccini", "rabbia", "trivalente",
        "leish", "letifend", "felv"
    ],
    "CHIRURGIA": [
        "castraz", "ovariect", "sterilizz", "intervento", "chirurg",
        "detartrasi", "estraz", "ovariectomia", "sutura"
    ],
    "MEDICINA": [
        "terapia", "terapie", "flebo", "day hospital", "pressione",
        "sedazione", "endovena", "ciclo", "cytopoint"
    ],
    "FAR": [
        "meloxidyl", "meloxidil", "apoquel", "konclav", "profenacarp", "cylanic",
        "mitex", "osurnia", "mometamax", "royal canin", "cortotic",
        "aristos", "stomorgyl", "previcox", "stronghold", "procox", "nexgard",
        "milbemax", "ronaxan", "panacur", "ciclosporin", "ciclosporina",
        "mg", "cpr", "compresse", "blister", "gocce", "sosp. orale", "ml"
    ],
    "DIAGNOSTICA PER IMMAGINI": [
        "ecografia", "radiografia", "radiolog", "radiograf", "lastra", "eco", "rx"
    ],
    "CHIP": ["microchip"],
    "ALTRE PRESTAZIONI": ["cremazion", "trasporto", "unghie", "otoematoma", "eutanasia"]
}
_FORCE = {"visita e terapia": "VISITE", "emedog": "MEDICINA", "cerenia": "MEDICINA"}


def _norm(s):
    return str(s).lower().strip()


def _has_any(text, keywords):
    for kw in keywords:
        if kw in text:
            return True
    return False


def rule_based_category(text):
    t = _norm(text)
    for phrase, cat in _FORCE.items():
        if phrase in t:
            return cat
    for cat, kwlist in _RULES.items():
        if _has_any(t, kwlist):
            return cat
    return "ALTRE PRESTAZIONI"


# === COMPLETA FAMIGLIA CATEGORIA ===
if desc_col:
    st.info("🔎 Rilevata colonna descrizione: completamento automatico delle categorie…")
    filled = 0
    for i, row in df.iterrows():
        if not row["FamigliaCategoria"] or row["FamigliaCategoria"].lower() == "nan":
            categoria = rule_based_category(row[desc_col])
            df.at[i, "FamigliaCategoria"] = categoria
            filled += 1
    st.success(f"Completate automaticamente {filled} righe.")


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
studio_isa = pd.concat([
    studio_isa,
    pd.DataFrame([{
        "FamigliaCategoria": "Totale",
        "Qtà": tot_qta,
        "Netto": tot_netto,
        "% Qtà": 100.0,
        "% Netto": 100.0
    }])
], ignore_index=True)


# === CREA PIVOT (identica a Excel) ===
pivot = (
    pd.pivot_table(df, values=["Perc", "Netto"], index=["FamigliaCategoria"], aggfunc="sum", fill_value=0)
      .reset_index()
      .rename(columns={"Perc": "Somma_Qta", "Netto": "Somma_Netto"})
)
pivot.loc[len(pivot)] = ["Totale", pivot["Somma_Qta"].sum(), pivot["Somma_Netto"].sum()]


# === CREA GRAFICO ===
fig, ax = plt.subplots(figsize=(8, 4))
pivot_no_total = pivot[pivot["FamigliaCategoria"] != "Totale"]
ax.bar(pivot_no_total["FamigliaCategoria"], pivot_no_total["Somma_Qta"], label="Somma_Qta")
ax.bar(pivot_no_total["FamigliaCategoria"], pivot_no_total["Somma_Netto"], label="Somma_Netto", alpha=0.7)
ax.legend()
ax.set_title("Pivot - Somma_Qta e Somma_Netto per FamigliaCategoria")
plt.xticks(rotation=45, ha="right")
plt.tight_layout()
buf = io.BytesIO()
plt.savefig(buf, format="png", dpi=150)
plt.close(fig)
buf.seek(0)


# === ESTRAI ANNO ===
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


# === SCRITTURA FILE EXCEL ===
wb = Workbook()
ws = wb.active
ws.title = "Report"

start_row, start_col = 3, 2

# Definisci colore per i totali (giallo chiaro)
total_fill = PatternFill(start_color="FFF4B084", end_color="FFF4B084", fill_type="solid")

# Tabella Studio ISA
isa_headers = ["FamigliaCategoria", "Qtà", "Netto", "% Qtà", "% Netto"]
for j, h in enumerate(isa_headers, start=start_col):
    ws.cell(row=start_row, column=j, value=h).font = Font(bold=True)
for i, row in enumerate(dataframe_to_rows(studio_isa, index=False, header=False), start=start_row+1):
    for j, v in enumerate(row, start=start_col):
        ws.cell(row=i, column=j, value=v)

# Totale in grassetto + giallo
tot_row = start_row + len(studio_isa)
for j in range(start_col, start_col + len(isa_headers)):
    cell = ws.cell(row=tot_row, column=j)
    cell.font = Font(bold=True)
    cell.fill = total_fill

# Pivot accanto
piv_col = start_col + 7
piv_headers = ["FamigliaCategoria", "Somma_Qta", "Somma_Netto"]
for j, h in enumerate(piv_headers, start=piv_col):
    ws.cell(row=start_row, column=j, value=h).font = Font(bold=True)
for i, row in enumerate(dataframe_to_rows(pivot, index=False, header=False), start=start_row+1):
    for j, v in enumerate(row, start=piv_col):
        ws.cell(row=i, column=j, value=v)

# Totale pivot in grassetto + giallo
tot_row_piv = start_row + len(pivot)
for j in range(piv_col, piv_col + len(piv_headers)):
    cell = ws.cell(row=tot_row_piv, column=j)
    cell.font = Font(bold=True)
    cell.fill = total_fill

# Grafico sotto pivot
img = XLImage(buf)
img.anchor = f"A{start_row + len(pivot) + 4}"
ws.add_image(img)

# Foglio DatiOriginali
ws_data = wb.create_sheet("DatiOriginali")
for r in dataframe_to_rows(df, index=False, header=True):
    ws_data.append(r)

# Auto-larghezza
for ws in [ws, ws_data]:
    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 40)

# Salva in memoria
output = io.BytesIO()
wb.save(output)
output.seek(0)

# === DOWNLOAD ===
st.success("✅ File generato con successo!")
st.download_button("📥 Scarica Excel", data=output, file_name=out_name,
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# === ANTEPRIME ===
col1, col2 = st.columns(2)
with col1:
    st.markdown("### Tabella Studio ISA")
    st.dataframe(studio_isa)
with col2:
    st.markdown("### Pivot per FamigliaCategoria")
    st.dataframe(pivot)


