import io
import re
import json
from datetime import datetime

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as XLImage


# =============== CONFIG STREAMLIT ===============
st.set_page_config(page_title="Studio ISA", page_icon="🐾", layout="centered")
st.title("🐾 Studio ISA (Web)")
st.caption("Upload Excel → Autoclassifica → Tabella + Pivot → Grafico → Download Excel")

# =============== UTILS: DIZIONARIO ===============
CATEGORIES = [
    "ALTRE PRESTAZIONI",
    "CHIP",
    "CHIRURGIA",
    "DIAGNOSTICA PER IMMAGINI",
    "FAR",
    "LABORATORIO",
    "MEDICINA",
    "VACCINI",
    "VISITE",
]

DEFAULT_KEYWORDS = {
    "ALTRE PRESTAZIONI": ["cremazion", "trasporto", "unghie", "otoematoma", "eutanasia"],
    "CHIP": ["microchip"],
    "CHIRURGIA": ["chirurg", "castraz", "ovariect", "sterilizz", "intervento", "sutura", "estraz", "detartrasi"],
    "DIAGNOSTICA PER IMMAGINI": ["rx", "radiograf", "radiolog", "eco", "ecografia", "lastra", "radiografia"],
    "FAR": [
        "meloxidyl", "meloxidil", "apoquel", "konclav", "profenacarp", "cylanic",
        "mitex", "osurnia", "mometamax", "royal canin", "cortotic",
        "aristos", "stomorgyl", "previcox", "stronghold", "procox", "nexgard",
        "milbemax", "letifend", "ronaxan", "panacur", "ciclosporin", "ciclosporina",
        "mg", "cpr", "compresse", "blister", "gocce", "sosp. orale", "ml"
    ],
    "LABORATORIO": [
        "analisi", "emocromo", "urine", "pu/cu", "urinocoltur",
        "coprolog", "feci", "giardia", "citolog", "citologia", "auricolare",
        "istolog", "test", "titolazione", "4 dx"
    ],
    "MEDICINA": [
        "terapia", "terapie", "flebo", "day hospital", "pressione",
        "sedazione", "endovena", "ciclo", "emedog", "cerenia", "cytopoint"
    ],
    "VACCINI": ["vacc", "vaccino", "vaccini", "rabbia", "trivalente", "leish", "letifend", "felv"],
    "VISITE": ["visita", "controllo", "check"],
}

def init_keywords():
    if "keywords" not in st.session_state:
        # base
        data = {cat: list(words) for cat, words in DEFAULT_KEYWORDS.items()}
        # se l'utente carica un json, verrà unito a runtime
        st.session_state["keywords"] = data

def merge_keywords(custom):
    """Unisce un dizionario custom nel dizionario in sessione, evitando duplicati."""
    kws = st.session_state["keywords"]
    for cat in CATEGORIES:
        kws.setdefault(cat, [])
    for cat, words in custom.items():
        if cat not in CATEGORIES:
            continue
        base = set(w.lower() for w in kws[cat])
        for w in words:
            if w.lower() not in base:
                kws[cat].append(w)
                base.add(w.lower())

def download_keywords_button():
    data = json.dumps(st.session_state["keywords"], ensure_ascii=False, indent=2)
    st.download_button("📥 Scarica dizionario (keywords.json)", data=data, file_name="keywords.json", mime="application/json")


# =============== CLASSIFICAZIONE ===============
def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip().lower())

def _has_any(text: str, keywords) -> bool:
    """Match preciso (parole intere; se la keyword contiene numeri tipo 50mg/10cpr, consenti match parziale)."""
    t = _norm(text)
    for kw in keywords:
        k = kw.lower().strip()
        if not k:
            continue
        if any(ch.isdigit() for ch in k):
            if k in t:
                return True
        else:
            if re.search(rf"(?<![a-zA-Z]){re.escape(k)}(?![a-zA-Z])", t):
                return True
    return False

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
        "sedazione", "endovena", "ciclo", "emedog", "cerenia", "cytopoint"
    ],
    "FAR": [
        "meloxidyl", "meloxidil", "apoquel", "konclav", "profenacarp", "cylanic",
        "mitex", "osurnia", "mometamax", "royal canin", "cortotic",
        "aristos", "stomorgyl", "previcox", "stronghold", "procox", "nexgard",
        "milbemax", "letifend", "ronaxan", "panacur", "ciclosporin", "ciclosporina",
        "mg", "cpr", "compresse", "blister", "gocce", "sosp. orale", "ml"
    ],
    "DIAGNOSTICA PER IMMAGINI": [
        "ecografia", "radiografia", "radiolog", "radiograf", "lastra", "eco", "rx"
    ],
    "CHIP": ["microchip"],
    "ALTRE PRESTAZIONI": ["cremazion", "trasporto", "unghie", "otoematoma", "eutanasia"]
}
_FORCE = {"visita e terapia": "VISITE"}

def rule_based_category(text: str) -> str:
    t = _norm(text)
    for phrase, cat in _FORCE.items():
        if phrase in t:
            return cat
    if _has_any(t, _RULES["LABORATORIO"]): return "LABORATORIO"
    if _has_any(t, _RULES["VISITE"]): return "VISITE"
    if _has_any(t, _RULES["VACCINI"]): return "VACCINI"
    if _has_any(t, _RULES["CHIRURGIA"]): return "CHIRURGIA"
    if _has_any(t, _RULES["MEDICINA"]): return "MEDICINA"
    if _has_any(t, _RULES["FAR"]): return "FAR"
    if _has_any(t, _RULES["DIAGNOSTICA PER IMMAGINI"]): return "DIAGNOSTICA PER IMMAGINI"
    if _has_any(t, _RULES["CHIP"]): return "CHIP"
    if _has_any(t, _RULES["ALTRE PRESTAZIONI"]): return "ALTRE PRESTAZIONI"
    return "ALTRE PRESTAZIONI"


# =============== MAPPATURA COLONNE ===============
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


# =============== SIDEBAR: Dizionario & Opzioni ===============
init_keywords()
with st.sidebar:
    st.subheader("🧠 Dizionario")
    uploaded_dict = st.file_uploader("Carica keywords.json (opzionale)", type=["json"], key="kwup")
    if uploaded_dict:
        try:
            custom = json.load(uploaded_dict)
            merge_keywords(custom)
            st.success("Dizionario caricato e unito ✅")
        except Exception as e:
            st.error(f"JSON non valido: {e}")

    download_keywords_button()
    st.divider()
    supervised = st.toggle("Apprendimento supervisionato", value=False, help="Suggerisce nuove parole chiave da apprendere in questa sessione")


# =============== UPLOAD EXCEL ===============
uploaded_file = st.file_uploader("📤 Carica file Excel (.xlsx)", type=["xlsx"])
if not uploaded_file:
    st.stop()

# =============== ELABORAZIONE ===============
progress = st.progress(0, text="🔍 Lettura file...")
try:
    df = pd.read_excel(uploaded_file)
    progress.progress(10, text="🔠 Mappo le colonne...")
    df = df.rename(columns=map_and_clean_columns(df.columns))

    # controlli minimi
    required = ['FamigliaCategoria', 'Perc', 'Netto']
    missing = [r for r in required if r not in df.columns]
    if missing:
        st.error(f"Mancano colonne obbligatorie dopo la mappatura: {missing}")
        st.stop()

    df['FamigliaCategoria'] = df['FamigliaCategoria'].astype(str).str.strip()

    # colonna descrizione per auto-fill
    desc_col = next((c for c in df.columns if "descrizione" in c.lower()), None)

    # auto fill delle categorie mancanti
    progress.progress(25, text="🧩 Completo FamigliaCategoria...")
    if desc_col:
        kws = st.session_state["keywords"]
        learned = {}
        filled = 0
        for i, row in df.iterrows():
            if not str(row["FamigliaCategoria"]).strip():
                descr = str(row[desc_col]).strip()
                cat = rule_based_category(descr)
                if cat == "ALTRE PRESTAZIONI" and supervised:
                    # apprendimento semplice: proponi token con numeri o parola lunga
                    t = re.sub(r"[=+.,;:!?\(\)\[\]]", " ", descr)
                    tokens = [p.lower() for p in t.split() if p.strip()]
                    token = next((p for p in tokens if any(ch.isdigit() for ch in p)), None) or next((p for p in tokens if len(p) >= 5), "")
                    if token:
                        learned.setdefault(token, "ALTRE PRESTAZIONI")
                df.at[i, "FamigliaCategoria"] = cat
                filled += 1

        if supervised and learned:
            st.info("🔎 Nuovi termini rilevati: conferma la categoria per apprenderli")
            for token in sorted(learned.keys()):
                cat_sel = st.selectbox(f"Termine: **{token}** → categoria:", CATEGORIES, index=CATEGORIES.index("ALTRE PRESTAZIONI"), key=f"learn_{token}")
                if st.button(f"Aggiungi '{token}' a {cat_sel}", key=f"add_{token}"):
                    st.session_state["keywords"][cat_sel].append(token)
                    st.success(f"Aggiunto '{token}' a {cat_sel}")

    # Studio ISA
    progress.progress(45, text="📊 Creo tabella Studio ISA...")
    studio_isa = (
        df.groupby('FamigliaCategoria', dropna=False)
          .agg({'Perc': 'sum', 'Netto': 'sum'})
          .reset_index()
          .rename(columns={'Perc': 'Qtà'})
    )
    tot_qta = studio_isa['Qtà'].sum()
    tot_netto = studio_isa['Netto'].sum()
    studio_isa['% Qtà'] = (studio_isa['Qtà'] / (tot_qta if tot_qta else 1) * 100).round(2)
    studio_isa['% Netto'] = (studio_isa['Netto'] / (tot_netto if tot_netto else 1) * 100).round(2)

    total_row = pd.DataFrame({
        'FamigliaCategoria': ['Totale'],
        'Qtà': [tot_qta],
        'Netto': [tot_netto],
        '% Qtà': [100.0],
        '% Netto': [100.0],
    })
    studio_isa = pd.concat([studio_isa, total_row], ignore_index=True)

    # Pivot (replica logica Excel)
    progress.progress(65, text="🧮 Creo pivot...")
    pivot = (
        pd.pivot_table(df, values=['Perc', 'Netto'], index=['FamigliaCategoria'], aggfunc='sum', fill_value=0)
          .reset_index()
          .rename(columns={'Perc': 'Somma_Qta', 'Netto': 'Somma_Netto'})
    )

    # Grafico
    progress.progress(80, text="📈 Genero grafico...")
    fig, ax = plt.subplots(figsize=(9, 4))
    pivot.plot(kind='bar', x='FamigliaCategoria', y=['Somma_Qta', 'Somma_Netto'], ax=ax)
    ax.set_title("Somma_Qta e Somma_Netto per FamigliaCategoria")
    ax.set_ylabel("Valore")
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    img_buf = io.BytesIO()
    plt.savefig(img_buf, format='png', dpi=120)
    plt.close(fig)
    img_buf.seek(0)

    # Anno dal primo campo data disponibile
    anno = None
    for c in df.columns:
        if "data" in c.lower():
            series = pd.to_datetime(df[c], errors='coerce').dropna()
            if not series.empty:
                anno = int(series.iloc[0].year)
                break
    if not anno:
        anno = datetime.now().year
    out_name = f"Studio_ISA_{anno}.xlsx"

    # Excel con openpyxl
    progress.progress(90, text="📦 Impagino Excel...")
    wb = Workbook()

    # Report
    ws_rep = wb.active
    ws_rep.title = "Report"
    # Titolo
    ws_rep["A1"] = "Studio ISA"
    ws_rep["A1"].font = Font(bold=True, size=14)

    # Tabella Studio ISA (a partire da riga 3, col B)
    start_row, start_col = 3, 2
    headers = ["FamigliaCategoria", "Qtà", "Netto", "% Qtà", "% Netto"]
    ws_rep.cell(row=start_row, column=start_col, value=headers[0]).font = Font(bold=True)
    for j, h in enumerate(headers[1:], start=start_col+1):
        ws_rep.cell(row=start_row, column=j, value=h).font = Font(bold=True)

    for i, r in enumerate(dataframe_to_rows(studio_isa, index=False, header=False), start=start_row+1):
        for j, v in enumerate(r, start=start_col):
            ws_rep.cell(row=i, column=j, value=v)

    # Totale in grassetto (ultima riga della tabella)
    tot_row = start_row + len(studio_isa)
    for j in range(start_col, start_col + len(headers)):
        ws_rep.cell(row=tot_row, column=j).font = Font(bold=True)

    # Pivot a destra (stessa riga, colonna start_col + 7)
    piv_col = start_col + 7
    ws_rep.cell(row=start_row, column=piv_col, value="FamigliaCategoria").font = Font(bold=True)
    ws_rep.cell(row=start_row, column=piv_col+1, value="Somma_Qta").font = Font(bold=True)
    ws_rep.cell(row=start_row, column=piv_col+2, value="Somma_Netto").font = Font(bold=True)

    for i, r in enumerate(dataframe_to_rows(pivot, index=False, header=False), start=start_row+1):
        ws_rep.cell(row=i, column=piv_col,   value=r[0])
        ws_rep.cell(row=i, column=piv_col+1, value=r[1])
        ws_rep.cell(row=i, column=piv_col+2, value=r[2])

    # Grafico sotto la pivot come immagine
    img = XLImage(img_buf)
    img_row_anchor = start_row + len(pivot) + 3
    img.anchor = f"{'A'}{img_row_anchor}"  # ancora semplice (col A). Puoi regolare se vuoi
    ws_rep.add_image(img)

    # Foglio DatiOriginali
    ws_data = wb.create_sheet("DatiOriginali")
    for r in dataframe_to_rows(df, index=False, header=True):
        ws_data.append(r)

    # Formattazione minima (auto-larghezza)
    for ws in [ws_rep, ws_data]:
        dims = {}
        for row in ws.rows:
            for cell in row:
                if cell.value is not None:
                    dims[cell.column_letter] = max(dims.get(cell.column_letter, 0), len(str(cell.value)))
        for col, width in dims.items():
            ws.column_dimensions[col].width = min(max(10, width + 2), 40)

    # Salva in memoria
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    progress.progress(100, text="✅ Fatto!")
    st.success("File generato con successo.")
    st.download_button(
        "📥 Scarica Excel",
        data=output,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Anteprime utili
    with st.expander("👀 Anteprima Studio ISA"):
        st.dataframe(studio_isa)
    with st.expander("👀 Anteprima Pivot"):
        st.dataframe(pivot)

except Exception as e:
    st.error(f"Errore: {e}")
