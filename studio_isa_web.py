import streamlit as st
import pandas as pd
import json, os, re, time
from datetime import datetime

# ========== CONFIG ==========
_RULES = {
    "ALTRE PRESTAZIONI": ["trasporto", "cremazione", "eutanasia", "unghie"],
    "CHIP": ["microchip", "chip"],
    "CHIRURGIA": ["intervento", "castrazione", "sterilizzazione", "ovariectomia", "chirurgico"],
    "DIAGNOSTICA PER IMMAGINI": ["rx", "radiografia", "eco", "ecografia"],
    "FAR": ["meloxidyl", "enrox", "apoquel", "konclav", "cytopoint", "cylan", "previcox", "aristos", "mitex", "mometa", "profenacarp", "stomorgyl", "stronghold", "nexgard", "milbemax", "royal", "procox"],
    "LABORATORIO": ["analisi", "esame", "citologia", "istologico", "emocromo", "urine", "coprologico", "giardia", "test", "feci", "titolazione", "urinocoltura"],
    "MEDICINA": ["terapia", "flebo", "emedog", "cerenia", "cura", "day hospital", "trattamento"],
    "VACCINI": ["vaccino", "letifend", "rabbia", "felv", "trivalente", "4dx"],
    "VISITE": ["visita", "controllo", "dermatologica"]
}

LOCAL_JSON = "studio_isa_memory.json"

# ========== FUNZIONI BASE ==========

def clean_text(s):
    return str(s).strip().lower()

def detect_category(desc):
    s = clean_text(desc)
    for cat, kws in _RULES.items():
        for kw in kws:
            if kw in s:
                return cat
    return None

def load_memory():
    if os.path.exists(LOCAL_JSON):
        with open(LOCAL_JSON, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_memory(mem):
    with open(LOCAL_JSON, "w", encoding="utf-8") as f:
        json.dump(mem, f, ensure_ascii=False, indent=2)

def classify_missing(df):
    df = df.copy()
    df["FamigliaCategoria"] = df["FamigliaCategoria"].astype(str)
    for i, row in df.iterrows():
        if not row["FamigliaCategoria"] or row["FamigliaCategoria"].lower() == "nan":
            desc = str(row.get("Descrizione (da archivio DrVeto)", ""))
            cat = detect_category(desc)
            if cat:
                df.at[i, "FamigliaCategoria"] = cat
    return df

# ========== STREAMLIT UI ==========

st.set_page_config(page_title="Studio ISA", layout="centered")

# Splash iniziale
st.markdown(
    """
    <div style='text-align:center;'>
        <h1 style='color:#1E90FF;'>💼 Studio ISA</h1>
        <p style='color:gray;'>Analisi automatizzata Excel con apprendimento intelligente</p>
    </div>
    """, unsafe_allow_html=True)

uploaded_file = st.file_uploader("📁 Carica il file Excel", type=["xlsx"])

if uploaded_file:
    with st.spinner("🔍 Lettura del file in corso..."):
        df = pd.read_excel(uploaded_file)

    rename_map = {}
    for c in df.columns:
        s = str(c).strip()
        if s == "%": rename_map[c] = "Perc"
        elif "netto" in s.lower() and "dopo" in s.lower(): rename_map[c] = "Netto"
        elif "famiglia" in s.lower() and "categoria" in s.lower(): rename_map[c] = "FamigliaCategoria"
        else: rename_map[c] = re.sub(r"[^\w]", "", s)
    df = df.rename(columns=rename_map)

    # Riempi i vuoti automaticamente
    df = classify_missing(df)

    # Carica memoria
    memory = load_memory()

    # Trova termini nuovi non ancora mappati
    desc_terms = sorted(set(df["Descrizione_daarchivioDrVeto"].astype(str).str.strip().unique()))
    known = set(memory.keys())
    new_terms = [t for t in desc_terms if t and t not in known]

    st.markdown(f"### 📊 {len(new_terms)} nuovi termini da classificare")

    if "idx" not in st.session_state:
        st.session_state.idx = 0
    if "saved" not in st.session_state:
        st.session_state.saved = False
    if "local_updates" not in st.session_state:
        st.session_state.local_updates = {}

    pending = new_terms
    idx = st.session_state.idx

    # ========================= CLASSIFICAZIONE =========================
    if pending:
        term = pending[idx]
        total_terms = len(pending)
        st.info(f"Termine {idx+1}/{total_terms}")

        default_cat = "ALTRE PRESTAZIONI"
        cat_options = list(_RULES.keys())

        # select persistente
        selected_cat = st.selectbox(
            f"Categoria per “{term}”:",
            cat_options,
            key=f"select_{term}"
        )

        c1, c2, c3 = st.columns([1, 1, 2])

        with c1:
            if st.button("✅ Salva locale", key=f"save_{idx}"):
                st.session_state.local_updates[term] = selected_cat
                st.session_state.saved = True

        with c2:
            if st.button("⏭️ Salta", key=f"skip_{idx}"):
                st.session_state.saved = True

        with c3:
            if st.button("💾 Salva tutto su GitHub", type="primary"):
                memory.update(st.session_state.local_updates)
                save_memory(memory)
                st.session_state.local_updates = {}
                st.session_state.idx = 0
                st.session_state.saved = False
                st.success("✅ Tutti i nuovi termini salvati!")
                st.experimental_rerun()

        # --- Avanzamento automatico ---
        if st.session_state.saved:
            st.session_state.saved = False
            st.session_state.idx += 1
            if st.session_state.idx >= len(pending):
                st.session_state.idx = 0
                st.success("🎉 Tutti classificati! Ora puoi salvare definitivamente.")
            st.experimental_rerun()

        st.progress((idx + 1) / total_terms)
    else:
        st.success("✅ Nessun nuovo termine da classificare!")

    # ========================= ELABORAZIONE =========================
    st.divider()
    if st.button("📈 Genera Report Studio ISA"):
        with st.spinner("Elaborazione pivot in corso..."):
            df["FamigliaCategoria"] = df["FamigliaCategoria"].fillna("ALTRE PRESTAZIONI")

            studio_isa = df.groupby("FamigliaCategoria", dropna=False).agg({
                "Perc": "sum",
                "Netto": "sum"
            }).reset_index().rename(columns={"Perc": "Qtà"})

            tot_qta = studio_isa["Qtà"].sum()
            tot_netto = studio_isa["Netto"].sum()
            studio_isa["% Qtà"] = (studio_isa["Qtà"]/tot_qta*100).round(2)
            studio_isa["% Netto"] = (studio_isa["Netto"]/tot_netto*100).round(2)

            # Totale
            totale = pd.DataFrame([{
                "FamigliaCategoria": "Totale",
                "Qtà": tot_qta,
                "Netto": tot_netto,
                "% Qtà": 100,
                "% Netto": 100
            }])
            studio_isa = pd.concat([studio_isa, totale], ignore_index=True)

            st.dataframe(
                studio_isa.style.highlight_max(axis=0, color="lightyellow")
                            .set_properties(**{"font-weight": "bold"}, subset=["FamigliaCategoria"])
            )

            # Salvataggio Excel finale
            year = datetime.now().year
            output_name = f"Studio_ISA_{year}.xlsx"
            studio_isa.to_excel(output_name, index=False)
            st.success(f"✅ File generato: {output_name}")
            with open(output_name, "rb") as f:
                st.download_button("⬇️ Scarica Excel", f, file_name=output_name)

else:
    st.info("👆 Carica un file Excel per iniziare l'analisi.")
