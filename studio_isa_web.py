import streamlit as st
import pandas as pd
import json, os, base64, requests, re
from io import BytesIO
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image as XLImage
import matplotlib.pyplot as plt

# AI
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.naive_bayes import MultinomialNB
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH


# =============== CONFIG =================
st.set_page_config(page_title="Studio ISA - DrVeto + VetsGo", layout="wide")

GITHUB_FILE_A = "dizionario_drveto.json"
GITHUB_FILE_B = "dizionario_vetsgo.json"
GITHUB_REPO = os.getenv("GITHUB_REPO")
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")

# AI MODELS (global placeholders)
model = None
vectorizer = None
model_B = None
vectorizer_B = None

# =============== UTILS =================
def norm(s):
    return re.sub(r"\s+", " ", str(s).strip().lower())


def any_kw_in(t, keys):
    return any(k in t for k in keys)


def coerce_numeric(s):
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce").fillna(0)
    s = (
        s.astype(str)
        .str.replace(r"\s", "", regex=True)
        .str.replace("€", "", regex=False)
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
    )
    return pd.to_numeric(s, errors="coerce").fillna(0)


def round_pct(values):
    values = pd.to_numeric(values, errors="coerce").fillna(0)
    total = Decimal(str(values.sum()))
    if total == 0:
        return values * 0
    raw = [Decimal(str(v)) * Decimal("100") / total for v in values]
    rounded = [r.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP) for r in raw]
    diff = Decimal("100.00") - sum(rounded)
    rounded[-1] = (rounded[-1] + diff).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return pd.Series([float(x) for x in rounded], index=values.index)


@st.cache_data(ttl=600, show_spinner=False)
def load_excel(file):
    return pd.read_excel(file)


# =============== GITHUB =================
def github_load_json(file_name):
    try:
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{file_name}"
        headers = {"Authorization": f"token {GITHUB_TOKEN}"} if GITHUB_TOKEN else {}
        r = requests.get(url, headers=headers, timeout=12)
        if r.status_code == 200:
            raw = json.loads(base64.b64decode(r.json()["content"]).decode("utf-8"))
            return {norm(k): v for k, v in raw.items()}
    except:
        pass
    return {}


def github_save_json(file_name, data):
    try:
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{file_name}"
        headers = {"Authorization": f"token {GITHUB_TOKEN}"}
        get = requests.get(url, headers=headers, timeout=12)
        sha = get.json().get("sha") if get.status_code == 200 else None

        encoded = base64.b64encode(json.dumps(data, ensure_ascii=False, indent=2).encode()).decode()
        payload = {"message": "Update ISA dictionary", "content": encoded, "branch": "main"}
        if sha:
            payload["sha"] = sha

        requests.put(url, headers=headers, data=json.dumps(payload), timeout=20)
    except:
        pass


# =============== CATEGORY RULES =================
RULES_A = {
    "LABORATORIO": ["analisi","emocromo","test","esame","coprolog","feci","giardia","leishmania","citolog","istolog","urinocolt","urine"],
    "VISITE": ["visita","controllo","consulto","dermatologic"],
    "TERAPIA": ["terapia","terapie"],
    "FAR": ["meloxidyl","konclav","enrox","profenacarp","apoquel","osurnia","cylan","mometa","aristos","cytopoint","milbemax","stomorgyl","previcox"],
    "CHIRURGIA": ["intervento","chirurg","castraz","sterilizz","ovariect","detartrasi","estraz","biopsia","orchiettomia","odontostomat"],
    "DIAGNOSTICA PER IMMAGINI": ["rx","radiograf","eco","ecografia","tac"],
    "MEDICINA": ["flebo","day hospital","trattamento","emedog","cerenia","endovena","pressione"],
    "VACCINI": ["vacc","letifend","rabbia","trivalente","felv"],
    "CHIP": ["microchip","chip"],
    "ALTRE PRESTAZIONI": ["trasporto","eutanasia","unghie","cremazion","otoematoma","pet corner","ricette","medicazione","manualità"]
}

RULES_B = {
    "Visite domiciliari o presso allevamenti": ["domicilio","allevamenti"],
    "Visite ambulatoriali": ["visita","controllo","consulto","terapia","trattam","manual","microchip","vacc","medicazione"],
    "Esami diagnostici per immagine": ["rx","eco","ecogra","tac","raggi","radi"],
    "Altri esami diagnostici": ["analisi","emocromo","prelievo","laboratorio"],
    "Interventi chirurgici": ["chirurg","castraz","ovariect","detartrasi","estraz","anest","endo"],
    "Altre attività": ["acconto"]
}

ORDER_B = list(RULES_B.keys()) + ["Totale"]


# =============== AI TRAIN =================
def train_ai_model(dictionary):
    if not dictionary:
        return None, None
    texts = list(dictionary.keys())
    labels = list(dictionary.values())
    vec = TfidfVectorizer(lowercase=True)
    X = vec.fit_transform(texts)
    m = MultinomialNB()
    m.fit(X, labels)
    return vec, m


# =============== CLASSIFICATION (helpers) =================
def classify_A(desc, fam, mem):
    """Rule-based + memory + (optional) AI suggestion (used only in auto-pass)."""
    global model, vectorizer
    d = norm(desc)

    fam_s = norm(fam)
    if fam_s and fam_s not in {"privato","professionista","nan","none",""}:
        return fam_s.upper()

    for k, v in mem.items():
        if norm(k) in d:
            return v

    # Pure rule fallback
    for cat, keys in RULES_A.items():
        if any_kw_in(d, keys):
            return cat

    return "ALTRE PRESTAZIONI"


def classify_B(prest, mem):
    """Rule-based + memory + (optional) AI suggestion (used only in auto-pass)."""
    d = norm(prest)

    for k, v in mem.items():
        if norm(k) in d:
            return v

    for cat, keys in RULES_B.items():
        if any_kw_in(d, keys):
            return cat

    return "Altre attività"


# =============== SIDEBAR =================
page = st.sidebar.radio("📌 Navigazione", ["Studio ISA", "Dashboard Annuale", "Registro IVA"])
auto_thresh = st.sidebar.slider("Soglia auto-apprendimento (AI)", 0.50, 0.99, 0.85, 0.01)
st.sidebar.caption("Se la confidenza del modello ≥ soglia, il termine viene appreso in automatico.")

# =============== MAIN =================
def main():
    global model, vectorizer, model_B, vectorizer_B

    st.title("📊 Studio ISA – DrVeto + VetsGo")
    file = st.file_uploader("Seleziona Excel", type=["xlsx","xls"])
    if not file:
        st.stop()

    # Load file only once
    if "df" not in st.session_state:
        df = load_excel(file)
        st.session_state.df = df
        mode = "B" if any("prestazioneprodotto" in c.replace(" ","").lower() for c in df.columns) else "A"
        st.session_state.mode = mode
        st.session_state.mem = github_load_json(GITHUB_FILE_A if mode=="A" else GITHUB_FILE_B)
        st.session_state.new = {}
        st.session_state.idx = 0
        st.session_state.auto_added = []  # [(term, cat, conf)]

    df = st.session_state.df.copy()
    mem = st.session_state.mem
    new = st.session_state.new
    mode = st.session_state.mode

    # Train AI
    if mode == "A":
        vectorizer, model = train_ai_model(mem | new)
    else:
        vectorizer_B, model_B = train_ai_model(mem | new)

    # ===== PROCESS A =====
    if mode == "A":
        desc = next(c for c in df.columns if "descrizione" in c.lower())
        fam = next((c for c in df.columns if "famiglia" in c.lower()), None)
        qta = next(c for c in df.columns if "quant" in c.lower() or c.strip()=="%")
        netto = next(c for c in df.columns if "netto" in c.lower() and "dopo" in c.lower())

        df[qta] = coerce_numeric(df[qta])
        df[netto] = coerce_numeric(df[netto])

        base = desc

        # ========== AUTO APPRENDIMENTO PASS ==========
        learned = {norm(k) for k in (mem | new).keys()}
        df["_clean"] = df[base].astype(str).map(norm)
        candidates = sorted([t for t in df["_clean"].unique() if t not in learned])

        auto_added_now = []
        if model and vectorizer and candidates:
            X = vectorizer.transform(candidates)
            probs = model.predict_proba(X)
            preds = model.classes_[probs.argmax(axis=1)]
            confs = probs.max(axis=1)
            for t, p, c in zip(candidates, preds, confs):
                if float(c) >= auto_thresh:
                    new[t] = p
                    auto_added_now.append((t, p, float(c)))

        if auto_added_now:
            st.session_state.new = new
            st.session_state.auto_added.extend(auto_added_now)

        # Classify rows (using classify_A that includes memory & rules)
        df["CategoriaFinale"] = df.apply(lambda r: classify_A(r[desc], r[fam] if fam else None, mem | new), axis=1)

    # ===== PROCESS B =====
    else:
        prest = next(c for c in df.columns if "prestazioneprodotto" in c.replace(" ","").lower())
        imp = next(c for c in df.columns if "totaleimpon" in c.lower())
        iva_col = next((c for c in df.columns if "totaleconiva" in c.replace(" ","").lower()), None)
        tot = next(c for c in df.columns if c.lower().strip()=="totale" or "totale" in c.lower())

        df[imp] = coerce_numeric(df[imp])
        if iva_col:
            df[iva_col] = coerce_numeric(df[iva_col])
        df[tot] = coerce_numeric(df[tot])

        base = prest

        # ========== AUTO APPRENDIMENTO PASS ==========
        learned = {norm(k) for k in (mem | new).keys()}
        df["_clean"] = df[base].astype(str).map(norm)
        candidates = sorted([t for t in df["_clean"].unique() if t not in learned])

        auto_added_now = []
        if model_B and vectorizer_B and candidates:
            X = vectorizer_B.transform(candidates)
            probs = model_B.predict_proba(X)
            preds = model_B.classes_[probs.argmax(axis=1)]
            confs = probs.max(axis=1)
            for t, p, c in zip(candidates, preds, confs):
                if float(c) >= auto_thresh:
                    new[t] = p
                    auto_added_now.append((t, p, float(c)))

        if auto_added_now:
            st.session_state.new = new
            st.session_state.auto_added.extend(auto_added_now)

        # Classify rows (using classify_B that includes memory & rules)
        df["CategoriaFinale"] = df[prest].apply(lambda x: classify_B(x, mem | new))

    # Remove Privato / Professionista
    df = df[~df["CategoriaFinale"].str.lower().isin(["privato","professionista"])]

    # ===== SHOW AUTO-LEARNED THIS RUN =====
    if st.session_state.auto_added:
        with st.expander(f"🤖 Auto-apprendimento: {len(st.session_state.auto_added)} nuovi termini (≥ {auto_thresh:.2f})"):
            auto_df = pd.DataFrame(st.session_state.auto_added, columns=["Termine", "Categoria", "Confidenza"])
            auto_df = auto_df.sort_values("Confidenza", ascending=False)
            st.dataframe(auto_df, use_container_width=True)

    # ===== LEARNING INTERFACE (manuale per ciò che resta) =====
    learned = {norm(k) for k in (mem | new).keys()}
    pending = [t for t in sorted(df["_clean"].unique()) if t not in learned]

    if pending:
        idx = st.session_state.idx
        if idx >= len(pending):
            idx = 0
            st.session_state.idx = 0
        term = pending[idx]
        opts = list(RULES_A.keys()) if mode=="A" else list(RULES_B.keys())
        last = st.session_state.get("last_cat", opts[0])
        default_index = opts.index(last) if last in opts else 0

        st.warning(f"🧠 Da classificare {idx+1}/{len(pending)} → “{term}”")
        cat_sel = st.selectbox("Categoria:", opts, index=default_index)

        if st.button("✅ Salva e prossimo"):
            new[norm(term)] = cat_sel
            st.session_state.new = new
            st.session_state.last_cat = cat_sel

            if idx + 1 >= len(pending):
                # Fine: salva su GitHub
                mem.update(new)
                github_save_json(GITHUB_FILE_A if mode=="A" else GITHUB_FILE_B, mem)
                st.session_state.mem = mem
                st.session_state.new = {}
                st.session_state.idx = 0
                st.success("🎉 Tutto classificato e salvato su GitHub!")
                st.rerun()

            st.session_state.idx = idx + 1
            st.rerun()

        st.stop()

    # ===== REPORT =====
    df = df.drop(columns=["_clean"], errors="ignore")

    if mode == "A":
        # DrVeto: Quantità(%) e Netto (dopo sconto)
        studio = df.groupby("CategoriaFinale").agg({qta:"sum", netto:"sum"}).reset_index()
        studio.columns = ["Categoria","Qtà","Netto"]
        studio["% Qtà"] = round_pct(studio["Qtà"])
        studio["% Netto"] = round_pct(studio["Netto"])
        studio.loc[len(studio)] = ["Totale", studio["Qtà"].sum(), studio["Netto"].sum(), 100, 100]
        ycol = "Netto"
        title = "Somma Netto per Categoria"

    else:
        # VetsGo: Imponibile + ConIVA, % su ConIVA
        studio = df.groupby("CategoriaFinale").agg({imp:"sum", iva_col:"sum"}).reset_index()
        studio.columns = ["Categoria","TotaleImponibile","TotaleConIVA"]
        studio["% Totale"] = round_pct(studio["TotaleConIVA"])
        studio.loc[len(studio)] = ["Totale", studio["TotaleImponibile"].sum(), studio["TotaleConIVA"].sum(), 100]
        studio["Categoria"] = pd.Categorical(studio["Categoria"], categories=ORDER_B, ordered=True)
        studio = studio.sort_values("Categoria")
        ycol = "TotaleConIVA"
        title = "Somma Totale con IVA per Categoria"

    st.dataframe(studio, use_container_width=True)

    fig, ax = plt.subplots(figsize=(8,5))
    ax.bar(studio["Categoria"], studio[ycol], color="steelblue")
    ax.set_title(title)
    plt.xticks(rotation=45, ha="right")
    buf = BytesIO(); plt.tight_layout(); plt.savefig(buf, format="png"); buf.seek(0)
    st.image(buf)

    wb = Workbook(); ws = wb.active; ws.title = "Report"
    for r in dataframe_to_rows(studio, index=False, header=True):
        ws.append(r)
    ws.add_image(XLImage(buf), f"A{len(studio)+3}")
    out = BytesIO(); wb.save(out)

    st.download_button("⬇️ Scarica Excel", data=out.getvalue(), file_name="StudioISA.xlsx")

    # ===== DASHBOARD =====
    if page == "Dashboard Annuale":
        st.header("📈 Dashboard Andamento Annuale")

        date_col = next(c for c in df.columns if "data" in c.replace(" ", "").lower())
        df[date_col] = (
            df[date_col].astype(str)
            .str.extract(r'(\d{1,4}[-/]\d{1,2}[-/]\d{2,4})')[0]
            .apply(lambda x: pd.to_datetime(x, dayfirst=True, errors="coerce"))
        )

        value_col = netto if mode=="A" else tot
        df["Anno"] = df[date_col].dt.year
        df["Mese"] = df[date_col].dt.to_period("M").astype(str)

        anni = sorted(df["Anno"].dropna().unique())
        if not anni:
            st.info("Nessuna data valida trovata per la dashboard.")
            st.stop()

        anno_sel = st.selectbox("Seleziona Anno:", anni, index=len(anni)-1)
        dfY = df[df["Anno"] == anno_sel]

        monthly = dfY.groupby("Mese")[value_col].sum().reset_index()
        all_months = pd.period_range(f"{anno_sel}-01", f"{anno_sel}-12", freq="M").astype(str)
        monthly = monthly.set_index("Mese").reindex(all_months, fill_value=0).reset_index().rename(columns={"index":"Mese"})

        st.subheader("Trend Fatturato Mensile")
        st.line_chart(monthly.set_index("Mese"))

        catshare = dfY.groupby("CategoriaFinale")[value_col].sum().reset_index()
        catshare["%"] = round_pct(catshare[value_col])
        st.subheader("Ripartizione per Categoria")
        st.bar_chart(catshare.set_index("CategoriaFinale")["%"])

        area = dfY.groupby(["Mese", "CategoriaFinale"])[value_col].sum().reset_index()
        area = area.pivot(index="Mese", columns="CategoriaFinale", values=value_col).fillna(0)
        st.subheader("Andamento Categorie nel Tempo")
        st.area_chart(area)

    # =============== REGISTRO IVA ===========
# --- HELPER: coercizione numerica robusta ---
if page == "Registro IVA":
    render_registro_iva()
def _coerce_numeric_series(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce").fillna(0)
    s = (s.astype(str)
         .str.replace(r"\s", "", regex=True)
         .str.replace("€", "", regex=False)
         .str.replace(".", "", regex=False)
         .str.replace(",", ".", regex=False))
    return pd.to_numeric(s, errors="coerce").fillna(0)

# --- HELPER: trova colonna per nome “fuzzy” ---
def _find_col(df: pd.DataFrame, *needles, required=True):
    cols = list(df.columns)
    low = [c.lower().replace(" ", "") for c in cols]
    for i, c in enumerate(low):
        if all(n.lower().replace(" ", "") in c for n in needles):
            return cols[i]
    if required:
        raise ValueError(f"Colonna non trovata: {needles}")
    return None

# --- HELPER: costruzione DOCX orizzontale in memoria ---
def _build_registro_iva_docx(df: pd.DataFrame, header: dict) -> bytes:
    doc = Document()

    # A4 landscape + margini
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    # A4: 29.7 x 21.0 cm → in landscape width>height
    section.page_width, section.page_height = Cm(29.7), Cm(21.0)
    # Margini
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)

    # Intestazione struttura
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(header.get("denominazione", ""))
    r.bold = True; r.font.size = Pt(14)
    if header.get("indirizzo"):
        doc.add_paragraph(header["indirizzo"]).alignment = WD_ALIGN_PARAGRAPH.CENTER
    riga2 = []
    if header.get("cap"): riga2.append(header["cap"])
    if header.get("citta"): riga2.append(header["citta"])
    if header.get("provincia"): riga2.append(f"({header['provincia']})")
    if riga2:
        doc.add_paragraph(" ".join(riga2)).alignment = WD_ALIGN_PARAGRAPH.CENTER
    riga3 = []
    if header.get("piva"): riga3.append(f"P.IVA {header['piva']}")
    if header.get("cf"): riga3.append(f"CF {header['cf']}")
    if riga3:
        doc.add_paragraph(" • ".join(riga3)).alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Spazio
    doc.add_paragraph("")

    # Titolo registro
    tp = doc.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    tr = tp.add_run("Registro IVA – Vendite")
    tr.bold = True; tr.font.size = Pt(12)

    # Tabella principale
    cols_order = [
        "Data", "Numero", "Cliente", "P.Iva", "Codice Fiscale",
        "Indirizzo", "CAP", "Città", "Totale Netto",
        "Totale ENPAV", "Totale Imponibile", "Totale IVA",
        "Totale Sconto", "Rit. d'acconto", "Totale"
    ]
    table = doc.add_table(rows=1, cols=len(cols_order))
    table.style = "Table Grid"
    hdr_cells = table.rows[0].cells
    for j, h in enumerate(cols_order):
        run = hdr_cells[j].paragraphs[0].add_run(h)
        run.bold = True

    # Righe tabella
    for _, row in df[cols_order].iterrows():
        cells = table.add_row().cells
        for j, h in enumerate(cols_order):
            val = row[h]
            if isinstance(val, float) or isinstance(val, int):
                text = f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            else:
                text = "" if pd.isna(val) else str(val)
            cells[j].paragraphs[0].add_run(text)

    # --- Pagina totali ---
    doc.add_page_break()

    doc.add_paragraph().add_run("Riepilogo Totali").bold = True

    # Tabella totali generali
    tot_table = doc.add_table(rows=1, cols=7)
    tot_table.style = "Table Grid"
    tot_hdr = ["Totale Netto", "Totale ENPAV", "Totale Imponibile", "Totale IVA",
               "Totale Sconto", "Rit. d'acconto", "Totale"]
    for j, h in enumerate(tot_hdr):
        r = tot_table.rows[0].cells[j].paragraphs[0].add_run(h); r.bold = True

    # Somme generali
    def _sum(col): return float(pd.to_numeric(df[col], errors="coerce").fillna(0).sum())
    totals = [
        _sum("Totale Netto"),
        _sum("Totale ENPAV"),
        _sum("Totale Imponibile"),
        _sum("Totale IVA"),
        _sum("Totale Sconto"),
        _sum("Rit. d'acconto"),
        _sum("Totale"),
    ]
    row = tot_table.add_row().cells
    for j, v in enumerate(totals):
        row[j].paragraphs[0].add_run(
            f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        )

    # Sezione IVA 22%
    doc.add_paragraph("")
    doc.add_paragraph().add_run("Riepilogo IVA 22%").bold = True

    iva22_table = doc.add_table(rows=1, cols=4)
    iva22_table.style = "Table Grid"
    hdr_22 = ["Tot. Netto (22%)", "Tot. ENPAV (22%)", "Tot. Imponibile (22%)", "Importo IVA (22%)"]
    for j, h in enumerate(hdr_22):
        r = iva22_table.rows[0].cells[j].paragraphs[0].add_run(h); r.bold = True

    # Calcolo 22%: prova a filtrare per colonna aliquota; se non c'è, usa tutte le righe
    aliquota_col = None
    for c in df.columns:
        cl = c.lower()
        if ("%iva" in cl) or ("aliquota" in cl) or ("aliq" in cl):
            aliquota_col = c; break

    if aliquota_col is not None:
        aliq = pd.to_numeric(pd.to_numeric(df[aliquota_col], errors="coerce").fillna(0))
        mask22 = aliq.eq(22) | aliq.eq(22.0)
        df22 = df[mask22]
        note_22 = ""
    else:
        df22 = df
        note_22 = "⚠️ Aliquota non trovata: considerati tutti i movimenti come 22%."

    def _sum22(col): return float(pd.to_numeric(df22[col], errors="coerce").fillna(0).sum())
    vals22 = [
        _sum22("Totale Netto"),
        _sum22("Totale ENPAV"),
        _sum22("Totale Imponibile"),
        _sum22("Totale IVA")
    ]
    row22 = iva22_table.add_row().cells
    for j, v in enumerate(vals22):
        row22[j].paragraphs[0].add_run(
            f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        )

    if note_22:
        doc.add_paragraph(note_22)

    # Esporta
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()

# --- RENDER DELLA PAGINA "Registro IVA" ---
def render_registro_iva():
    st.header("📄 Registro IVA (Word)")

    # Dati intestazione struttura
    st.subheader("Intestazione")
    c1, c2, c3 = st.columns([1.2,1,1])
    with c1:
        denominazione = st.text_input("Denominazione", placeholder="Clinica Veterinaria ...")
        indirizzo = st.text_input("Indirizzo", placeholder="Via/Piazza ...")
    with c2:
        cap = st.text_input("CAP", placeholder="00000")
        citta = st.text_input("Città", placeholder="Città")
    with c3:
        provincia = st.text_input("Provincia (sigla)", placeholder="MI")
        piva = st.text_input("Partita IVA", placeholder="IT........")
        cf = st.text_input("Codice Fiscale", placeholder="...")

    uploaded = st.file_uploader("📁 Seleziona Excel (Registro IVA)", type=["xlsx","xls"])
    if not uploaded:
        st.info("Carica il file Excel del Registro IVA.")
        return

    df_raw = pd.read_excel(uploaded)

    # Normalizzazione nomi colonne e coercizione numerica sui totali
    try:
        col_data   = _find_col(df_raw, "data")
        col_num    = _find_col(df_raw, "numero")
        col_cli    = _find_col(df_raw, "cliente")
        col_piva   = _find_col(df_raw, "p.iva", required=False) or _find_col(df_raw, "piva", required=False) or "P.Iva"
        col_cf     = _find_col(df_raw, "codice", "fiscale", required=False) or "Codice Fiscale"
        col_addr   = _find_col(df_raw, "indirizzo", required=False) or "Indirizzo"
        col_cap    = _find_col(df_raw, "cap", required=False) or "CAP"
        col_citta  = _find_col(df_raw, "citt", required=False) or "Città"

        col_tnetto = _find_col(df_raw, "totale", "netto")
        col_enpav  = _find_col(df_raw, "totale", "enpav")
        col_imp    = _find_col(df_raw, "totale", "imponibile")
        col_iva    = _find_col(df_raw, "totale", "iva")
        col_sconto = _find_col(df_raw, "totale", "sconto")
        col_rit    = _find_col(df_raw, "rit", "acconto")
        col_tot    = _find_col(df_raw, "totale")

    except ValueError as e:
        st.error(f"Colonne mancanti: {e}")
        return

    df = pd.DataFrame({
        "Data": df_raw[col_data],
        "Numero": df_raw[col_num],
        "Cliente": df_raw[col_cli],
        "P.Iva": df_raw.get(col_piva, ""),
        "Codice Fiscale": df_raw.get(col_cf, ""),
        "Indirizzo": df_raw.get(col_addr, ""),
        "CAP": df_raw.get(col_cap, ""),
        "Città": df_raw.get(col_citta, ""),
        "Totale Netto": _coerce_numeric_series(df_raw[col_tnetto]),
        "Totale ENPAV": _coerce_numeric_series(df_raw[col_enpav]),
        "Totale Imponibile": _coerce_numeric_series(df_raw[col_imp]),
        "Totale IVA": _coerce_numeric_series(df_raw[col_iva]),
        "Totale Sconto": _coerce_numeric_series(df_raw[col_sconto]),
        "Rit. d'acconto": _coerce_numeric_series(df_raw[col_rit]),
        "Totale": _coerce_numeric_series(df_raw[col_tot]),
    })

    # Anteprima
    st.subheader("Anteprima")
    st.dataframe(df.head(30), use_container_width=True)

    # Genera DOCX
    if st.button("🖨️ Genera Word (Registro IVA)"):
        header = {
            "denominazione": denominazione,
            "indirizzo": indirizzo,
            "cap": cap,
            "citta": citta,
            "provincia": provincia,
            "piva": piva,
            "cf": cf,
        }
        docx_bytes = _build_registro_iva_docx(df, header)
        anno = pd.to_datetime(df["Data"], errors="coerce").dt.year.dropna()
        year_str = str(int(anno.mode()[0])) if not anno.empty else str(datetime.now().year)
        st.download_button(
            "⬇️ Scarica Registro IVA (DOCX)",
            data=docx_bytes,
            file_name=f"Registro_IVA_{year_str}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )


if __name__ == "__main__":
    main()


