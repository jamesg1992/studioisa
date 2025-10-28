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
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT, WD_SECTION_START
from docx.oxml.ns import qn


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
    if page == "Registro IVA":
        render_registro_iva()
        st.stop()
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
def render_registro_iva():
    st.header("📄 Registro IVA - Vendite")

    c1, c2 = st.columns(2)
    with c1:
        struttura = st.text_input("Nome Struttura")
        indirizzo = st.text_input("Indirizzo (Via, CAP, Città)")
        piva = st.text_input("Partita IVA")
    with c2:
        note_header = st.text_input("Nota intestazione destra (opzionale)", "")

    file = st.file_uploader("Carica il file Registro IVA (Excel)", type=["xlsx"])
    if not file or not struttura:
        return

    # --- Lettura e normalizzazione ---
    df = pd.read_excel(file)

    # Nomi colonne attesi dal tracciato
    expected_cols = [
        "Data", "Numero", "Cliente", "P.Iva", "Codice Fiscale",
        "Indirizzo", "CAP", "Città", "Totale Netto", "Totale ENPAV",
        "Totale Imponibile", "Totale IVA", "Totale Sconto",
        "Rit. d'acconto", "Totale"
    ]
    # Manteniamo l'ordine/filtriamo all'occorrenza (se mancano alcune colonne, gestiamo graceful)
    cols_presenti = [c for c in expected_cols if c in df.columns]
    df = df[cols_presenti].copy()

    # Coercizione data (dd/mm/yyyy hh:mm:ss o simili)
    if "Data" not in df.columns:
        st.error("❌ Colonna 'Data' non trovata nel file.")
        return

    df["Data"] = (
        df["Data"].astype(str)
        .str.extract(r"(\d{1,4}[-/]\d{1,2}[-/]\d{2,4})")[0]
        .apply(lambda x: pd.to_datetime(x, dayfirst=True, errors="coerce"))
    )

    if df["Data"].notna().sum() == 0:
        st.error("❌ Nessuna data valida riconosciuta nella colonna 'Data'.")
        return

    data_min = df["Data"].min().strftime("%d/%m/%Y")
    data_max = df["Data"].max().strftime("%d/%m/%Y")
    anno = int(df["Data"].dt.year.dropna().mode()[0])

    # Coercizione importi dove presenti
    ren_num = {
        "Totale Netto": "tot_netto",
        "Totale ENPAV": "tot_enpav",
        "Totale Imponibile": "tot_imponibile",
        "Totale IVA": "tot_iva",
        "Totale Sconto": "tot_sconto",
        "Rit. d'acconto": "tot_rit",
        "Totale": "totale"
    }
    df = df.rename(columns={c: ren_num.get(c, c) for c in df.columns})
    for c in ren_num.values():
        if c in df.columns:
            df[c] = coerce_numeric(df[c])

    # --- Crea documento Word in orizzontale ---
    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    # A4 landscape
    section.page_width = Inches(11.69)
    section.page_height = Inches(8.27)
    section.left_margin = Inches(0.4)
    section.right_margin = Inches(0.4)
    section.top_margin = Inches(0.4)
    section.bottom_margin = Inches(0.4)

    # Font base Aptos Narrow 8 (fallback a Calibri Narrow se non disponibile)
    style = doc.styles["Normal"]
    style.font.name = "Aptos Narrow"
    try:
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Aptos Narrow")
    except Exception:
        pass
    style.font.size = Pt(8)

    # --- Header: tabella 2 colonne (sx/dx) ---
    header = section.header
    hdr_table = header.add_table(rows=1, cols=2, width=Inches(11.4))
    hdr_table.autofit = False
    hdr_left, hdr_right = hdr_table.rows[0].cells

    # Sinistra
    pL = cell_left.paraghaphs[0]
    pL.alignment = WD_ALIGN_PARAGRAPH.LEFT

    run1 = pL.add_run(struttura + "\n")
    run1.font.name = "Segoe UI"
    run1.font.size = Pt(14)
    run1.bold = False

    run2 = pL.add_run(indirizzo + "\n")
    run2.font.name = "Segoe UI"
    run2.font.size = Pt(12)
    run2.bold = False

    run3 = pL.add_run(f"P.IVA {piva}")
    run3.font.name = "Aptos Narrow"
    run3._element.rPr.rFonts.set(qn('w:eastAsia'), "Aptos Narrow")
    run3.font.size = Pt(10)
    run3.bold = True

    # Destra (ANNO, intervallo date, Pag. X di Y)
    pR = cell_right.paragraphs[0]
    pR.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    run4 = pR.add_run(f"ANNO {anno}\n")
    run4.font.name = "Calibri"
    run4.font.size = Pt(10)
    run4.bold = False

    run5 = pR.add_run(f"Entrate dal {data_min} al {data_max}")
    run5.font.name = "Aptos Narrow"
    run5._element.rPr.rFonts.set(qn('w:eastAsia'), "Aptos Narrow")
    run5.font.size = Pt(10)
    run5.bold = True

    # Nuova riga con "Pag. X di Y"
    pR2 = header.add_paragraph()
    pR2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _ = pR2.add_run("Pag. ")
    # Campo PAGE
    from docx.oxml import OxmlElement
    from docx.oxml.ns import nsmap

    def _add_field(paragraph, instr_text):
        """Inserisce un campo Word (es. PAGE, NUMPAGES)."""
        r = paragraph.add_run()
        fldChar1 = OxmlElement("w:fldChar")
        fldChar1.set(qn("w:fldCharType"), "begin")
        r._r.append(fldChar1)

        instr = OxmlElement("w:instrText")
        instr.set(qn("xml:space"), "preserve")
        instr.text = instr_text
        r._r.append(instr)

        fldChar2 = OxmlElement("w:fldChar")
        fldChar2.set(qn("w:fldCharType"), "separate")
        r._r.append(fldChar2)

        t = OxmlElement("w:t")
        t.text = "1"
        r._r.append(t)

        fldChar3 = OxmlElement("w:fldChar")
        fldChar3.set(qn("w:fldCharType"), "end")
        r._r.append(fldChar3)

    _add_field(pR2, " PAGE ")
    _ = pR2.add_run(" di ")
    _add_field(pR2, " NUMPAGES ")

    # --- Tabella dati (tutte le colonne presenti) ---
    # Ripristina intestazioni “pulite” come le vede l’utente
    inv_ren = {v: k for k, v in ren_num.items()}
    vis_cols = [inv_ren.get(c, c) for c in df.columns]  # nomi come in input

    tbl = doc.add_table(rows=1, cols=len(df.columns))
    tbl.style = "Table Grid"
    tbl.autofit = True

    # Header row
    for j, col in enumerate(vis_cols):
        cell = tbl.cell(0, j)
        cell.text = str(col)

    # Ripeti header su ogni pagina
    try:
        tr = tbl.rows[0]._tr
        trPr = tr.get_or_add_trPr()
        tblHeader = OxmlElement('w:tblHeader')
        trPr.append(tblHeader)
    except Exception:
        pass

    # Righe dati
    for _, row in df.iterrows():
        cells = tbl.add_row().cells
        for j, col in enumerate(df.columns):
            val = row[col]
            if pd.api.types.is_number(val):
                cells[j].text = f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            elif isinstance(val, pd.Timestamp):
                cells[j].text = val.strftime("%d/%m/%Y")
            else:
                cells[j].text = "" if pd.isna(val) else str(val)

    # --- Totali finali (ultima pagina) ---
    def fmt(n):
        return f"{n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    doc.add_paragraph()  # spazio
    pTot = doc.add_paragraph("Totali Finali:", style="Normal")
    pTot.runs[0].font.bold = True

    tot_netto = fmt(df["tot_netto"].sum()) if "tot_netto" in df.columns else "0,00"
    tot_enpav = fmt(df["tot_enpav"].sum()) if "tot_enpav" in df.columns else "0,00"
    tot_imponibile = fmt(df["tot_imponibile"].sum()) if "tot_imponibile" in df.columns else "0,00"
    tot_iva = fmt(df["tot_iva"].sum()) if "tot_iva" in df.columns else "0,00"
    tot_sconto = fmt(df["tot_sconto"].sum()) if "tot_sconto" in df.columns else "0,00"
    tot_rit = fmt(df["tot_rit"].sum()) if "tot_rit" in df.columns else "0,00"
    tot_totale = fmt(df["totale"].sum()) if "totale" in df.columns else "0,00"

    doc.add_paragraph(f"• Totale Netto (IVA 22%): {tot_netto} €")
    doc.add_paragraph(f"• Totale ENPAV (IVA 22%): {tot_enpav} €")
    doc.add_paragraph(f"• Totale Imponibile (IVA 22%): {tot_imponibile} €")
    doc.add_paragraph(f"• Importo IVA (22%): {tot_iva} €")
    doc.add_paragraph(f"• Totale Sconto: {tot_sconto} €")
    doc.add_paragraph(f"• Ritenuta d'acconto: {tot_rit} €")
    doc.add_paragraph(f"• Totale complessivo: {tot_totale} €")

    # --- Esporta Word ---
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    st.download_button("⬇️ Scarica Registro IVA (Word)", buf, file_name=f"Registro_IVA_{anno}.docx")


if __name__ == "__main__":
    main()






