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


# =============== CLASSIFICATION =================
def classify_A(desc, fam, mem):
    global model, vectorizer
    d = norm(desc)

    fam_s = norm(fam)
    if fam_s and fam_s not in {"privato","professionista","nan","none",""}:
        return fam_s.upper()

    for k,v in mem.items():
        if norm(k) in d:
            return v

    if model and vectorizer:
        X = vectorizer.transform([d])
        pred = model.predict(X)[0]
        conf = model.predict_proba(X).max()
        if conf >= 0.75:
            return pred

    for cat, keys in RULES_A.items():
        if any_kw_in(d, keys):
            return cat

    return "ALTRE PRESTAZIONI"


def classify_B(prest, mem):
    global model_B, vectorizer_B
    d = norm(prest)

    for k,v in mem.items():
        if norm(k) in d:
            return v

    if model_B and vectorizer_B:
        X = vectorizer_B.transform([d])
        pred = model_B.predict(X)[0]
        conf = model_B.predict_proba(X).max()
        if conf >= 0.75:
            return pred

    for cat, keys in RULES_B.items():
        if any_kw_in(d, keys):
            return cat

    return "Altre attività"


# =============== MAIN =================
page = st.sidebar.radio("📌 Navigazione", ["Studio ISA", "Dashboard Annuale"])


def main():

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

    df = st.session_state.df.copy()
    mem = st.session_state.mem
    new = st.session_state.new
    mode = st.session_state.mode

    # Train AI
    global model, vectorizer, model_B, vectorizer_B
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
        df["CategoriaFinale"] = df.apply(lambda r: classify_A(r[desc], r[fam] if fam else None, mem|new), axis=1)


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
        df["CategoriaFinale"] = df[prest].apply(lambda x: classify_B(x, mem|new))


    # Remove Privato / Professionista
    df = df[~df["CategoriaFinale"].str.lower().isin(["privato","professionista"])]


    # ===== LEARNING INTERFACE =====
    learned = {norm(k) for k in (mem|new).keys()}
    df["_clean"] = df[base].astype(str).map(norm)
    pending = [t for t in sorted(df["_clean"].unique()) if t not in learned]

    if pending:
        idx = st.session_state.idx
        term = pending[idx]
        opts = list(RULES_A.keys()) if mode=="A" else list(RULES_B.keys())
        last = st.session_state.get("last_cat", opts[0])

        st.warning(f"🧠 Da classificare {idx+1}/{len(pending)} → “{term}”")
        cat_sel = st.selectbox("Categoria:", opts, index=opts.index(last) if last in opts else 0)

        if st.button("✅ Salva e prossimo"):
            new[norm(term)] = cat_sel
            st.session_state.last_cat = cat_sel

            if idx + 1 >= len(pending):
                mem.update(new)
                github_save_json(GITHUB_FILE_A if mode=="A" else GITHUB_FILE_B, mem)
                st.success("🎉 Salvataggio completato!")
                st.session_state.idx = 0
                st.session_state.new = {}
                st.stop()

            st.session_state.new = new
            st.session_state.idx += 1
            st.rerun()

        st.stop()


    # ===== REPORT =====
    if mode == "A":
        studio = df.groupby("CategoriaFinale").agg({qta:"sum", netto:"sum"}).reset_index()
        studio.columns = ["Categoria","Qtà","Netto"]
        studio["% Qtà"] = round_pct(studio["Qtà"])
        studio["% Netto"] = round_pct(studio["Netto"])
        studio.loc[len(studio)] = ["Totale", studio["Qtà"].sum(), studio["Netto"].sum(), 100, 100]
        ycol = "Netto"
        title = "Somma Netto per Categoria"

    else:
        studio = df.groupby("CategoriaFinale").agg({imp:"sum", iva_col:"sum"}).reset_index()
        studio.columns = ["Categoria","TotaleImponibile","TotaleConIVA"]
        studio["% Totale"] = round_pct(studio["TotaleConIVA"])
        studio.loc[len(studio)] = ["Totale", studio["TotaleImponibile"].sum(), studio["TotaleConIVA"].sum(), 100]
        studio["Categoria"] = pd.Categorical(studio["Categoria"], categories=ORDER_B, ordered=True)
        studio = studio.sort_values("Categoria")
        ycol = "TotaleConIVA"
        title = "Somma Totale con IVA per Categoria"

    st.dataframe(studio)

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
        anno_sel = st.selectbox("Seleziona Anno:", anni, index=len(anni)-1)
        dfY = df[df["Anno"] == anno_sel]

        monthly = dfY.groupby("Mese")[value_col].sum().reset_index()
        st.line_chart(monthly.set_index("Mese"))

        catshare = dfY.groupby("CategoriaFinale")[value_col].sum().reset_index()
        catshare["%"] = round_pct(catshare[value_col])
        st.bar_chart(catshare.set_index("CategoriaFinale")["%"])

        area = dfY.groupby(["Mese", "CategoriaFinale"])[value_col].sum().reset_index()
        area = area.pivot(index="Mese", columns="CategoriaFinale", values=value_col).fillna(0)
        st.area_chart(area)


if __name__ == "__main__":
    main()
