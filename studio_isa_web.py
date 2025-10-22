import streamlit as st
import pandas as pd
import json, os, base64, requests, re
from io import BytesIO
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill
import matplotlib.pyplot as plt

# ───────────────── CONFIG ─────────────────
st.set_page_config(page_title="Studio ISA v7", layout="wide")
GITHUB_FILE = os.getenv("GITHUB_FILE", "keywords_memory.json")
GITHUB_REPO = os.getenv("GITHUB_REPO")  # es: "utente/repo"
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")  # personal access token

# Categorie/keywords predefinite (Tipo A)
RULES_A = {
    "LABORATORIO": ["analisi","emocromo","test","esame","coprolog","feci","giardia","leishmania","citolog","istolog","urinocolt","urine"],
    "VISITE": ["visita","controllo","consulto","dermatologic"],
    "FAR": ["meloxidyl","konclav","enrox","profenacarp","apoquel","osurnia","cylan","mometa","aristos","cytopoint","milbemax","stomorgyl","previcox","royal","stronghold","nexgard","procox"],
    "CHIRURGIA": ["intervento","chirurg","castraz","sterilizz","ovariect","detartrasi","estraz","biopsia","orchiettomia","odontostomat"],
    "DIAGNOSTICA PER IMMAGINI": ["rx","radiograf","eco","ecografia","tac"],
    "MEDICINA": ["terapia","terapie","flebo","day hospital","trattamento","emedog","cerenia","endovena","pressione"],
    "VACCINI": ["vacc","letifend","rabbia","trivalente","felv"],
    "CHIP": ["microchip","chip"],
    "ALTRE PRESTAZIONI": ["trasporto","eutanasia","unghie","cremazion","otoematoma","pet corner","ricette","medicazione","manualità"]
}

# Categorie/keywords specifiche (Tipo B)
RULES_B = {
    "Domicilio": ["visite domiciliari","allevamenti","domicilio"],
    "Terapia": ["visite ambulatoriali","terapia","trattamenti","vaccinazioni","ambulatorio","manualità","pet corner","visite","ricette","medicazione","microchip","controllo"],
    "Radiologia": ["esami diagnostici per immagine","radiologia","eco","ecografia","tac","rx","raggi"],
    "Laboratorio": ["altri esami diagnostici","esami biochimici","laboratorio","malattie infettive","emocromo","prelievo"],
    "Chirurgia": ["interventi chirurgici","avulsione","endoscopia","eutanasia","sedazione","anestesia","chirurgia","odontostomat","orchiettomia","asportazione","biopsia","ovariectomia"],
    "Ostetricia": ["assistenza al parto","ostetricia","parto"],
    "Consulenza": ["attività di consulenza","perizia","collaborazione","telemedicina","consulto"],
    "Inseminazione": ["inseminazione artificiale"],
    "Altre attività": ["acconto"]
}

# ──────────────── GITHUB I/O ────────────────
def github_load_json():
    """Legge il JSON dal repo GitHub, se configurato; altrimenti {}."""
    try:
        if not (GITHUB_REPO and GITHUB_FILE):
            return {}
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_FILE}"
        headers = {"Authorization": f"token {GITHUB_TOKEN}"} if GITHUB_TOKEN else {}
        r = requests.get(url, headers=headers, timeout=15)
        if r.status_code == 200 and "content" in r.json():
            content = base64.b64decode(r.json()["content"]).decode("utf-8")
            return json.loads(content)
        return {}
    except Exception:
        return {}

def github_save_json(data: dict):
    """Scrive/aggiorna il JSON nel repo GitHub, se configurato; no-op se non configurato."""
    try:
        if not (GITHUB_REPO and GITHUB_FILE and GITHUB_TOKEN):
            st.info("ℹ️ Salvataggio GitHub non configurato (manca token o repo).")
            return
        url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_FILE}"
        headers = {"Authorization": f"token {GITHUB_TOKEN}"}
        get_res = requests.get(url, headers=headers, timeout=15)
        sha = get_res.json().get("sha") if get_res.status_code == 200 else None

        encoded = base64.b64encode(json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")).decode("utf-8")
        payload = {"message": "Aggiornamento dizionario Studio ISA", "content": encoded, "branch": "main"}
        if sha: payload["sha"] = sha

        put_res = requests.put(url, headers=headers, data=json.dumps(payload), timeout=15)
        if put_res.status_code in (200, 201):
            st.success("✅ Dizionario aggiornato su GitHub!")
        else:
            st.error(f"❌ Errore GitHub: {put_res.status_code}")
    except Exception as e:
        st.error(f"❌ Salvataggio GitHub fallito: {e}")

# ──────────────── UTILS ────────────────
@st.cache_data(show_spinner=False)
def load_excel(file):
    return pd.read_excel(file)

def norm(s: str) -> str:
    return re.sub(r"\s+", " ", str(s).strip().lower())

def any_kw_in(text: str, kws: list[str]) -> bool:
    t = norm(text)
    return any(norm(k) in t for k in kws)

def detect_file_type(columns: list[str]) -> str:
    """Ritorna 'A' (DrVeto) o 'B' (VetsGo) in base alle colonne presenti."""
    cols_low = [c.lower() for c in columns]
    # Tipo B se ha PrestazioneProdotto + Totali
    if any("prestazioneprodotto" in c.replace(" ", "") for c in cols_low) and \
       any("totaleimpon" in c for c in cols_low) and \
       any("totaleconiva" in c.replace(" ", "") for c in cols_low) and \
       any(re.fullmatch(r".*\btotale\b.*", c) for c in cols_low):
        return "B"
    # Tipo A se ha 'descrizione' e colonne classiche
    if any("descrizione" in c for c in cols_low) and \
       any(c.strip() == "%" for c in columns) and \
       any("netto" in c and "dopo" in c for c in cols_low):
        return "A"
    # fallback: prova a inferire
    if any("prestazioneprodotto" in c.replace(" ", "") for c in cols_low):
        return "B"
    return "A"

def pick_column(df: pd.DataFrame, *candidates: str) -> str | None:
    """Trova la colonna per nome (case-insensitive, ignora spazi/simboli)."""
    def key(s): return re.sub(r"[^\w]", "", s).lower()
    cols = {key(c): c for c in df.columns}
    for cand in candidates:
        k = key(cand)
        for ck, orig in cols.items():
            if k in ck:
                return orig
    return None

# ─────────────── Classificatori ───────────────
def classify_A(desc: str, fam_val: str | None, memory: dict) -> str:
    """Tipo A: usa FamigliaCategoria se presente, poi memory, poi RULES_A."""
    if pd.notna(fam_val) and str(fam_val).strip():
        return str(fam_val).strip()
    d = norm(desc)
    if not d:
        return "ALTRE PRESTAZIONI"
    for key, cat in memory.items():
        if norm(key) in d:
            return cat
    for cat, kws in RULES_A.items():
        if any_kw_in(d, kws):
            return cat
    return "ALTRE PRESTAZIONI"

def classify_B(prest: str, cat_val: str | None, memory: dict) -> str:
    """Tipo B: usa Categoria se presente, poi memory, poi RULES_B."""
    if pd.notna(cat_val) and str(cat_val).strip():
        return str(cat_val).strip()
    d = norm(prest)
    if not d:
        return "Altre attività"
    for key, cat in memory.items():
        if norm(key) in d:
            return cat
    for cat, kws in RULES_B.items():
        if any_kw_in(d, kws):
            return cat
    return "Altre attività"

# ──────────────── MAIN ────────────────
def main():
    st.title("📊 Studio ISA DrVeto - VetsGo")

    up = st.file_uploader("📁 Seleziona file Excel", type=["xlsx","xls"])
    if not up:
        st.info("Carica un file per iniziare.")
        return

    # Ricarica sempre il dizionario più recente
    latest_memory = github_load_json()

    # Inizializza sessione per nuovo file
    if "df" not in st.session_state or up.name != st.session_state.get("last_file"):
        st.session_state.df = load_excel(up)
        st.session_state.user_memory = latest_memory
        st.session_state.local_updates = {}
        st.session_state.pending_terms = []
        st.session_state.idx = 0
        st.session_state.last_file = up.name
    else:
        st.session_state.user_memory = latest_memory

    df = st.session_state.df.copy()
    user_memory = st.session_state.user_memory

    # Identifica tipo file
    ftype = detect_file_type(list(df.columns))
    st.caption(f"🔎 Tipo file rilevato: **{('A – DrVeto' if ftype=='A' else 'B – Gestionale nuovo')}**")

    # ──────────────── TIPO A ────────────────
    if ftype == "A":
        col_desc = pick_column(df, "Descrizione nel documento", "Descrizione (da archivio DrVeto)", "Descrizione")
        col_fam  = pick_column(df, "FamigliaCategoria", "Famiglia prestazione / categoria prodotto", "Famiglia")
        col_netto= next((c for c in df.columns if ("netto" in c.lower() and "dopo" in c.lower())), None)
        col_perc = next((c for c in df.columns if c.strip() == "%"), None)

        if not all([col_desc, col_fam, col_netto, col_perc]):
            st.error("❌ Colonne richieste per il Tipo A non trovate.")
            st.stop()

        # Classifica
        df["FamigliaCategoria"] = df.apply(lambda r: classify_A(r[col_desc], r[col_fam], user_memory), axis=1)

        # Apprendimento: termini non noti
        uniq_terms = sorted({str(v).strip() for v in df[col_desc].dropna().unique()}, key=lambda s: s.casefold())
        pending = []
        for term in uniq_terms:
            t = norm(term)
            if not any(norm(k) in t for k in user_memory.keys()) and \
               not any(any_kw_in(t, kws) for kws in RULES_A.values()):
                pending.append(term)

        # UI apprendimento
        if pending:
            if "idx" not in st.session_state: st.session_state.idx = 0
            if st.session_state.idx >= len(pending): st.session_state.idx = 0
            term = pending[st.session_state.idx]
            st.warning(f"🧠 Nuovo termine (Tipo A): {term} [{st.session_state.idx+1}/{len(pending)}]")

            default_opts = list(RULES_A.keys())
            c1, c2 = st.columns([2,1])
            with c1:
                cat_sel = st.selectbox("Seleziona categoria:", default_opts, key=f"a_sel_{st.session_state.idx}")
                cat_free = st.text_input("…oppure categoria personalizzata (opzionale)", key=f"a_free_{st.session_state.idx}")
                final_cat = cat_free.strip() if cat_free.strip() else cat_sel
            with c2:
                if st.button("✅ Salva e continua"):
                    user_memory[term] = final_cat
                    github_save_json(user_memory)
                    st.session_state.user_memory = user_memory
                    st.session_state.idx += 1
                    st.rerun()
            st.stop()

        # Report (Tipo A): FamigliaCategoria / Qtà / Netto / % Qtà / % Netto
        st.success("✅ Nessun nuovo termine. Genero il report (DrVeto)…")

        studio_isa = (
            df.groupby("FamigliaCategoria", dropna=False)
              .agg({col_perc:"sum", col_netto:"sum"})
              .reset_index()
              .rename(columns={col_perc:"Qtà", col_netto:"Netto"})
        )
        # opzionale: escludi righe sgradite
        studio_isa = studio_isa[~studio_isa["FamigliaCategoria"].str.lower().isin(["privato","none"])]

        tot_qta = studio_isa["Qtà"].sum()
        tot_netto = studio_isa["Netto"].sum()
        studio_isa["% Qtà"] = (studio_isa["Qtà"]/tot_qta*100).round(2) if tot_qta else 0
        studio_isa["% Netto"] = (studio_isa["Netto"]/tot_netto*100).round(2) if tot_netto else 0
        studio_isa = pd.concat([studio_isa, pd.DataFrame([["Totale", tot_qta, tot_netto, 100, 100]], columns=studio_isa.columns)], ignore_index=True)

        st.subheader("📄 Tabella Studio ISA (DrVeto)")
        st.dataframe(
            studio_isa.style
              .apply(lambda r: ['background-color: #fff8b3' if r["FamigliaCategoria"]=="Totale" else '' for _ in r], axis=1)
              .set_properties(subset=["FamigliaCategoria"], **{"font-weight":"bold"})
              .format({"Qtà":"{:,.0f}", "Netto":"{:,.2f}", "% Qtà":"{:.2f}", "% Netto":"{:.2f}"})
        )

        # Grafico: Somma Netto
        st.subheader("📊 Somma Netto per FamigliaCategoria")
        chart_data = studio_isa[studio_isa["FamigliaCategoria"]!="Totale"]
        fig, ax = plt.subplots(figsize=(8,5))
        ax.bar(chart_data["FamigliaCategoria"], chart_data["Netto"])
        ax.set_title("Somma Netto per FamigliaCategoria")
        plt.xticks(rotation=45, ha="right")
        buf = BytesIO(); plt.tight_layout(); plt.savefig(buf, format="png"); buf.seek(0)

        # Excel export
        wb = Workbook(); ws = wb.active; ws.title = "Report_A"
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
            c=ws.cell(row=tot_row_idx,column=j); c.font=Font(bold=True); c.fill=total_fill
        img=XLImage(buf); img.anchor=f"A{tot_row_idx+3}"; ws.add_image(img)
        out=BytesIO(); wb.save(out)
        st.download_button("⬇️ Scarica Excel (DrVeto)", data=out.getvalue(), file_name=f"StudioISA_DrVeto_{datetime.now().year}.xlsx")

        st.stop()

    # ──────────────── TIPO B ────────────────
    else:
        col_prest = pick_column(df, "PrestazioneProdotto", "Prestazione prodotto", "Prestazione")
        col_cat   = pick_column(df, "Categoria")
        col_imp   = pick_column(df, "TotaleImponibile", "Totale imponibile", "Imponibile")
        col_iva   = pick_column(df, "Totaleconiva", "Totale con iva", "TotaleConIVA", "Totale IVA")
        col_tot   = pick_column(df, "Totale")

        if not all([col_prest, col_cat, col_imp, col_iva, col_tot]):
            st.error("❌ Colonne richieste per il Tipo B non trovate.")
            st.stop()

        # Compila Categoria con regole + memoria
        df["Categoria"] = df.apply(lambda r: classify_B(r[col_prest], r[col_cat], user_memory), axis=1)

        # Apprendimento: nuovi termini su PrestazioneProdotto
        uniq_terms = sorted({str(v).strip() for v in df[col_prest].dropna().unique()}, key=lambda s: s.casefold())
        pending = []
        for term in uniq_terms:
            t = norm(term)
            if not any(norm(k) in t for k in user_memory.keys()) and \
               not any(any_kw_in(t, kws) for kws in RULES_B.values()):
                pending.append(term)

        if pending:
            if "idx" not in st.session_state: st.session_state.idx = 0
            if st.session_state.idx >= len(pending): st.session_state.idx = 0
            term = pending[st.session_state.idx]
            st.warning(f"🧠 Nuovo termine (Tipo B): {term} [{st.session_state.idx+1}/{len(pending)}]")

            default_opts = list(RULES_B.keys())
            c1, c2 = st.columns([2,1])
            with c1:
                cat_sel = st.selectbox("Seleziona categoria:", default_opts, key=f"b_sel_{st.session_state.idx}")
                cat_free = st.text_input("…oppure categoria personalizzata (opzionale)", key=f"b_free_{st.session_state.idx}")
                final_cat = cat_free.strip() if cat_free.strip() else cat_sel
            with c2:
                if st.button("✅ Salva e continua"):
                    user_memory[term] = final_cat
                    github_save_json(user_memory)
                    st.session_state.user_memory = user_memory
                    st.session_state.idx += 1
                    st.rerun()
            st.stop()

        # Report (Tipo B): Categoria / TotaleImponibile / TotaleConIVA / Totale / % Totale
        st.success("✅ Nessun nuovo termine. Genero il report (VetsGo)…")

        studio_b = (
            df.groupby("Categoria", dropna=False)
              .agg({col_imp:"sum", col_iva:"sum", col_tot:"sum"})
              .reset_index()
              .rename(columns={col_imp:"TotaleImponibile", col_iva:"TotaleConIVA", col_tot:"Totale"})
        )

        # Ordina per Totale desc
        studio_b = studio_b.sort_values("Totale", ascending=False, ignore_index=True)

        tot_totale = studio_b["Totale"].sum()
        studio_b["% Totale"] = (studio_b["Totale"]/tot_totale*100).round(2) if tot_totale else 0.0
        studio_b = pd.concat([studio_b, pd.DataFrame([["Totale", studio_b["TotaleImponibile"].sum(), studio_b["TotaleConIVA"].sum(), tot_totale, 100]], columns=studio_b.columns)], ignore_index=True)

        st.subheader("📄 Tabella Studio ISA (VetsGo)")
        st.dataframe(
            studio_b.style
              .apply(lambda r: ['background-color: #fff8b3' if r["Categoria"]=="Totale" else '' for _ in r], axis=1)
              .set_properties(subset=["Categoria"], **{"font-weight":"bold"})
              .format({"TotaleImponibile":"{:,.2f}", "TotaleConIVA":"{:,.2f}", "Totale":"{:,.2f}", "% Totale":"{:.2f}"})
        )

        # Grafico: Somma Totale per Categoria
        st.subheader("📊 Somma Totale per Categoria")
        chart_data = studio_b[studio_b["Categoria"]!="Totale"]
        fig, ax = plt.subplots(figsize=(8,5))
        ax.bar(chart_data["Categoria"], chart_data["Totale"])
        ax.set_title("Somma Totale per Categoria")
        plt.xticks(rotation=45, ha="right")
        buf = BytesIO(); plt.tight_layout(); plt.savefig(buf, format="png"); buf.seek(0)

        # Excel export
        wb = Workbook(); ws = wb.active; ws.title = "Report_B"
        start_row, start_col = 3, 2
        total_fill = PatternFill(start_color="FFF4B084", end_color="FFF4B084", fill_type="solid")
        headers = ["Categoria","TotaleImponibile","TotaleConIVA","Totale","% Totale"]
        for j,h in enumerate(headers,start=start_col):
            ws.cell(row=start_row,column=j,value=h).font = Font(bold=True)
        for i,row in enumerate(dataframe_to_rows(studio_b,index=False,header=False),start=start_row+1):
            for j,v in enumerate(row,start=start_col):
                ws.cell(row=i,column=j,value=v)
        tot_row_idx = start_row+len(studio_b)
        for j in range(start_col, start_col+len(headers)):
            c=ws.cell(row=tot_row_idx,column=j); c.font=Font(bold=True); c.fill=total_fill
        img=XLImage(buf); img.anchor=f"A{tot_row_idx+3}"; ws.add_image(img)
        out=BytesIO(); wb.save(out)
        st.download_button("⬇️ Scarica Excel (VetsGo)", data=out.getvalue(), file_name=f"StudioISA_VetsGo_{datetime.now().year}.xlsx")

if __name__ == "__main__":
    main()

