import os
import re
import json
import difflib
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd
import xlwings as xw

# Per usare COM in un thread (xlwings) in modo sicuro
try:
    import pythoncom
except ImportError:
    pythoncom = None

# =========================
#  CONFIG & COSTANTI
# =========================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
KEYWORDS_PATH = os.path.join(BASE_DIR, "keywords.json")

APP_TITLE = "Studio ISA - Alcyon Italia SpA - 2025"

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

# =========================
#  DIZIONARIO (load/save)
# =========================

def load_keywords(path: str) -> dict:
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            data = {}
    else:
        data = {}
    # garantisci categorie & unisci default senza duplicati
    for cat in CATEGORIES:
        data.setdefault(cat, [])
    for cat, words in DEFAULT_KEYWORDS.items():
        known = set(w.lower() for w in data.get(cat, []))
        for w in words:
            if w.lower() not in known:
                data[cat].append(w)
                known.add(w.lower())
    return data

def save_keywords(path: str, data: dict) -> None:
    clean = {}
    for cat, words in data.items():
        seen = set()
        out = []
        for w in words:
            w = str(w).strip()
            if not w:
                continue
            lw = w.lower()
            if lw not in seen:
                out.append(w)
                seen.add(lw)
        clean[cat] = out
    with open(path, "w", encoding="utf-8") as f:
        json.dump(clean, f, ensure_ascii=False, indent=2)

# =========================
#  CLASSIFICAZIONE
# =========================

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
            # confini stretti per evitare 'rx' in 'enrox'
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
    # priorità ottimizzata
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

def best_match_category(text: str, keywords: dict, cutoff: int = 92) -> tuple[str, str, int]:
    t = _norm(text)
    best_cat, best_kw, best_score = "ALTRE PRESTAZIONI", "", 0
    # parola intera
    for cat, words in keywords.items():
        for kw in words:
            k = kw.lower().strip()
            if re.search(rf"(?<![a-zA-Z]){re.escape(k)}(?![a-zA-Z])", t):
                return cat, kw, 100
    # similarità prudente
    for cat, words in keywords.items():
        for kw in words:
            score = int(difflib.SequenceMatcher(None, kw.lower().strip(), t).ratio() * 100)
            if score > best_score:
                best_score, best_cat, best_kw = score, cat, kw
    if best_score >= cutoff:
        return best_cat, best_kw, best_score
    return "ALTRE PRESTAZIONI", "", best_score

def extract_token_for_learning(text: str) -> str:
    t = re.sub(r"[=+.,;:!?\(\)\[\]]", " ", str(text))
    parts = [p.strip().lower() for p in t.split() if p.strip()]
    for p in parts:
        if any(ch.isdigit() for ch in p):
            return p
    for p in parts:
        if len(p) >= 5:
            return p
    return parts[0] if parts else ""

# =========================
#  AUTO-COMPILAZIONE
# =========================

def map_and_clean_columns(cols):
    rename_map = {}
    for c in cols:
        s = str(c).strip()
        if s == "%": rename_map[c] = "Perc"
        elif "netto" in s.lower() and "dopo" in s.lower(): rename_map[c] = "Netto"
        elif ("famiglia" in s.lower() and "categoria" in s.lower()) or ("famiglia" in s.lower() and "/" in s):
            rename_map[c] = "FamigliaCategoria"
        else:
            cleaned = re.sub(r"[^\w]", "", s)
            rename_map[c] = cleaned if cleaned else "Col"
    return rename_map

def auto_fill_famiglia_categoria(df, keywords, supervised=False, parent=None):
    desc_col = next((c for c in df.columns if "descrizione" in c.lower()), None)
    if not desc_col:
        return df, keywords, 0

    df["FamigliaCategoria"] = df["FamigliaCategoria"].fillna("").astype(str)
    filled = 0
    unknown_tokens = {}

    for idx, row in df.iterrows():
        if not row["FamigliaCategoria"].strip():
            descr = str(row[desc_col]).strip()
            if not descr:
                df.at[idx, "FamigliaCategoria"] = "ALTRE PRESTAZIONI"
                continue

            # 1) regole con priorità
            cat_rule = rule_based_category(descr)
            if cat_rule != "ALTRE PRESTAZIONI":
                df.at[idx, "FamigliaCategoria"] = cat_rule
                filled += 1
                continue

            # 2) dizionario/fuzzy prudente
            cat, kw, score = best_match_category(descr, keywords, cutoff=92)
            if cat == "ALTRE PRESTAZIONI" or score < 92:
                if supervised:
                    token = extract_token_for_learning(descr)
                    if token:
                        unknown_tokens[token] = cat_rule  # default di suggerimento
                else:
                    df.at[idx, "FamigliaCategoria"] = cat_rule
            else:
                df.at[idx, "FamigliaCategoria"] = cat
                filled += 1

    # Apprendimento: chiedi 1 volta per token
    if supervised and unknown_tokens:
        for token, suggested in unknown_tokens.items():
            chosen = prompt_category_for_token(token, suggested, parent)
            if chosen and chosen in CATEGORIES:
                keywords.setdefault(chosen, [])
                if token not in [w.lower() for w in keywords[chosen]]:
                    keywords[chosen].append(token)

        # seconda passata con dizionario aggiornato
        for idx, row in df.iterrows():
            if not row["FamigliaCategoria"].strip():
                descr = str(row[desc_col]).strip()
                cat_rule = rule_based_category(descr)
                if cat_rule != "ALTRE PRESTAZIONI":
                    df.at[idx, "FamigliaCategoria"] = cat_rule
                    filled += 1
                else:
                    cat, kw, score = best_match_category(descr, keywords, cutoff=92)
                    df.at[idx, "FamigliaCategoria"] = cat if score >= 92 else "ALTRE PRESTAZIONI"
                    filled += 1

    return df, keywords, filled

# =========================
#  POPUP SUPERVISIONE
# =========================

def prompt_category_for_token(token: str, suggested: str, parent) -> str:
    top = tk.Toplevel(parent)
    top.title("Apprendimento supervisionato")
    top.grab_set()

    tk.Label(top, text="Nuovo termine rilevato:", font=("Segoe UI", 9, "bold")).pack(padx=12, pady=(10, 2), anchor="w")
    tk.Label(top, text=f"“{token}”").pack(padx=12, pady=(0, 10), anchor="w")
    tk.Label(top, text="Scegli la categoria:", font=("Segoe UI", 9)).pack(padx=12, anchor="w")

    var = tk.StringVar(value=suggested if suggested in CATEGORIES else "ALTRE PRESTAZIONI")
    cmb = ttk.Combobox(top, textvariable=var, values=CATEGORIES, state="readonly", width=34)
    cmb.pack(padx=12, pady=6)

    result = {"choice": None}
    ttk.Button(top, text="Conferma", command=lambda: (setattr(type('obj', (), {'x': None})(), 'x', None), top.destroy(), result.update(choice=var.get()))).pack(pady=10)

    # centra
    top.update_idletasks()
    w, h = 360, 180
    sw = top.winfo_screenwidth()
    sh = top.winfo_screenheight()
    x = (sw - w)//2
    y = (sh - h)//2
    top.geometry(f"{w}x{h}+{x}+{y}")

    top.wait_window()
    return result["choice"]

# =========================
#  EXCEL: PIVOT & REPORT
# =========================

def build_excel(input_file, supervised=False, parent=None):
    """Funzione pesante: lettura, classificazione, pivot, grafico, salvataggio."""
    # Inizializza COM nel thread (se in thread separato)
    if pythoncom is not None:
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass

    keywords = load_keywords(KEYWORDS_PATH)

    df = pd.read_excel(input_file)
    df = df.rename(columns=map_and_clean_columns(df.columns))
    df, keywords, n_filled = auto_fill_famiglia_categoria(df, keywords, supervised, parent)

    # controlli minimi
    for req in ["FamigliaCategoria", "Perc", "Netto"]:
        if req not in df.columns:
            raise ValueError(f"Manca la colonna {req}")

    df["FamigliaCategoria"] = df["FamigliaCategoria"].astype(str).str.strip()

    # anno dal primo valore in colonna "Data ..."
    date_col = next((c for c in df.columns if "data" in c.lower()), None)
    anno = "ND"
    if date_col:
        first = df[date_col].dropna()
        if not first.empty:
            try:
                anno = pd.to_datetime(first.iloc[0]).year
            except Exception:
                pass

    folder = os.path.dirname(input_file)
    output_file = os.path.join(folder, f"Studio ISA {anno}.xlsx")

    # Excel invisibile
    app = xw.App(visible=False)
    wb = app.books.add()

    # DatiOriginali
    ws_data = wb.sheets.add("DatiOriginali")
    ws_data.range("A1").value = df
    rng = ws_data.range((1, 1), (df.shape[0] + 1, df.shape[1])).api
    tbl = ws_data.api.ListObjects.Add(1, rng, 0, 1, 1)
    tbl.Name = "TblDati"
    tbl.ShowHeaders = True

    # Report
    ws_report = wb.sheets.add("Report")

    studio = df.groupby("FamigliaCategoria", dropna=False).agg({"Perc": "sum", "Netto": "sum"}).reset_index().rename(columns={"Perc": "Qtà"})
    tq, tn = studio["Qtà"].sum(), studio["Netto"].sum()
    studio["% Qtà"] = (studio["Qtà"] / tq * 100).round(2) if tq else 0
    studio["% Netto"] = (studio["Netto"] / tn * 100).round(2) if tn else 0

    start_row, start_col = 3, 2
    ws_report.range((start_row, start_col)).value = ["FamigliaCategoria", "Qtà", "Netto", "% Qtà", "% Netto"]
    ws_report.range((start_row, start_col), (start_row, start_col + 4)).api.Font.Bold = True
    ws_report.range((start_row + 1, start_col)).value = studio.values
    ws_report.range((start_row + 1 + len(studio), start_col)).value = ["Totale", tq, tn, 100, 100]
    ws_report.range((start_row + 1 + len(studio), start_col), (start_row + 1 + len(studio), start_col + 4)).api.Font.Bold = True

    # Pivot a destra
    pivot_col = start_col + 7
    pc = wb.api.PivotCaches().Create(SourceType=1, SourceData=tbl.Range)
    pt = pc.CreatePivotTable(TableDestination=ws_report.range((start_row, pivot_col)).api, TableName="PivotFamiglia")
    pt.PivotFields("FamigliaCategoria").Orientation = 1
    pt.AddDataField(pt.PivotFields("Perc"), "Somma_Qta", -4157)
    pt.AddDataField(pt.PivotFields("Netto"), "Somma_Netto", -4157)

    # Grafico sotto pivot (colonne raggruppate)
    chart = ws_report.api.ChartObjects().Add(
        Left=ws_report.range((start_row, pivot_col)).api.Left,
        Top=ws_report.range((start_row + pt.TableRange2.Rows.Count + 2, pivot_col)).api.Top,
        Width=600, Height=350
    ).Chart
    chart.SetSourceData(pt.TableRange2)
    chart.ChartType = 51  # xlColumnClustered
    chart.HasTitle = True
    chart.ChartTitle.Text = "Pivot - Somma_Qta e Somma_Netto per FamigliaCategoria"

    wb.save(output_file)
    wb.close()
    app.quit()

    save_keywords(KEYWORDS_PATH, keywords)
    return output_file, n_filled

# =========================
#  GESTIONE DIZIONARIO (GUI)
# =========================

def open_dictionary_manager(root):
    kws = load_keywords(KEYWORDS_PATH)

    win = tk.Toplevel(root)
    win.title("Gestione Dizionario")
    win.geometry("560x380")
    win.grab_set()

    tk.Label(win, text="Categoria:", font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=10, pady=(10, 2))
    cat_var = tk.StringVar(value=CATEGORIES[0])
    cat_combo = ttk.Combobox(win, textvariable=cat_var, values=CATEGORIES, state="readonly", width=38)
    cat_combo.pack(anchor="w", padx=10)

    tk.Label(win, text="Parole nella categoria:", font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=10, pady=(12, 2))
    frame_list = tk.Frame(win)
    frame_list.pack(fill="both", expand=True, padx=10, pady=(0, 8))
    listbox = tk.Listbox(frame_list, selectmode="extended")
    listbox.pack(side="left", fill="both", expand=True)
    scrollbar = ttk.Scrollbar(frame_list, orient="vertical", command=listbox.yview)
    scrollbar.pack(side="right", fill="y")
    listbox.config(yscrollcommand=scrollbar.set)

    def refresh_list():
        listbox.delete(0, "end")
        words = sorted(set(kws.get(cat_var.get(), [])), key=str.lower)
        for w in words:
            listbox.insert("end", w)

    cat_combo.bind("<<ComboboxSelected>>", lambda e: refresh_list())

    tk.Label(win, text="Aggiungi parola:", font=("Segoe UI", 9)).pack(anchor="w", padx=10)
    add_var = tk.StringVar()
    tk.Entry(win, textvariable=add_var).pack(anchor="w", padx=10, fill="x")

    btns = tk.Frame(win); btns.pack(anchor="w", padx=10, pady=8)

    def add_word():
        w = add_var.get().strip()
        if not w:
            return
        cat = cat_var.get()
        kws.setdefault(cat, [])
        if w.lower() not in [x.lower() for x in kws[cat]]:
            kws[cat].append(w)
            add_var.set("")
            refresh_list()

    def remove_selected():
        cat = cat_var.get()
        sel = list(listbox.curselection())
        if not sel:
            return
        values = [listbox.get(i) for i in sel]
        kws[cat] = [w for w in kws.get(cat, []) if w not in values]
        refresh_list()

    ttk.Button(btns, text="➕ Aggiungi", command=add_word).pack(side="left", padx=(0, 8))
    ttk.Button(btns, text="🗑️ Rimuovi selezionati", command=remove_selected).pack(side="left")

    action = tk.Frame(win); action.pack(fill="x", padx=10, pady=10)

    def do_save():
        save_keywords(KEYWORDS_PATH, kws)
        messagebox.showinfo("Salvato", f"✅ Dizionario aggiornato in:\n{KEYWORDS_PATH}")

    ttk.Button(action, text="💾 Salva modifiche", command=do_save).pack(side="left")
    ttk.Button(action, text="Chiudi", command=win.destroy).pack(side="right")

    refresh_list()

# =========================
#  SPLASH SCREEN (solo AUTO)
# =========================

class Splash:
    def __init__(self, root):
        self.root = root
        self.top = tk.Toplevel(root)
        self.top.overrideredirect(True)  # stile “splash”
        self.top.configure(bg="#ffffff")
        self.top.attributes("-topmost", True)

        # contenuto
        frame = tk.Frame(self.top, bg="#ffffff", bd=2, relief="flat")
        frame.pack(padx=18, pady=14)

        tk.Label(frame, text="💼 Elaborazione in corso…", font=("Segoe UI", 11, "bold"), bg="#ffffff").pack(pady=(2, 4))
        tk.Label(frame, text="Attendere qualche secondo", font=("Segoe UI", 9), fg="#666", bg="#ffffff").pack(pady=(0, 8))

        self.pb = ttk.Progressbar(frame, mode="indeterminate", length=260)
        self.pb.pack(pady=(2, 2))
        self.pb.start(12)

        # centra al centro dello schermo
        self.top.update_idletasks()
        w, h = 340, 120
        sw = self.top.winfo_screenwidth()
        sh = self.top.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.top.geometry(f"{w}x{h}+{x}+{y}")

    def close(self):
        try:
            self.pb.stop()
        except Exception:
            pass
        self.top.destroy()

# =========================
#  AVVIO ELABORAZIONE
# =========================

def on_select_file_auto(root):
    path = filedialog.askopenfilename(title="Seleziona file Excel", filetypes=[("Excel files", "*.xlsx;*.xls")])
    if not path:
        return

    splash = Splash(root)

    def worker():
        try:
            out, n = build_excel(path, supervised=False, parent=root)
            root.after(0, lambda: [splash.close(), messagebox.showinfo("Completato", f"✅ File salvato in:\n{out}\nCategorie compilate: {n}")])
        except Exception as e:
            root.after(0, lambda: [splash.close(), messagebox.showerror("Errore", str(e))])

    threading.Thread(target=worker, daemon=True).start()

def on_select_file_supervised(root):
    path = filedialog.askopenfilename(title="Seleziona file Excel", filetypes=[("Excel files", "*.xlsx;*.xls")])
    if not path:
        return
    # Nessuno splash qui, per non coprire i popup di apprendimento
    try:
        out, n = build_excel(path, supervised=True, parent=root)
        messagebox.showinfo("Completato", f"✅ File salvato in:\n{out}\nCategorie compilate: {n}")
    except Exception as e:
        messagebox.showerror("Errore", str(e))

# =========================
#  GUI PRINCIPALE
# =========================

def main():
    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("460x260")

    tk.Label(root, text="Studio ISA", font=("Segoe UI", 12, "bold")).pack(pady=(14, 4))
    tk.Label(root, text="Genera Tabella + Pivot + Grafico + File Output Studio ISA <anno>.xlsx", fg="#555").pack(pady=(0, 8))

    ttk.Button(root, text="📂 Elabora file automaticamente", width=38, command=lambda: on_select_file_auto(root)).pack(pady=6)
    ttk.Button(root, text="🧠 Elabora file manualmente", width=38, command=lambda: on_select_file_supervised(root)).pack(pady=6)
    ttk.Button(root, text="🗂️ Gestisci Dizionario", width=38, command=lambda: open_dictionary_manager(root)).pack(pady=(10, 6))
    ttk.Button(root, text="❌ Esci", width=38, command=root.destroy).pack(pady=6)

    root.mainloop()

if __name__ == "__main__":
    main()
