import streamlit as st
import pandas as pd
import xlwings as xw
import re
import io
from datetime import datetime

st.set_page_config(page_title="Studio ISA", page_icon="🐾", layout="centered")

st.title("🐾 Studio ISA")
st.write("Carica il file Excel per generare automaticamente la tabella pivot e il report.")

# === UPLOAD FILE ===
uploaded_file = st.file_uploader("📤 Carica file Excel", type=["xlsx", "xls"])

# barra di progresso (placeholder)
progress_text = st.empty()
progress_bar = st.progress(0)

if uploaded_file:
    try:
        # === LEGGE IL FILE ===
        progress_text.text("🔍 Lettura del file in corso...")
        df = pd.read_excel(uploaded_file)
        progress_bar.progress(10)

        # === FUNZIONE DI MAPPATURA COLONNE ===
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
                    cleaned = re.sub(r'[^\w]', '', s)
                    rename_map[c] = cleaned if cleaned else "Col"
            return rename_map

        rename_map = map_and_clean_columns(df.columns)
        df = df.rename(columns=rename_map)

        required = ['FamigliaCategoria', 'Perc', 'Netto']
        missing = [r for r in required if r not in df.columns]
        if missing:
            st.error(f"Mancano colonne obbligatorie dopo la mappatura: {missing}")
            st.stop()

        df['FamigliaCategoria'] = df['FamigliaCategoria'].astype(str).str.strip()
        progress_bar.progress(25)

        # === CREA FILE EXCEL IN MEMORIA ===
        output = io.BytesIO()
        app = xw.App(visible=False)
        wb = app.books.add()

        # --- Dati originali ---
        ws_data = wb.sheets.add("DatiOriginali")
        ws_data.range("A1").value = df
        progress_bar.progress(40)

        # --- Tabella Excel ---
        last_row, last_col = df.shape
        rng = ws_data.range((1,1), (last_row+1, last_col)).api
        tbl = ws_data.api.ListObjects.Add(1, rng, 0, 1, 1)
        tbl.Name = "TblDati"
        tbl.ShowHeaders = True

        # --- Foglio report ---
        ws_report = wb.sheets.add("Report")
        ws_report.activate()

        # --- Tabella Studio ISA ---
        studio_isa = df.groupby('FamigliaCategoria', dropna=False).agg({
            'Perc': 'sum',
            'Netto': 'sum'
        }).reset_index().rename(columns={'Perc':'Qtà'})

        tot_qta = studio_isa['Qtà'].sum()
        tot_netto = studio_isa['Netto'].sum()

        studio_isa['% Qtà'] = (studio_isa['Qtà'] / tot_qta * 100).round(2) if tot_qta != 0 else 0
        studio_isa['% Netto'] = (studio_isa['Netto'] / tot_netto * 100).round(2) if tot_netto != 0 else 0

        isa_start_row = 3
        isa_start_col = 2
        ws_report.range((isa_start_row, isa_start_col)).value = ["FamigliaCategoria", "Qtà", "Netto", "% Qtà", "% Netto"]
        ws_report.range((isa_start_row+1, isa_start_col)).value = studio_isa.values

        # Grassetto intestazioni
        ws_report.range((isa_start_row, isa_start_col), (isa_start_row, isa_start_col+4)).api.Font.Bold = True
        progress_bar.progress(60)

        # Totali
        tot_row_idx = isa_start_row + 1 + len(studio_isa)
        totali = ["Totale",
                studio_isa['Qtà'].sum(),
                studio_isa['Netto'].sum(),
                100,
                100]
        ws_report.range((tot_row_idx, isa_start_col)).value = totali
        ws_report.range((tot_row_idx, isa_start_col), (tot_row_idx, isa_start_col+4)).api.Font.Bold = True

        # === PIVOT ===
        pc = wb.api.PivotCaches().Create(SourceType=1, SourceData=tbl.Range)
        pivot_start_row = isa_start_row
        pivot_start_col = isa_start_col + 7
        pivot_dest = ws_report.range((pivot_start_row, pivot_start_col)).api

        pt = pc.CreatePivotTable(TableDestination=pivot_dest, TableName="PivotFamiglia")
        pt.PivotFields("FamigliaCategoria").Orientation = 1
        pt.AddDataField(pt.PivotFields("Perc"), "Somma_Qta", -4157)
        pt.AddDataField(pt.PivotFields("Netto"), "Somma_Netto", -4157)

        # === GRAFICO ===
        chart_obj = ws_report.api.ChartObjects().Add(
            Left=ws_report.range((pivot_start_row, pivot_start_col)).api.Left,
            Top=ws_report.range((pivot_start_row + pt.TableRange2.Rows.Count + 2, pivot_start_col)).api.Top,
            Width=600,
            Height=350
        )
        chart = chart_obj.Chart
        chart.SetSourceData(pt.TableRange2)
        chart.ChartType = 51  # xlColumnClustered
        chart.HasTitle = True
        chart.ChartTitle.Text = "Pivot - Somma_Qta e Somma_Netto per FamigliaCategoria"

        progress_bar.progress(90)

        # === NOME FILE OUTPUT ===
        anno = None
        for c in df.columns:
            if "data" in c.lower():
                try:
                    data_col = pd.to_datetime(df[c], errors='coerce')
                    anno = data_col.dt.year.dropna().iloc[0]
                    break
                except Exception:
                    continue
        if anno is None:
            anno = datetime.now().year

        output_filename = f"Studio_ISA_{anno}.xlsx"

        # === SALVA E PREPARA DOWNLOAD ===
        wb.save(output)
        wb.close()
        app.quit()
        progress_bar.progress(100)

        st.success("✅ Elaborazione completata!")
        st.download_button(
            label="📥 Scarica file Excel",
            data=output.getvalue(),
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Errore: {e}")
        try:
            app.quit()
        except:
            pass