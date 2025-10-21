import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
from datetime import datetime
import re
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image as XLImage

st.set_page_config(page_title="Studio ISA", page_icon="🐾", layout="centered")

st.title("🐾 Studio ISA")
st.write("Carica il file Excel per generare automaticamente la tabella Studio ISA, la pivot e il grafico.")

uploaded_file = st.file_uploader("📤 Carica file Excel", type=["xlsx", "xls"])
progress_bar = st.progress(0)

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

if uploaded_file:
    try:
        progress_bar.progress(10)
        df = pd.read_excel(uploaded_file)
        rename_map = map_and_clean_columns(df.columns)
        df = df.rename(columns=rename_map)
        required = ['FamigliaCategoria', 'Perc', 'Netto']
        missing = [r for r in required if r not in df.columns]
        if missing:
            st.error(f"Mancano colonne obbligatorie dopo la mappatura: {missing}")
            st.stop()

        df['FamigliaCategoria'] = df['FamigliaCategoria'].astype(str).str.strip()
        progress_bar.progress(25)

        # === Tabella Studio ISA ===
        studio_isa = df.groupby('FamigliaCategoria', dropna=False).agg({
            'Perc': 'sum',
            'Netto': 'sum'
        }).reset_index().rename(columns={'Perc': 'Qtà'})

        tot_qta = studio_isa['Qtà'].sum()
        tot_netto = studio_isa['Netto'].sum()

        studio_isa['% Qtà'] = (studio_isa['Qtà'] / tot_qta * 100).round(2)
        studio_isa['% Netto'] = (studio_isa['Netto'] / tot_netto * 100).round(2)

        total_row = pd.DataFrame({
            'FamigliaCategoria': ['Totale'],
            'Qtà': [tot_qta],
            'Netto': [tot_netto],
            '% Qtà': [100],
            '% Netto': [100]
        })
        studio_isa = pd.concat([studio_isa, total_row], ignore_index=True)

        progress_bar.progress(50)

        # === Crea Pivot (pandas style) ===
        pivot = pd.pivot_table(
            df,
            values=['Perc', 'Netto'],
            index=['FamigliaCategoria'],
            aggfunc='sum',
            fill_value=0
        ).reset_index()
        pivot = pivot.rename(columns={'Perc': 'Somma_Qta', 'Netto': 'Somma_Netto'})

        # === Grafico ===
        fig, ax = plt.subplots(figsize=(8, 4))
        pivot.plot(
            kind='bar',
            x='FamigliaCategoria',
            y=['Somma_Qta', 'Somma_Netto'],
            ax=ax
        )
        ax.set_title("Somma_Qta e Somma_Netto per FamigliaCategoria")
        ax.set_ylabel("Valore")
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()

        img_buf = io.BytesIO()
        plt.savefig(img_buf, format='png')
        plt.close(fig)
        img_buf.seek(0)
        progress_bar.progress(75)

        # === Crea Excel con openpyxl ===
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "Report"

        ws1.append(["Studio ISA"])
        for r in dataframe_to_rows(studio_isa, index=False, header=True):
            ws1.append(r)

        ws1.append([])
        ws1.append(["Pivot"])
        for r in dataframe_to_rows(pivot, index=False, header=True):
            ws1.append(r)

        # Inserisci grafico come immagine
        img = XLImage(img_buf)
        img.anchor = f"A{len(ws1['A']) + 2}"
        ws1.add_image(img)

        # === Trova anno da colonna data (se esiste) ===
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

        # === Salva in memoria e consenti il download ===
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        progress_bar.progress(100)
        st.success("✅ Elaborazione completata!")
        st.download_button(
            label="📥 Scarica file Excel",
            data=output,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Errore: {e}")
