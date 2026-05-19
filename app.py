import streamlit as st
import pandas as pd
import re
import datetime
import io

# Nastavení vzhledu stránky a záložky v prohlížeči
st.set_page_config(page_title="Wolfnet - Převodník dat z mappingu", page_icon="🌐")

# --- HLAVIČKA APLIKACE S POPISY ---
st.title("🌐 Wolfnet - Převodník dat z mappingu")
st.markdown("""
**Autor:** Honza Kropáček  
**Datum:** 19.05.2026  
---
""")
st.write("Nahrajte exportovaný CSV soubor a stáhněte si upravený Excel.")

# 1. Webový prvek pro nahrání souboru
vstupni_soubor = st.file_uploader("Vyberte CSV soubor", type=['csv'])

if vstupni_soubor is not None:
    with st.spinner('Zpracovávám data...'):
        
        # 2. Načtení dat přímo z nahraného souboru
        df = pd.read_csv(vstupni_soubor, sep=';')

        # --- Rozpad adresy ---
        df['ulice_cp'] = df['tid'].str.split(',').str[0].str.strip()
        df['mesto'] = df['tid'].str.split(',').str[1].str.strip()
        df['mesto'] = df['mesto'].str.replace(r'^\d{3}\s?\d{2}\s+', '', regex=True)

        # --- Určení typu připojení (kabel / WIFI) ---
        df['druh_pripojeni'] = 'WIFI'
        
        mask = df['_state'].str.contains('CATV|FTTB|FTTH', na=False, regex=True)
        df.loc[mask, 'druh_pripojeni'] = 'kabel'

        # --- Očištění textů (pro všechny řádky, tedy i pro WIFI) ---
        df['_state'] = df['_state'].str.replace('Připojeno - ', '', regex=False)
        df['_state'] = df['_state'].str.replace(r' od \d{2}\.\d{2}\.\d{4}', '', regex=True)
        df['_state'] = df['_state'].str.replace(r'\s*\(.*?\)', '', regex=True)
        df['_state'] = df['_state'].str.strip()

        # --- Převod na čísla a úprava procent ---
        df['number_of_apartments'] = pd.to_numeric(df['number_of_apartments'], errors='coerce').astype('Int64')
        df['number_of_floors'] = pd.to_numeric(df['number_of_floors'], errors='coerce').astype('Int64')
        df['_penetration'] = pd.to_numeric(df['_penetration'], errors='coerce') / 100

        # --- Odstranění nechtěných sloupců ---
        df = df.drop(columns=['price_level_id', 'ctu_vhcn', 'metadata'])

        # --- Přejmenování sloupců přesně podle zadání ---
        df = df.rename(columns={
            'ar_address_place_code': 'RUIAN',
            'tid': 'adresa',
            '_state': 'technologie',
            'number_of_apartments': 'počet bytů',
            'number_of_floors': 'počet pater',
            '_active_connections_points': 'připojených',
            '_penetration': 'penetrace v %'
        })

        # 3. Vytvoření Excelu do paměti
        output = io.BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df.to_excel(writer, index=False, sheet_name='Data')
        
        workbook = writer.book
        worksheet = writer.sheets['Data']
        
        worksheet.freeze_panes(1, 0)
        (max_row, max_col) = df.shape
        worksheet.autofilter(0, 0, max_row, max_col - 1)
        
        # Aplikace formátu procent podle nového českého názvu sloupce
        idx_penetration = df.columns.get_loc('penetrace v %') 
        pct_format = workbook.add_format({'num_format': '0.00%'})
        worksheet.set_column(idx_penetration, idx_penetration, None, pct_format)
        
        writer.close()
        output.seek(0)
        
        # 4. Příprava názvu pro stažení
        dnes = datetime.datetime.now().strftime("%d.%m.%Y")
        vychozi_nazev = f"schéma sítě wolfnet {dnes}.xlsx"
        
        st.success("Data byla úspěšně zpracována!")

        # 5. Tlačítko pro stažení
        st.download_button(
            label="📥 Stáhnout upravený Excel",
            data=output,
            file_name=vychozi_nazev,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
