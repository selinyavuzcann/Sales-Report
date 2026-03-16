import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import PatternFill, Font, Alignment

def process_data(data_file):
    # Load the original data sheets
    df_personel = pd.read_excel(data_file, sheet_name='PERSONEL SATIŞ AYLIK')
    df_marka = pd.read_excel(data_file, sheet_name='MARKA SATIŞ')
    
    # --- Part 1: Prepare Personel Data (Columns A-H) ---
    personel_part = pd.DataFrame()
    personel_part['Satıcı Ad Soyad'] = df_personel['Satici DESC']
    personel_part['Mağaza Lokasyon'] = df_personel['Lokasyon DESC']
    
    # Mağaza Ciro Calculation (Mapping from Marka sheet)
    marka_ciro_map = df_marka.groupby('Lokasyon Donusum DESC')['SAP Satış Net Tutar KDV\'siz HPD'].sum().to_dict()
    personel_part['Mağaza Ciro'] = personel_part['Mağaza Lokasyon'].map(marka_ciro_map)
    
    personel_part['Satıcı Ciro'] = df_personel['Pos Kasa Satış Net Tutar HPD']
    personel_part['Başarı Yüzdesi'] = None # Placeholder
    personel_part['Çalışma Saati'] = None # Blank for User Input
    personel_part['Satılan Adet'] = df_personel['Pos Kasa Satış Net Miktar']
    personel_part['Verimlilik'] = None # Placeholder

    # --- Part 2: Prepare Marka Data (Columns I-M) ---
    marka_part = pd.DataFrame()
    marka_part['MARKA_LOKASYON'] = df_marka['Lokasyon Donusum DESC']
    marka_part['MODEL ADI'] = df_marka['Urun Model ID']
    marka_part['ADET'] = df_marka['SAP Satış Net Miktar']
    marka_part['CİRO'] = df_marka['SAP Satış Net Tutar KDV\'siz HPD']
    marka_part['STOK'] = df_marka['Stok Kullanilabilir Miktar (Tahditsiz Stok Full)']

    # --- Part 3: Combine them without limits ---
    # axis=1 ensures they are side-by-side. 
    output_df = pd.concat([personel_part, marka_part], axis=1)

    # 4. Create Excel with Multiple Sheets
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        output_df.to_excel(writer, sheet_name='Sales Report', index=False)
        df_personel.to_excel(writer, sheet_name='PERSONEL SATIŞ AYLIK', index=False)
        df_marka.to_excel(writer, sheet_name='MARKA SATIŞ', index=False)
        
        workbook = writer.book
        worksheet = workbook['Sales Report']
        
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        center_align = Alignment(horizontal="center")

        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align

        # Identify limits
        len_personel = len(df_personel)
        len_total = len(output_df)

        # Loop through all rows for general number formatting
        for i in range(2, len_total + 2):
            # General Number Formats for Marka Data (K, L, M)
            for col in ['K', 'L', 'M']:
                worksheet[f'{col}{i}'].number_format = '#,##0'

            # Apply Personel-specific formulas and formats ONLY for Personel rows
            if i <= len_personel + 1:
                # E: Başarı Yüzdesi (Satıcı Ciro / Mağaza Ciro)
                worksheet[f'E{i}'] = f'=IFERROR(D{i}/C{i}, 0)'
                # H: Verimlilik (F / G)
                worksheet[f'H{i}'] = f'=IFERROR(F{i}/G{i}, 0)'

                # Formatting for Personel columns
                for col in ['C', 'D', 'G']:
                    worksheet[f'{col}{i}'].number_format = '#,##0'
                
                worksheet[f'E{i}'].number_format = '0.00%'
                worksheet[f'H{i}'].number_format = '0.00%'

        # Auto-adjust column width
        for col in worksheet.columns:
            column = col[0].column_letter
            worksheet.column_dimensions[column].width = 20

    return output.getvalue()

# --- Streamlit UI ---
st.set_page_config(page_title="Sales Report", layout="wide")
st.title("Sales Report")

uploaded_file = st.file_uploader("Upload Data Excel File", type=['xlsx'])

if uploaded_file:
    if st.button("Generate"):
        try:
            with st.spinner('Processing...'):
                final_data = process_data(uploaded_file)
                st.success("Final report is ready!")
                st.download_button(
                    label="Download",
                    data=final_data,
                    file_name="Aylık Satış Raporu.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Error: {e}")
