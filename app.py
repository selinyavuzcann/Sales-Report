import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import PatternFill, Font, Alignment

def process_data(data_file):
    # Load the data sheets
    df_personel = pd.read_excel(data_file, sheet_name='PERSONEL SATIŞ AYLIK')
    df_marka = pd.read_excel(data_file, sheet_name='MARKA SATIŞ')
    
    # Define the final template structure (A through M)
    output_df = pd.DataFrame(columns=[
        'Satıcı Ad Soyad',   # A
        'Mağaza Lokasyon',   # B
        'Mağaza Ciro',       # C
        'Satıcı Ciro',       # D
        'Başarı Yüzdesi',    # E
        'Çalışma Saati',     # F
        'Satılan Adet',      # G
        'Verimlilik',        # H
        'MARKA_LOKASYON',    # I
        'MODEL ADI',         # J
        'ADET',              # K
        'CİRO',              # L
        'STOK'               # M
    ])

    # 1. Mapping PERSONEL SATIŞ AYLIK
    output_df['Satıcı Ad Soyad'] = df_personel['SATICI']
    output_df['Mağaza Lokasyon'] = df_personel['LOKASYON']
    output_df['Satıcı Ciro'] = df_personel['CİRO']
    output_df['Satılan Adet'] = df_personel['ADET']

    # 2. Mapping MARKA SATIŞ Summary for Store Ciro (Column C)
    marka_ciro_map = df_marka.groupby('LOKASYON')['CİRO'].sum().to_dict()
    output_df['Mağaza Ciro'] = output_df['Mağaza Lokasyon'].map(marka_ciro_map)

    # 3. Mapping MARKA SATIŞ Granular Data (Columns I-M)
    output_df['MARKA_LOKASYON'] = df_marka['LOKASYON']
    output_df['MODEL ADI'] = df_marka['MODEL ADI']
    output_df['ADET'] = df_marka['ADET']
    output_df['CİRO'] = df_marka['CİRO']
    output_df['STOK'] = df_marka['STOK']

    # 4. Create Excel with Professional Formatting & Colors
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        output_df.to_excel(writer, sheet_name='Sales Report', index=False)
        
        # Add reference sheets
        df_personel.to_excel(writer, sheet_name='PERSONEL SATIŞ AYLIK', index=False)
        df_marka.to_excel(writer, sheet_name='MARKA SATIŞ', index=False)
        
        workbook = writer.book
        worksheet = workbook['Sales Report']
        
        # Define Professional Styles
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid") # Dark Blue
        header_font = Font(color="FFFFFF", bold=True) # White Bold Text
        center_align = Alignment(horizontal="center")

        # Apply Header Styling
        for cell in worksheet[1]: # First row
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align

        # Professional Formatting Loop for Data
        for i in range(2, len(output_df) + 2):
            # --- FORMULAS ---
            # E: Başarı Yüzdesi (C/D)
            worksheet[f'E{i}'] = f'=IFERROR(C{i}/D{i}, 0)'
            # H: Verimlilik (F/G)
            worksheet[f'H{i}'] = f'=IFERROR(F{i}/G{i}, "")'

            # --- NUMBER FORMATS ---
            # Ciros: Whole numbers with thousands separator
            worksheet[f'C{i}'].number_format = '#,##0'
            worksheet[f'D{i}'].number_format = '#,##0'
            worksheet[f'L{i}'].number_format = '#,##0'
            
            # Column E & Column H: Both in Percentage Format as requested
            worksheet[f'E{i}'].number_format = '0.0%'
            worksheet[f'H{i}'].number_format = '0.0%'
            
            # Stocks and Quantities: Whole numbers
            worksheet[f'G{i}'].number_format = '#,##0'
            worksheet[f'K{i}'].number_format = '#,##0'
            worksheet[f'M{i}'].number_format = '#,##0'

    return output.getvalue()

# --- Streamlit UI ---
st.set_page_config(page_title="Pro Sales Report", layout="wide")

st.title("Sales Report")


uploaded_file = st.file_uploader("Upload Data Excel File", type=['xlsx'])

if uploaded_file:
    if st.button("Generate"):
        try:
            with st.spinner('Generating...'):
                final_data = process_data(uploaded_file)
                
              
                st.download_button(
                    label="📥 Download",
                    data=final_data,
                    file_name="Satış Raporu.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Error encountered: {e}")
            st.info("Check sheet names: 'PERSONEL SATIŞ AYLIK' and 'MARKA SATIŞ'")
