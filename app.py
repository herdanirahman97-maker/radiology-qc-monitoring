import streamlit as st
import pandas as pd
import base64
import os

st.set_page_config(page_title="Mandaya Radiology QC & Monitoring", layout="wide")

# Function to encode image for HTML
def get_image_base64(image_path):
    if os.path.exists(image_path):
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    return ""

LOGO_PATH = "logo.png"

# Automatically find and sort all monthly Excel files (1. Januari, 2. Februari, etc.)
all_files = [f for f in os.listdir('.') if f.endswith('.xlsx') and f[0].isdigit()]
# Sort files numerically by the number at the beginning of the filename
all_files = sorted(all_files, key=lambda x: int(x.split('.')[0]))

if not all_files:
    st.error("Tidak ada file Excel (.xlsx) yang ditemukan di folder!")
else:
    st.sidebar.header("Preview Data Bulan Sebelumnya")
    
    # Dropdown 1: Select Month
    selected_file = st.sidebar.selectbox("Pilih Bulan Database:", all_files)
    
    # Load the chosen file
    xls = pd.ExcelFile(selected_file)
    
    # Find raw data sheets (ending in "Oct" for raw gform data)
    raw_data_sheets = [s for s in xls.sheet_names if str(s).endswith("Oct")]
    
    # Dropdown 2: Select Room/Modality
    selected_sheet = st.sidebar.selectbox("Pilih Ruangan / Modality:", raw_data_sheets)
    
    if selected_sheet:
        # Load raw data from the selected month and room
        df = pd.read_excel(xls, sheet_name=selected_sheet)
        
        # Identify columns dynamically
        suhu_col = [c for c in df.columns if 'Suhu' in c][0] if any('Suhu' in c for c in df.columns) else None
        kel_col = [c for c in df.columns if 'Kelembapan' in c][0] if any('Kelembapan' in c for c in df.columns) else None
        
        # Build dictionary for fast lookup by (Day, Shift)
        data_dict = {}
        if suhu_col and kel_col:
            for _, row in df.dropna(subset=['Tanggal Pengisian', 'Jadwal Dinas']).iterrows():
                try:
                    day = int(float(row['Tanggal Pengisian']))
                    shift = str(row['Jadwal Dinas']).strip().upper()
                    suhu = float(row[suhu_col]) if pd.notna(row[suhu_col]) else None
                    kel = float(row[kel_col]) if pd.notna(row[kel_col]) else None
                    nama = str(row['Nama Petugas']).strip() if pd.notna(row['Nama Petugas']) else ""
                    
                    data_dict[(day, shift)] = {'suhu': suhu, 'kel': kel, 'nama': nama}
                except:
                    continue

        # Generate Custom HTML Grid
        logo_base64 = get_image_base64(LOGO_PATH)
        img_html = f"<img src='data:image/png;base64,{logo_base64}' width='150'>" if logo_base64 else "<b>[LOGO MISSING]</b>"
        
        ruang_name = selected_sheet.replace(" Oct", "")
        # Extract month name from filename (e.g. "1. Januari.xlsx" -> "JANUARI")
        bulan_name = selected_file.split(".")[1].replace("xlsx", "").strip().upper()
        
        html = f"""
        <div style='display: flex; align-items: center; margin-bottom: 20px;'>
            <div style='flex: 0 0 auto;'>{img_html}</div>
            <div style='flex: 1 1 auto; text-align: center;'>
                <h3 style='margin: 0; font-family: sans-serif; font-weight: bold;'>Form Digital Monitoring Suhu dan Kelembapan</h3>
                <h4 style='margin: 0; font-family: sans-serif; font-weight: bold;'>Departemen Radiologi Tahun 2026</h4>
            </div>
        </div>
        <div style='font-family: sans-serif; font-weight: bold; font-size: 14px; margin-bottom: 10px;'>
            RUANG : {ruang_name}<br>
            BULAN : {bulan_name} 2026
        </div>
        
        <table style='width:100%; border-collapse: collapse; text-align: center; font-size: 11px; font-family: sans-serif;'>
        """

        # --- SUHU SECTION ---
        html += "<tr><th colspan='4' style='border: 1px solid black; background-color: #f2f2f2; font-weight: bold;'>SUHU</th>"
        html += "<th colspan='90' style='border: 1px solid black; background-color: #f2f2f2; font-weight: bold;'>Target Temperatur (18 - 23)°C</th></tr>"
        
        html += "<tr><td style='border: 1px solid black; font-weight: bold; width: 30px;'>Tgl</td>"
        for day in range(1, 32):
            bg = "#E0F7FA" if day % 2 == 0 else "#FFFFFF"
            html += f"<td colspan='3' style='border: 1px solid black; font-weight: bold; background-color: {bg};'>{day}</td>"
        html += "</tr>"
        
        html += "<tr><td style='border: 1px solid black; font-weight: bold;'>Dinas</td>"
        for day in range(1, 32):
            bg = "#E0F7FA" if day % 2 == 0 else "#FFFFFF"
            for shift in ['P', 'S', 'M']:
                html += f"<td style='border: 1px solid black; font-weight: bold; background-color: {bg}; width: 15px;'>{shift}</td>"
        html += "</tr>"
        
        # Suhu Rows (30 down to 15)
        for temp in range(30, 14, -1):
            html += "<tr>"
            label_color = "red" if temp < 18 or temp > 23 else "black"
            html += f"<td style='border: 1px solid black; font-weight: bold; color: {label_color};'>{temp}</td>"
            for day in range(1, 32):
                bg = "#E0F7FA" if day % 2 == 0 else "#FFFFFF"
                for shift in ['P', 'S', 'M']:
                    val = data_dict.get((day, shift), {}).get('suhu')
                    dot = ""
                    if val is not None and round(val) == temp:
                        dot_color = "red" if val < 18 or val > 23 else "black"
                        dot = f"<span style='color: {dot_color}; font-size: 14px;'>●</span>"
                    html += f"<td style='border: 1px solid black; background-color: {bg};'>{dot}</td>"
            html += "</tr>"
            
        # --- KELEMBAPAN SECTION ---
        html += "<tr><th colspan='4' style='border: 1px solid black; background-color: #f2f2f2; font-weight: bold;'>KELEMBAPAN</th>"
        html += "<th colspan='90' style='border: 1px solid black; background-color: #f2f2f2; font-weight: bold;'>Target kelembapan ( 40 - 60 % )</th></tr>"
        
        html += "<tr><td style='border: 1px solid black; font-weight: bold;'>Tgl</td>"
        for day in range(1, 32):
            bg = "#E0F7FA" if day % 2 == 0 else "#FFFFFF"
            html += f"<td colspan='3' style='border: 1px solid black; font-weight: bold; background-color: {bg};'>{day}</td>"
        html += "</tr>"
        
        html += "<tr><td style='border: 1px solid black; font-weight: bold;'>Dinas</td>"
        for day in range(1, 32):
            bg = "#E0F7FA" if day % 2 == 0 else "#FFFFFF"
            for shift in ['P', 'S', 'M']:
                html += f"<td style='border: 1px solid black; font-weight: bold; background-color: {bg};'>{shift}</td>"
        html += "</tr>"
        
        # Kelembapan Rows (65 down to 30, step -5)
        for hum in range(65, 29, -5):
            html += "<tr>"
            label_color = "red" if hum < 40 or hum > 60 else "black"
            html += f"<td style='border: 1px solid black; font-weight: bold; color: {label_color};'>{hum}</td>"
            for day in range(1, 32):
                bg = "#E0F7FA" if day % 2 == 0 else "#FFFFFF"
                for shift in ['P', 'S', 'M']:
                    val = data_dict.get((day, shift), {}).get('kel')
                    dot = ""
                    if val is not None and round(val / 5) * 5 == hum:
                        dot_color = "red" if val < 40 or val > 60 else "black"
                        dot = f"<span style='color: {dot_color}; font-size: 14px;'>●</span>"
                    html += f"<td style='border: 1px solid black; background-color: {bg};'>{dot}</td>"
            html += "</tr>"

        # --- NAMA PETUGAS SECTION ---
        html += "<tr><td style='border: 1px solid black; font-weight: bold; font-size: 10px;'>NAMA</td>"
        for day in range(1, 32):
            bg = "#E0F7FA" if day % 2 == 0 else "#FFFFFF"
            for shift in ['P', 'S', 'M']:
                nama = data_dict.get((day, shift), {}).get('nama', '')
                html += f"<td style='border: 1px solid black; background-color: {bg}; font-size: 9px; font-weight: bold;'>{nama}</td>"
        html += "</tr>"
        
        html += "</table>"
        
        st.markdown(html, unsafe_allow_html=True)
