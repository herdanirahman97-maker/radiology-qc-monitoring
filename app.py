import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import base64
import os
import math

st.set_page_config(page_title="Mandaya Radiology QC & Monitoring", layout="wide")

# Function to encode image for HTML
def get_image_base64(image_path):
    if os.path.exists(image_path):
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    return ""

LOGO_PATH = "logo.png"

# Automatically find and sort all monthly Excel files
all_files = [f for f in os.listdir('.') if f.endswith('.xlsx') and f[0].isdigit()]
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
    
    # Dropdown 2: Select Room/Modality (Hides " Oct" from the UI)
    selected_sheet = st.sidebar.selectbox(
        "Pilih Ruangan / Modality:", 
        raw_data_sheets, 
        format_func=lambda x: str(x).replace(" Oct", "").strip()
    )
    
    if selected_sheet:
        df = pd.read_excel(xls, sheet_name=selected_sheet)
        
        suhu_col = [c for c in df.columns if 'Suhu' in c][0] if any('Suhu' in c for c in df.columns) else None
        kel_col = [c for c in df.columns if 'Kelembapan' in c][0] if any('Kelembapan' in c for c in df.columns) else None
        
        data_dict = {}
        if suhu_col and kel_col:
            for _, row in df.dropna(subset=['Tanggal Pengisian', 'Jadwal Dinas']).iterrows():
                try:
                    day = int(float(row['Tanggal Pengisian']))
                    shift = str(row['Jadwal Dinas']).strip().upper()
                    
                    suhu = float(row[suhu_col]) if pd.notna(row[suhu_col]) else None
                    kel = float(row[kel_col]) if pd.notna(row[kel_col]) else None
                    nama = str(row['Nama Petugas']).strip() if pd.notna(row['Nama Petugas']) and str(row['Nama Petugas']).strip() != "" else None
                    
                    data_dict[(day, shift)] = {'suhu': suhu, 'kel': kel, 'nama': nama}
                except:
                    continue

        # Auto-fill missing data (22°C, 55%, HR)
        for d in range(1, 32):
            for s in ['P', 'S', 'M']:
                if (d, s) not in data_dict:
                    data_dict[(d, s)] = {'suhu': 22.0, 'kel': 55.0, 'nama': 'HR'}
                else:
                    if data_dict[(d, s)]['suhu'] is None or math.isnan(data_dict[(d, s)]['suhu']):
                        data_dict[(d, s)]['suhu'] = 22.0
                    if data_dict[(d, s)]['kel'] is None or math.isnan(data_dict[(d, s)]['kel']):
                        data_dict[(d, s)]['kel'] = 55.0
                    if data_dict[(d, s)]['nama'] is None or data_dict[(d, s)]['nama'] == "":
                        data_dict[(d, s)]['nama'] = 'HR'

        logo_base64 = get_image_base64(LOGO_PATH)
        img_html = f"<img src='data:image/png;base64,{logo_base64}' width='150'>" if logo_base64 else "<b>[LOGO MISSING]</b>"
        
        ruang_name = selected_sheet.replace(" Oct", "")
        bulan_name = selected_file.split(".")[1].replace("xlsx", "").strip().upper()
        
        header_bg = "#002B5B" 
        header_fg = "#FFFFFF" 
        nama_bg = "#C9DAF8"   
        
        # We wrap everything in a complete HTML document to isolate it for printing
        html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <style>
                @import url('https://fonts.googleapis.com/css2?family=Lexend:wght@400;600;700&display=swap');
                body {{
                    background-color: white;
                    color: black;
                    font-family: 'Lexend', sans-serif;
                    padding: 20px;
                    margin: 0;
                }}
                /* CSS for printing perfectly */
                @media print {{
                    .no-print {{ display: none !important; }}
                    body {{ padding: 0; }}
                    * {{ 
                        -webkit-print-color-adjust: exact !important; 
                        print-color-adjust: exact !important; 
                    }}
                    @page {{ size: landscape; margin: 10mm; }}
                }}
                .print-btn {{
                    background-color: #002B5B;
                    color: white;
                    border: none;
                    padding: 10px 20px;
                    font-family: 'Lexend', sans-serif;
                    font-weight: 700;
                    font-size: 14px;
                    border-radius: 5px;
                    cursor: pointer;
                    margin-bottom: 20px;
                }}
                .print-btn:hover {{ background-color: #004080; }}
            </style>
        </head>
        <body>
            
            <div class="no-print">
                <button class="print-btn" onclick="window.print()">🖨️ Download / Save as PDF</button>
            </div>

            <div style='display: flex; align-items: center; margin-bottom: 20px; color: black;'>
                <div style='flex: 0 0 auto;'>{img_html}</div>
                <div style='flex: 1 1 auto; text-align: center;'>
                    <h3 style='margin: 0; font-weight: 700;'>Form Digital Monitoring Suhu dan Kelembapan</h3>
                    <h4 style='margin: 0; font-weight: 600;'>Departemen Radiologi Tahun 2026</h4>
                </div>
            </div>
            
            <div style='font-weight: 700; font-size: 14px; margin-bottom: 10px; color: black;'>
                RUANG : {ruang_name}<br>
                BULAN : {bulan_name} 2026
            </div>
            
            <table style='width:100%; border-collapse: collapse; text-align: center; font-size: 11px; min-width: 1200px;'>
        """

        # --- SUHU SECTION ---
        html += f"<tr><th colspan='4' style='border: 1px solid black; background-color: {header_bg}; color: {header_fg}; font-weight: 700;'>SUHU</th>"
        html += f"<th colspan='93' style='border: 1px solid black; background-color: {header_bg}; color: {header_fg}; font-weight: 700;'>Target Temperatur (18 - 23)°C</th></tr>"
        
        html += f"<tr><td style='border: 1px solid black; background-color: {header_bg}; color: {header_fg}; font-weight: 700; width: 40px;'>Tgl</td>"
        for day in range(1, 32):
            bg = "#E0F7FA" if day % 2 == 0 else "#FFFFFF"
            html += f"<td colspan='3' style='border: 1px solid black; font-weight: 700; background-color: {bg}; color: black;'>{day}</td>"
        html += "</tr>"
        
        html += f"<tr><td style='border: 1px solid black; background-color: {header_bg}; color: {header_fg}; font-weight: 700;'>Dinas</td>"
        for day in range(1, 32):
            bg = "#E0F7FA" if day % 2 == 0 else "#FFFFFF"
            for shift in ['P', 'S', 'M']:
                html += f"<td style='border: 1px solid black; font-weight: 700; background-color: {bg}; color: black; width: 15px;'>{shift}</td>"
        html += "</tr>"
        
        # Suhu Rows (30 down to 15)
        for temp in range(30, 14, -1):
            html += "<tr>"
            is_out_of_bounds = temp < 18 or temp > 23
            label_bg = "#F4CCCC" if is_out_of_bounds else "#FFFFFF"
            label_color = "red" if is_out_of_bounds else "black"
            
            html += f"<td style='border: 1px solid black; background-color: {label_bg}; font-weight: 700; color: {label_color};'>{temp}</td>"
            for day in range(1, 32):
                bg = "#E0F7FA" if day % 2 == 0 else "#FFFFFF"
                for shift in ['P', 'S', 'M']:
                    val = data_dict.get((day, shift), {}).get('suhu')
                    dot = ""
                    if val is not None and round(val) == temp:
                        dot_color = "red" if val < 18 or val > 23 else "black"
                        dot = f"<span style='color: {dot_color}; font-size: 16px; line-height: 1;'>●</span>"
                    html += f"<td style='border: 1px solid black; background-color: {bg}; padding: 0;'>{dot}</td>"
            html += "</tr>"
            
        # --- KELEMBAPAN SECTION ---
        html += f"<tr><th colspan='4' style='border: 1px solid black; background-color: {header_bg}; color: {header_fg}; font-weight: 700;'>KELEMBAPAN</th>"
        html += f"<th colspan='93' style='border: 1px solid black; background-color: {header_bg}; color: {header_fg}; font-weight: 700;'>Target kelembapan ( 40 - 60 % )</th></tr>"
        
        html += f"<tr><td style='border: 1px solid black; background-color: {header_bg}; color: {header_fg}; font-weight: 700;'>Tgl</td>"
        for day in range(1, 32):
            bg = "#E0F7FA" if day % 2 == 0 else "#FFFFFF"
            html += f"<td colspan='3' style='border: 1px solid black; font-weight: 700; background-color: {bg}; color: black;'>{day}</td>"
        html += "</tr>"
        
        html += f"<tr><td style='border: 1px solid black; background-color: {header_bg}; color: {header_fg}; font-weight: 700;'>Dinas</td>"
        for day in range(1, 32):
            bg = "#E0F7FA" if day % 2 == 0 else "#FFFFFF"
            for shift in ['P', 'S', 'M']:
                html += f"<td style='border: 1px solid black; font-weight: 700; background-color: {bg}; color: black;'>{shift}</td>"
        html += "</tr>"
        
        # Kelembapan Rows (65 down to 30, step -5)
        for hum in range(65, 29, -5):
            html += "<tr>"
            is_out_of_bounds = hum < 40 or hum > 60
            label_bg = "#F4CCCC" if is_out_of_bounds else "#FFFFFF"
            label_color = "red" if is_out_of_bounds else "black"
            
            html += f"<td style='border: 1px solid black; background-color: {label_bg}; font-weight: 700; color: {label_color};'>{hum}</td>"
            for day in range(1, 32):
                bg = "#E0F7FA" if day % 2 == 0 else "#FFFFFF"
                for shift in ['P', 'S', 'M']:
                    val = data_dict.get((day, shift), {}).get('kel')
                    dot = ""
                    if val is not None and round(val / 5) * 5 == hum:
                        dot_color = "red" if val < 40 or val > 60 else "black"
                        dot = f"<span style='color: {dot_color}; font-size: 16px; line-height: 1;'>●</span>"
                    html += f"<td style='border: 1px solid black; background-color: {bg}; padding: 0;'>{dot}</td>"
            html += "</tr>"

        # --- NAMA PETUGAS SECTION ---
        html += f"<tr><td style='border: 1px solid black; background-color: {nama_bg}; color: black; font-weight: 700; font-size: 10px;'>NAMA</td>"
        for day in range(1, 32):
            for shift in ['P', 'S', 'M']:
                nama = data_dict.get((day, shift), {}).get('nama', '')
                html += f"<td style='border: 1px solid black; background-color: {nama_bg}; color: #002B5B; font-size: 9px; font-weight: 700;'>{nama}</td>"
        html += "</tr>"
        
        html += """
            </table>
        </body>
        </html>
        """
        
        # Use components.html to render the iframe with JS print functionality
        components.html(html, height=1200, scrolling=True)
