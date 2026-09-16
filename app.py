import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import base64
import os
import math

st.set_page_config(page_title="Mandaya Radiology QC & Monitoring", layout="wide")

def get_image_base64(image_path):
    if os.path.exists(image_path):
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode()
    return ""

LOGO_PATH = "logo.png"

all_files = [f for f in os.listdir('.') if f.endswith('.xlsx') and f[0].isdigit()]
all_files = sorted(all_files, key=lambda x: int(x.split('.')[0]))

# --- PHASE 1: TEMPERATURE & HUMIDITY HTML BUILDER ---
def build_room_html(sheet_name, file_name, df, logo_base64):
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

    img_html = f"<img src='data:image/png;base64,{logo_base64}' width='150'>" if logo_base64 else "<b>[LOGO MISSING]</b>"
    ruang_name = sheet_name.replace(" Oct", "")
    bulan_name = file_name.split(".")[1].replace("xlsx", "").strip().upper()
    
    header_bg, header_fg, nama_bg = "#002B5B", "#FFFFFF", "#C9DAF8"
    
    html = f"""
    <div class="page-container">
        <div style='display: flex; align-items: center; justify-content: space-between; margin-bottom: 20px; color: black; border-bottom: 2px solid #002B5B; padding-bottom: 15px;'>
            <div style='flex: 0 0 auto;'>{img_html}</div>
            <div style='flex: 1 1 auto; text-align: center;'>
                <h1 style='margin: 0; font-weight: 700; font-size: 26px; letter-spacing: 0.5px;'>Form Digital Monitoring Suhu dan Kelembapan</h1>
                <h3 style='margin: 6px 0 0 0; font-weight: 600; font-size: 16px; color: #333;'>Departemen Radiologi Tahun 2026</h3>
            </div>
            <div style='flex: 0 0 150px;'></div>
        </div>
        
        <div style='font-weight: 700; font-size: 14px; margin-bottom: 10px; color: black;'>
            RUANG : {ruang_name}<br>
            BULAN : {bulan_name} 2026
        </div>
        
        <table style='width:100%; border-collapse: collapse; text-align: center; font-size: 11px; min-width: 1100px;'>
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
    
    for temp in range(30, 14, -1):
        html += "<tr>"
        is_out = temp < 18 or temp > 23
        label_bg = "#F4CCCC" if is_out else "#FFFFFF"
        label_color = "red" if is_out else "black"
        
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
    
    for hum in range(65, 29, -5):
        html += "<tr>"
        is_out = hum < 40 or hum > 60
        label_bg = "#F4CCCC" if is_out else "#FFFFFF"
        label_color = "red" if is_out else "black"
        
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
    
    html += f"""
        </table>
        <div class="trademark">Source & maintained by Herdani Rahman Account</div>
    </div>
    """
    return html

# --- PHASE 2: QC FORM HTML RENDERER ---
def build_qc_html(sheet_name, file_name, df, logo_base64):
    img_html = f"<img src='data:image/png;base64,{logo_base64}' width='150'>" if logo_base64 else "<b>[LOGO MISSING]</b>"
    bulan_name = file_name.split(".")[1].replace("xlsx", "").strip().upper()
    
    # Convert Excel QC sheet directly into an elegant styled HTML table representation
    df_clean = df.fillna("")
    table_html = df_clean.to_html(index=False, header=False, border=0, classes="qc-table")
    
    html = f"""
    <div class="page-container">
        <div style='display: flex; align-items: center; justify-content: space-between; margin-bottom: 15px; color: black; border-bottom: 2px solid #002B5B; padding-bottom: 10px;'>
            <div style='flex: 0 0 auto;'>{img_html}</div>
            <div style='flex: 1 1 auto; text-align: center;'>
                <h1 style='margin: 0; font-weight: 700; font-size: 24px;'>Form Digital Daily Quality Control</h1>
                <h3 style='margin: 4px 0 0 0; font-weight: 600; font-size: 15px; color: #333;'>Departemen Radiologi Tahun 2026</h3>
            </div>
            <div style='flex: 0 0 150px;'></div>
        </div>
        
        <div style='font-weight: 700; font-size: 14px; margin-bottom: 10px; color: black;'>
            MODALITY / SHEET : {sheet_name}<br>
            BULAN : {bulan_name} 2026
        </div>
        
        <div class="table-wrapper">
            {table_html}
        </div>
        
        <div class="trademark">Source & maintained by Herdani Rahman</div>
    </div>
    """
    return html


# --- MAIN APP LAYOUT & TABS ---
if not all_files:
    st.error("Tidak ada file Excel (.xlsx) yang ditemukan di folder!")
else:
    st.sidebar.header("Navigasi Menu Utama")
    app_mode = st.sidebar.radio("Pilih Modul:", ["🌡️ Monitoring Suhu & Kelembapan", "📋 Daily Quality Control (QC)"])
    
    selected_file = st.sidebar.selectbox("Pilih Bulan Database:", all_files)
    xls = pd.ExcelFile(selected_file)
    
    logo_b64 = get_image_base64(LOGO_PATH)
    
    master_html_start = """
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Lexend:wght@400;600;700&display=swap');
            body {
                background-color: white;
                color: black;
                font-family: 'Lexend', sans-serif;
                margin: 0;
                display: flex;
                flex-direction: column;
                align-items: center;
                justify-content: center;
            }
            .page-container {
                width: 100%;
                max-width: 1350px;
                background: white;
                padding: 20px;
                box-sizing: border-box;
                margin: auto;
            }
            .qc-table {
                width: 100%;
                border-collapse: collapse;
                font-size: 10px;
            }
            .qc-table td, .qc-table th {
                border: 1px solid black;
                padding: 4px;
                text-align: center;
            }
            @media print {
                .no-print { display: none !important; }
                @page { size: landscape; margin: 0; }
                body { padding: 10mm; display: block; }
                .page-container { margin: 0 auto; width: 100%; max-width: none; }
                .page-break { page-break-after: always; }
                * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
            }
            .print-btn {
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
            }
            .print-btn:hover { background-color: #004080; }
            .trademark {
                text-align: left; 
                font-size: 11px; 
                color: #888888; 
                margin-top: 10px; 
                font-weight: 400; 
                font-style: italic;
                letter-spacing: 0.3px;
            }
        </style>
    </head>
    <body>
        <div class="no-print">
            <button class="print-btn" onclick="window.print()">🖨️ Download / Save as PDF</button>
        </div>
    """
    master_html_end = "</body></html>"
    
    if app_mode == "🌡️ Monitoring Suhu & Kelembapan":
        raw_data_sheets = [s for s in xls.sheet_names if str(s).endswith("Oct")]
        selected_sheet = st.sidebar.selectbox(
            "Pilih Ruangan / Modality:", 
            raw_data_sheets, 
            format_func=lambda x: str(x).replace(" Oct", "").strip()
        )
        view_mode = st.sidebar.radio("Mode Tampilan:", ["Tampilkan 1 Ruangan", "Cetak Semua Ruangan (1 Bulan)"])
        
        if view_mode == "Tampilkan 1 Ruangan" and selected_sheet:
            df = pd.read_excel(xls, sheet_name=selected_sheet)
            room_html = build_room_html(selected_sheet, selected_file, df, logo_b64)
            components.html(master_html_start + room_html + master_html_end, height=850, scrolling=True)
            
        elif view_mode == "Cetak Semua Ruangan (1 Bulan)":
            st.info("💡 Memuat semua ruangan suhu & kelembapan untuk dicetak.")
            all_rooms_html = ""
            for i, sheet in enumerate(raw_data_sheets):
                df = pd.read_excel(xls, sheet_name=sheet)
                all_rooms_html += build_room_html(sheet, selected_file, df, logo_b64)
                if i < len(raw_data_sheets) - 1:
                    all_rooms_html += "<div class='page-break'></div>"
            components.html(master_html_start + all_rooms_html + master_html_end, height=850, scrolling=True)
            
    elif app_mode == "📋 Daily Quality Control (QC)":
        qc_sheets = [s for s in xls.sheet_names if "QC" in s]
        selected_qc = st.sidebar.selectbox("Pilih Lembar QC Modality:", qc_sheets)
        qc_view_mode = st.sidebar.radio("Mode Tampilan QC:", ["Tampilkan 1 QC Sheet", "Cetak Semua QC Sheets (1 Bulan)"])
        
        if qc_view_mode == "Tampilkan 1 QC Sheet" and selected_qc:
            df_qc = pd.read_excel(xls, sheet_name=selected_qc, header=None)
            qc_html = build_qc_html(selected_qc, selected_file, df_qc, logo_b64)
            components.html(master_html_start + qc_html + master_html_end, height=900, scrolling=True)
            
        elif qc_view_mode == "Cetak Semua QC Sheets (1 Bulan)":
            st.info("💡 Memuat seluruh lembar QC modality untuk dicetak.")
            all_qc_html = ""
            for i, qsheet in enumerate(qc_sheets):
                df_qc = pd.read_excel(xls, sheet_name=qsheet, header=None)
                all_qc_html += build_qc_html(qsheet, selected_file, df_qc, logo_b64)
                if i < len(qc_sheets) - 1:
                    all_qc_html += "<div class='page-break'></div>"
            components.html(master_html_start + all_qc_html + master_html_end, height=900, scrolling=True)
