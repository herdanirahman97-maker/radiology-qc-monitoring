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
SIG1_PATH = "sig1.png"
SIG2_PATH = "sig2.png"
SIG3_PATH = "sig3.png"

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
                <h1 style='margin: 0; font-weight: 700; font-size: 26px;'>Form Digital Monitoring Suhu dan Kelembapan</h1>
                <h3 style='margin: 6px 0 0 0; font-weight: 600; font-size: 16px; color: #333;'>Departemen Radiologi Tahun 2026</h3>
            </div>
            <div style='flex: 0 0 150px;'></div>
        </div>
        
        <div style='font-weight: 700; font-size: 14px; margin-bottom: 10px; color: black;'>
            RUANG : {ruang_name}<br>
            BULAN : {bulan_name} 2026
        </div>
        
        <table style='width:100%; border-collapse: collapse; text-align: center; font-size: 11px;'>
    """

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

    html += f"<tr><td style='border: 1px solid black; background-color: {nama_bg}; color: black; font-weight: 700; font-size: 10px;'>NAMA</td>"
    for day in range(1, 32):
        for shift in ['P', 'S', 'M']:
            nama = data_dict.get((day, shift), {}).get('nama', '')
            html += f"<td style='border: 1px solid black; background-color: {nama_bg}; color: #002B5B; font-size: 9px; font-weight: 700;'>{nama}</td>"
    html += "</tr>"
    
    html += f"""
        </table>
        <div class="trademark">Source & maintained by Herdani Rahman</div>
    </div>
    """
    return html

# --- PHASE 2: DYNAMIC QC FORM BUILDER ---
def build_qc_html(sheet_name, file_name, df, logo_base64, sig1_b64, sig2_b64, sig3_b64):
    img_html = f"<img src='data:image/png;base64,{logo_base64}' width='150'>" if logo_base64 else "<b>[LOGO MISSING]</b>"
    bulan_name = file_name.split(".")[1].replace("xlsx", "").strip().upper()
    
    modality_title = sheet_name.replace("QC ", "")
    
    # Automatically locate the row index for "NAMA PEKERJA RADIASI"
    name_row_idx = None
    for idx, row in df.iterrows():
        row_str = str(row.values)
        if "NAMA" in row_str and "RADIASI" in row_str:
            name_row_idx = idx
            break
            
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
        
        <div style='font-weight: 700; font-size: 13px; margin-bottom: 10px; color: black;'>
            MODALITAS : {modality_title}<br>
            BULAN : {bulan_name} 2026
        </div>
    """
    
    table_html = "<table class='qc-table'>"
    
    for idx, row in df.iterrows():
        if idx < 8:
            continue
            
        row_vals = [str(x) if pd.notna(x) else "" for x in row.values]
        
        is_category = (row_vals[1] == "" and row_vals[4] == "" and row_vals[0] != "" and row_vals[0] != "NO")
        is_header = (row_vals[0] == "NO")
        is_names = (idx == name_row_idx) or ("NAMA PEKERJA RADIASI" in str(row.values))
        
        if name_row_idx is not None and idx > name_row_idx:
            # Stop after rendering the name row and signature block area
            break
            
        if is_header:
            table_html += "<tr style='background-color: #002B5B; color: white; font-weight: bold;'>"
            table_html += "<td style='border: 1px solid black; padding: 6px; text-align: center;' width='45px'>NO</td>"
            table_html += "<td colspan='2' style='border: 1px solid black; padding: 6px; text-align: center;' width='200px'>KEGIATAN</td>"
            table_html += "<td colspan='2' style='border: 1px solid black; padding: 6px; text-align: center;' width='250px'>PARAMETER</td>"
            for d in range(1, 32):
                table_html += f"<td style='border: 1px solid black; padding: 6px; text-align: center;' width='40px'>{d}</td>"
            table_html += "</tr>"
        elif is_category:
            table_html += f"<tr style='background-color: #cfe2f3; font-weight: bold; text-align: left;'><td colspan='36' style='border: 1px solid black; padding: 8px; color: #000;'>{row_vals[0]}</td></tr>"
        elif is_names:
            table_html += "<tr style='background-color: #C9DAF8; font-weight: bold; height: 32px;'>"
            table_html += "<td colspan='5' style='border: 1px solid black; padding: 5px; text-align: center;'>NAMA PEKERJA RADIASI</td>"
            for day_idx in range(1, 32):
                # Search across all columns in this row for the day's value or map directly by offset
                val = ""
                for col_c in range(len(row_vals)):
                    # Check if cell matches day number or if we can pull from mapped position
                    pass
                # Fallback extraction from expected index offset
                val = row_vals[day_idx + 5] if (day_idx + 5) < len(row_vals) else ""
                if val == "" and (day_idx + 6) < len(row_vals):
                    val = row_vals[day_idx + 6]
                table_html += f"<td style='border: 1px solid black; padding: 4px; font-size: 9px; text-align: center;'>{val}</td>"
            table_html += "</tr>"
        else:
            table_html += "<tr style='height: 30px;'>"
            table_html += f"<td style='border: 1px solid black; padding: 5px; text-align: center;' width='45px'>{row_vals[0]}</td>"
            table_html += f"<td colspan='2' style='border: 1px solid black; padding: 5px; text-align: left;' width='200px'>{row_vals[1] if row_vals[1] != '' else row_vals[2]}</td>"
            table_html += f"<td colspan='2' style='border: 1px solid black; padding: 5px; text-align: left;' width='250px'>{row_vals[4] if row_vals[4] != '' else row_vals[5]}</td>"
            
            for day_idx in range(1, 32):
                val = row_vals[day_idx + 5] if (day_idx + 5) < len(row_vals) else ""
                table_html += f"<td style='border: 1px solid black; padding: 4px; text-align: center;'>{val}</td>"
            table_html += "</tr>"
            
    table_html += "</table>"
    html += table_html
    
    sig1_img = f"<img src='data:image/png;base64,{sig1_b64}' width='110'>" if sig1_b64 else "<br><br>"
    sig2_img = f"<img src='data:image/png;base64,{sig2_b64}' width='110'>" if sig2_b64 else "<br><br>"
    sig3_img = f"<img src='data:image/png;base64,{sig3_b64}' width='110'>" if sig3_b64 else "<br><br>"
    
    html += f"""
    <div style="display: flex; justify-content: space-between; margin-top: 30px; align-items: flex-start; width: 100%;">
        <div style="border: 1px solid black; width: 300px; font-size: 10px;">
            <div style="background-color: #d9d9d9; padding: 6px; font-weight: bold; border-bottom: 1px solid black;">Keterangan :</div>
            <div style="padding: 8px; line-height: 1.4;">
                ✓ : Memenuhi / Lulus Uji / Lengkap<br>
                X : Tidak Memenuhi / Tidak Lulus Uji / Tidak Lengkap
            </div>
        </div>
        
        <table style="border-collapse: collapse; width: 550px; font-size: 10px; text-align: center;">
            <tr>
                <td style="border: 1px solid black; background-color: #d9d9d9; font-weight: bold; padding: 5px;" width="33%">Disiapkan Oleh</td>
                <td style="border: 1px solid black; background-color: #d9d9d9; font-weight: bold; padding: 5px;" colspan="2">Mengetahui</td>
            </tr>
            <tr>
                <td style="border: 1px solid black; height: 60px; vertical-align: middle;">{sig1_img}</td>
                <td style="border: 1px solid black; height: 60px; vertical-align: middle;" width="33%">{sig2_img}</td>
                <td style="border: 1px solid black; height: 60px; vertical-align: middle;" width="33%">{sig3_img}</td>
            </tr>
            <tr>
                <td style="border: 1px solid black; font-weight: bold; padding: 4px;">Rheinner Nicholaus, Amd. Rad</td>
                <td style="border: 1px solid black; font-weight: bold; padding: 4px;">Joko Harjanto S.Tr.Rad</td>
                <td style="border: 1px solid black; font-weight: bold; padding: 4px;">dr Christopher Silman Sp.Rad, Ph.D</td>
            </tr>
            <tr>
                <td style="border: 1px solid black; padding: 4px;">PIC. Fasilitas</td>
                <td style="border: 1px solid black; padding: 4px;">Koordinator Radiologi</td>
                <td style="border: 1px solid black; padding: 4px;">Ho. Dept Radiologi</td>
            </tr>
        </table>
    </div>
    <div class="trademark">Source & maintained by Herdani Rahman</div>
    </div>
    """
    return html

# --- MAIN APP LAYOUT ---
if not all_files:
    st.error("Tidak ada file Excel (.xlsx) yang ditemukan di folder!")
else:
    st.sidebar.header("Navigasi Menu Utama")
    app_mode = st.sidebar.radio("Pilih Modul:", ["🌡️ Monitoring Suhu & Kelembapan", "📋 Daily Quality Control (QC)"])
    
    selected_file = st.sidebar.selectbox("Pilih Bulan Database:", all_files)
    xls = pd.ExcelFile(selected_file)
    
    logo_b64 = get_image_base64(LOGO_PATH)
    sig1_b64 = get_image_base64(SIG1_PATH)
    sig2_b64 = get_image_base64(SIG2_PATH)
    sig3_b64 = get_image_base64(SIG3_PATH)
    
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
                overflow-x: auto;
            }
            .page-container {
                width: 1450px;
                background: white;
                padding: 20px;
                box-sizing: border-box;
                margin: 0 auto;
            }
            .qc-table {
                width: 100%;
                border-collapse: collapse;
                font-size: 10px;
                table-layout: fixed;
            }
            .qc-table td {
                border: 1px solid black;
                overflow: hidden;
            }
            @media print {
                .no-print { display: none !important; }
                @page { size: landscape; margin: 0; }
                body { padding: 10mm; overflow: visible !important; }
                .page-container { margin: 0 auto; width: 100%; }
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
                margin: 15px 0 15px 20px;
            }
            .print-btn:hover { background-color: #004080; }
            .trademark {
                text-align: left; 
                font-size: 11px; 
                color: #888888; 
                margin-top: 15px; 
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
            qc_html = build_qc_html(selected_qc, selected_file, df_qc, logo_b64, sig1_b64, sig2_b64, sig3_b64)
            components.html(master_html_start + qc_html + master_html_end, height=950, scrolling=True)
            
        elif qc_view_mode == "Cetak Semua QC Sheets (1 Bulan)":
            st.info("💡 Memuat seluruh lembar QC modality untuk dicetak.")
            all_qc_html = ""
            for i, qsheet in enumerate(qc_sheets):
                df_qc = pd.read_excel(xls, sheet_name=qsheet, header=None)
                all_qc_html += build_qc_html(qsheet, selected_file, df_qc, logo_b64, sig1_b64, sig2_b64, sig3_b64)
                if i < len(qc_sheets) - 1:
                    all_qc_html += "<div class='page-break'></div>"
            components.html(master_html_start + all_qc_html + master_html_end, height=950, scrolling=True)
