import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import base64
import os
import math

st.set_page_config(page_title="Mandaya Radiology QC & Monitoring", layout="wide")

def get_image_base64(base_name):
    possible_names = [f"{base_name}.png", f"{base_name}.jpg", f"{base_name}.jpeg", base_name]
    for file_name in possible_names:
        if os.path.exists(file_name):
            with open(file_name, "rb") as img_file:
                mime_type = "image/png" if file_name.endswith(".png") else "image/jpeg"
                b64_str = base64.b64encode(img_file.read()).decode()
                return f"data:{mime_type};base64,{b64_str}"
    return ""

LOGO_PATH = "logo"
SIG1_PATH = "sig1"
SIG2_PATH = "sig2"
SIG3_PATH = "sig3"

all_files = [f for f in os.listdir('.') if f.endswith('.xlsx') and f[0].isdigit()]
all_files = sorted(all_files, key=lambda x: int(x.split('.')[0]))

# --- PHASE 1: TEMPERATURE & HUMIDITY HTML BUILDER ---
def build_room_html(sheet_name, file_name, df, logo_data_uri):
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
                raw_nama = str(row['Nama Petugas']).strip() if pd.notna(row['Nama Petugas']) else "HR"
                nama = "HR" if "#REF!" in raw_nama or raw_nama == "" else raw_nama
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
                if data_dict[(d, s)]['nama'] is None or data_dict[(d, s)]['nama'] == "" or "#REF!" in str(data_dict[(d, s)]['nama']):
                    data_dict[(d, s)]['nama'] = 'HR'

    img_html = f"<img src='{logo_data_uri}' width='150'>" if logo_data_uri else "<b>[LOGO MISSING]</b>"
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
            raw_nama = data_dict.get((day, shift), {}).get('nama', 'HR')
            nama = "HR" if "#REF!" in raw_nama or raw_nama == "" else raw_nama
            html += f"<td style='border: 1px solid black; background-color: {nama_bg}; color: #002B5B; font-size: 9px; font-weight: 700;'>{nama}</td>"
    html += "</tr>"
    
    html += f"""
        </table>
        <div class="trademark">Source & maintained by Herdani Rahman Account</div>
    </div>
    """
    return html

# --- PHASE 2: DYNAMIC COLUMN MAPPING QC FORM BUILDER ---
def build_qc_html(sheet_name, file_name, df, logo_data_uri, sig1_uri, sig2_uri, sig3_uri):
    img_html = f"<img src='{logo_data_uri}' width='150'>" if logo_data_uri else "<b>[LOGO MISSING]</b>"
    bulan_name = file_name.split(".")[1].replace("xlsx", "").strip().upper()
    modality_title = sheet_name.replace("QC ", "")
    
    # 1. Dynamically find the parameter column index from header row (index 8)
    param_col_idx = 5  # default fallback
    for r_idx in [8, 7, 6]:
        if r_idx < len(df):
            row_h = [str(x).upper() for x in df.iloc[r_idx].values]
            for c_idx, val in enumerate(row_h):
                if "PARAMETER" in val:
                    param_col_idx = c_idx
                    break
            if param_col_idx != 5:
                break

    name_row_vals = None
    for idx, row in df.iterrows():
        row_str = str(row.values)
        if "NAMA" in row_str and "RADIASI" in row_str:
            name_row_vals = [str(x) if pd.notna(x) else "" for x in row.values]
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
        row_text_joined = " ".join(row_vals).upper()
        
        if idx >= 42 or any(kwd in row_text_joined for kwd in ["DISIAPKAN", "MENGETAHUI", "RHEINNER", "JOKO HARJANTO", "CHRISTOPHER", "PIC.", "KOORDINATOR", "HO. DEPT"]):
            continue
            
        is_header = ("NO" in row_vals[0].upper() or "KEGIATAN" in row_text_joined)
        is_names = ("NAMA PEKERJA RADIASI" in row_text_joined)
        
        non_empty_texts = [v.strip() for i, v in enumerate(row_vals[:param_col_idx]) if v.strip() != "" and v.strip() != "NO"]
        is_category = (len(non_empty_texts) == 1 and not is_names and not is_header and row_vals[0] != "")
        
        if is_header:
            table_html += "<tr style='background-color: #002B5B; color: white; font-weight: bold;'>"
            table_html += "<td style='border: 1px solid black; padding: 6px; text-align: center;' width='45px'>NO</td>"
            table_html += "<td colspan='2' style='border: 1px solid black; padding: 6px; text-align: center;' width='200px'>KEGIATAN</td>"
            table_html += "<td colspan='2' style='border: 1px solid black; padding: 6px; text-align: center;' width='250px'>PARAMETER</td>"
            for d in range(1, 32):
                table_html += f"<td style='border: 1px solid black; padding: 6px; text-align: center;' width='40px'>{d}</td>"
            table_html += "</tr>"
        elif is_category:
            cat_text = non_empty_texts[0]
            table_html += f"<tr style='background-color: #cfe2f3; font-weight: bold; text-align: left;'><td colspan='36' style='border: 1px solid black; padding: 8px; color: #000;'>{cat_text}</td></tr>"
        elif is_names:
            continue
        else:
            if row_vals[0] == "" and all(v == "" for v in row_vals[1:param_col_idx]):
                continue
                
            # Extract Kegiatan (col 1) and Parameter (dynamically at param_col_idx)
            kegiatan = row_vals[1].strip() if len(row_vals) > 1 else ""
            parameter = row_vals[param_col_idx].strip() if param_col_idx < len(row_vals) else ""
            
            # Fallback if parameter is empty
            if not parameter:
                for c in range(2, param_col_idx + 1):
                    if c < len(row_vals) and row_vals[c].strip() != "" and row_vals[c].strip() not in ["✓", "X"]:
                        parameter = row_vals[c].strip()
                        break
            
            table_html += "<tr style='height: 30px;'>"
            table_html += f"<td style='border: 1px solid black; padding: 5px; text-align: center;' width='45px'>{row_vals[0]}</td>"
            table_html += f"<td colspan='2' style='border: 1px solid black; padding: 5px; text-align: left;' width='200px'>{kegiatan}</td>"
            table_html += f"<td colspan='2' style='border: 1px solid black; padding: 5px; text-align: left;' width='250px'>{parameter}</td>"
            
            # Days 1 to 31 checklist values (starting right after parameter column)
            for day_idx in range(1, 32):
                val = ""
                target_col = param_col_idx + day_idx
                if target_col < len(row_vals):
                    cand = row_vals[target_col].strip()
                    if cand in ["✓", "X"] or (len(cand) <= 2 and cand != ""):
                        val = cand
                if not val or val == "" or "#REF!" in val:
                    val = "✓"
                table_html += f"<td style='border: 1px solid black; padding: 4px; text-align: center;'>{val}</td>"
            table_html += "</tr>"
            
    # --- RENDER NAMA PEKERJA RADIASI AT THE ABSOLUTE BOTTOM ---
    table_html += "<tr style='background-color: #C9DAF8; font-weight: bold; height: 32px;'>"
    table_html += "<td colspan='5' style='border: 1px solid black; padding: 5px; text-align: center;'>NAMA PEKERJA RADIASI</td>"
    for day_idx in range(1, 32):
        val = ""
        if name_row_vals:
            target_col = param_col_idx + day_idx
            if target_col < len(name_row_vals):
                cand = name_row_vals[target_col].strip()
                if cand and cand != "NAMA PEKERJA RADIASI" and not any(bad in cand.upper() for bad in ["JOKO", "CHRISTOPHER", "RHEINNER"]):
                    val = cand
        if not val or "#REF!" in val:
            val = "HR"
        table_html += f"<td style='border: 1px solid black; padding: 4px; font-size: 9px; text-align: center;'>{val}</td>"
    table_html += "</tr>"

    table_html += "</table>"
    html += table_html
    
    sig1_img = f"<img src='{sig1_uri}' width='110'>" if sig1_uri else "<div style='color: #888; font-size: 10px; padding: 20px;'>[Tanda Tangan 1]</div>"
    sig2_img = f"<img src='{sig2_uri}' width='110'>" if sig2_uri else "<div style='color: #888; font-size: 10px; padding: 20px;'>[Tanda Tangan 2]</div>"
    sig3_img = f"<img src='{sig3_uri}' width='110'>" if sig3_uri else "<div style='color: #888; font-size: 10px; padding: 20px;'>[Tanda Tangan 3]</div>"
    
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
    <div class="trademark">Source & maintained by Herdani Rahman Account</div>
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
            components.html(master_html_start + all_qc_html + master_html_end, height=850, scrolling=True)
