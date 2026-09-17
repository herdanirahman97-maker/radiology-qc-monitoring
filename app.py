import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import base64
import os
import json
import math
from datetime import datetime

st.set_page_config(page_title="Mandaya Radiology QC & Monitoring", layout="wide")

# --- DATABASE LOKAL (JSON STORAGE) ---
DB_QC_FILE = "db_qc_local.json"
DB_SUHU_FILE = "db_suhu_local.json"

def load_local_db(filepath):
    if os.path.exists(filepath):
        with open(filepath, "r") as f:
            try:
                return json.load(f)
            except:
                return {}
    return {}

def save_local_db(filepath, data):
    with open(filepath, "w") as f:
        json.dump(data, f, indent=4)

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
active_file = all_files[0] if all_files else "8. Agustus.xlsx"
xls_global = pd.ExcelFile(active_file) if os.path.exists(active_file) else None

OFFICER_INITIALS = ["JK", "RN", "ND", "BA", "DT", "WN", "NA", "SS", "PR", "AR", "AG", "SN", "PP", "RK", "LD", "RR", "FH", "HR", "VR", "AL", "WF", "EK"]
ACTIVE_MONTHS = ["Oktober", "November", "Desember"]

# --- GLOBAL CSS: PREMIUM WEB DESIGNER THEME (MODERN HOSPITAL UI) ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
    
    .stApp {
        background-color: #F8FAFC !important;
        font-family: 'Plus Jakarta Sans', sans-serif;
        color: #0F172A !important;
    }
    
    p, span, label, div, h1, h2, h3, h4, h5, h6, .stMarkdown {
        color: #0F172A !important;
    }
    
    /* Header Card Styling */
    .header-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        padding: 24px 32px;
        box-shadow: 0 4px 20px -2px rgba(15, 23, 42, 0.05);
        margin-bottom: 28px;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }

    /* Selectbox Styling */
    .stSelectbox div[data-baseweb="select"] {
        background-color: #FFFFFF !important;
        border: 1px solid #CBD5E1 !important;
        border-radius: 12px !important;
        box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
    }
    
    .stSelectbox div[data-baseweb="select"] span {
        color: #0F172A !important;
        font-weight: 600 !important;
    }

    div[data-baseweb="popover"], div[data-baseweb="menu"], ul[data-baseweb="menu"], div[role="listbox"] {
        background-color: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 12px !important;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.1) !important;
    }
    
    div[data-baseweb="popover"] *, div[data-baseweb="menu"] *, ul[data-baseweb="menu"] *, li[role="option"] *, div[role="listbox"] * {
        color: #0F172A !important;
        background-color: #FFFFFF !important;
        font-weight: 600 !important;
    }
    
    li[role="option"]:hover {
        background-color: #F1F5F9 !important;
        color: #002B5B !important;
    }

    /* QC Question Card */
    .qc-question-card {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 18px 24px;
        margin-bottom: 16px;
        box-shadow: 0 2px 6px rgba(15, 23, 42, 0.02);
        transition: all 0.2s ease;
    }
    .qc-question-card:hover {
        border-color: #CBD5E1;
        box-shadow: 0 4px 12px rgba(15, 23, 42, 0.05);
    }

    /* Custom Modern Submit Button */
    div.stButton > button, form button[type="submit"] {
        background: linear-gradient(135deg, #002B5B 0%, #004080 100%) !important;
        color: #FFFFFF !important;
        font-family: 'Plus Jakarta Sans', sans-serif !important;
        font-weight: 700 !important;
        font-size: 15px !important;
        padding: 14px 32px !important;
        border-radius: 12px !important;
        border: none !important;
        box-shadow: 0 4px 12px rgba(0, 43, 91, 0.25) !important;
        transition: all 0.3s ease !important;
        width: 100% !important;
    }
    
    div.stButton > button:hover, form button[type="submit"]:hover {
        background: linear-gradient(135deg, #001E3F 0%, #002B5B 100%) !important;
        box-shadow: 0 6px 16px rgba(0, 43, 91, 0.35) !important;
        transform: translateY(-1px);
    }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# --- MAPPING RUANGAN KE SHEET QC ---
ROOM_TO_QC_MAP = {
    "R.Teknik MRI": None,
    "R.Teknik CT Scan": None,
    "Lemari ALKES": None,
    "Lemari CD": None,
    "MRI": "QC MRI",
    "Mammografi": "QC Mammografi",
    "USG": "QC USG",
    "CTScan": "QC CT-Scan",
    "Fluoroskopi": "QC Fluoroskopi",
    "Mobile X-Ray": "QC Mobile",
    "Konvensional DR": "QC Konvensional DR"
}

# --- HTML BUILDER SUHU & KELEMBAPAN ---
def build_room_html(sheet_name, month_name, logo_data_uri):
    data_dict = {}
    local_suhu_db = load_local_db(DB_SUHU_FILE)
    file_key = f"{month_name}"
    lookup_sheet = sheet_name if sheet_name.endswith("Oct") else f"{sheet_name} Oct"

    if file_key in local_suhu_db and lookup_sheet in local_suhu_db[file_key]:
        for entry in local_suhu_db[file_key][lookup_sheet]:
            d = int(entry['tanggal'])
            s = entry['dinas'].upper()
            data_dict[(d, s)] = {
                'suhu': float(entry['suhu']), 
                'kel': float(entry['kelembapan']), 
                'nama': entry['petugas']
            }

    img_html = f"<img src='{logo_data_uri}' width='150'>" if logo_data_uri else "<b>[LOGO MISSING]</b>"
    ruang_name = sheet_name.replace(" Oct", "")
    header_bg, header_fg, nama_bg = "#002B5B", "#FFFFFF", "#C9DAF8"
    
    html = f"""
    <div class="page-container" style="background: white; padding: 25px; margin-bottom: 35px; border: 1px solid #CBD5E1; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.05);">
        <div style='display: flex; align-items: center; justify-content: space-between; margin-bottom: 20px; color: black; border-bottom: 2px solid #002B5B; padding-bottom: 15px;'>
            <div style='flex: 0 0 auto;'>{img_html}</div>
            <div style='flex: 1 1 auto; text-align: center;'>
                <h1 style='margin: 0; font-weight: 700; font-size: 26px; color: #002B5B;'>Form Digital Monitoring Suhu dan Kelembapan</h1>
                <h3 style='margin: 6px 0 0 0; font-weight: 600; font-size: 16px; color: #333;'>Departemen Radiologi Tahun 2026</h3>
            </div>
            <div style='flex: 0 0 150px;'></div>
        </div>
        
        <div style='font-weight: 700; font-size: 14px; margin-bottom: 10px; color: #002B5B;'>
            RUANG : {ruang_name}<br>
            BULAN : {month_name.upper()} 2026
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
                if val is not None and not math.isnan(val) and round(val) == temp:
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
                if val is not None and not math.isnan(val) and round(val / 5) * 5 == hum:
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

# --- HTML BUILDER QC ---
def build_qc_html(sheet_name, month_name, df, logo_data_uri, sig1_uri, sig2_uri, sig3_uri):
    img_html = f"<img src='{logo_data_uri}' width='150'>" if logo_data_uri else "<b>[LOGO MISSING]</b>"
    modality_title = sheet_name.replace("QC ", "")
    
    param_col_idx = 5
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
            
    local_qc_db = load_local_db(DB_QC_FILE)
    file_key = f"{month_name}"
    sheet_overrides = local_qc_db.get(file_key, {}).get(sheet_name, {})

    html = f"""
    <div class="page-container" style="background: white; padding: 25px; margin-bottom: 35px; border: 1px solid #CBD5E1; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.05);">
        <div style='display: flex; align-items: center; justify-content: space-between; margin-bottom: 15px; color: black; border-bottom: 2px solid #002B5B; padding-bottom: 10px;'>
            <div style='flex: 0 0 auto;'>{img_html}</div>
            <div style='flex: 1 1 auto; text-align: center;'>
                <h1 style='margin: 0; font-weight: 700; font-size: 24px; color: #002B5B;'>Form Digital Daily Quality Control</h1>
                <h3 style='margin: 4px 0 0 0; font-weight: 600; font-size: 15px; color: #333;'>Departemen Radiologi Tahun 2026</h3>
            </div>
            <div style='flex: 0 0 150px;'></div>
        </div>
        
        <div style='font-weight: 700; font-size: 13px; margin-bottom: 10px; color: #002B5B;'>
            MODALITAS : {modality_title}<br>
            BULAN : {month_name.upper()} 2026
        </div>
    """
    
    table_html = "<table class='qc-table'>"
    row_counter = 0
    
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
                
            kegiatan = row_vals[1].strip() if len(row_vals) > 1 else ""
            parameter = row_vals[param_col_idx].strip() if param_col_idx < len(row_vals) else ""
            
            if not parameter:
                for c in range(2, param_col_idx + 1):
                    if c < len(row_vals) and row_vals[c].strip() != "" and row_vals[c].strip() not in ["✓", "X"]:
                        parameter = row_vals[c].strip()
                        break
            
            table_html += "<tr style='height: 30px;'>"
            table_html += f"<td style='border: 1px solid black; padding: 5px; text-align: center;' width='45px'>{row_vals[0]}</td>"
            table_html += f"<td colspan='2' style='border: 1px solid black; padding: 5px; text-align: left;' width='200px'>{kegiatan}</td>"
            table_html += f"<td colspan='2' style='border: 1px solid black; padding: 5px; text-align: left;' width='250px'>{parameter}</td>"
            
            for day_idx in range(1, 32):
                val = ""
                if str(day_idx) in sheet_overrides and str(row_counter) in sheet_overrides[str(day_idx)]:
                    val = sheet_overrides[str(day_idx)][str(row_counter)].get('status', '')
                table_html += f"<td style='border: 1px solid black; padding: 4px; text-align: center;'>{val}</td>"
            table_html += "</tr>"
            row_counter += 1
            
    table_html += "<tr style='background-color: #C9DAF8; font-weight: bold; height: 32px;'>"
    table_html += "<td colspan='5' style='border: 1px solid black; padding: 5px; text-align: center;'>NAMA PEKERJA RADIASI</td>"
    for day_idx in range(1, 32):
        val = ""
        if str(day_idx) in sheet_overrides and 'worker' in sheet_overrides[str(day_idx)]:
            val = sheet_overrides[str(day_idx)]['worker']
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
                X : Tidak Memenuhi / Tidak Lulus Uji / Lengkap
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

logo_b64 = get_image_base64(LOGO_PATH)
sig1_b64 = get_image_base64(SIG1_PATH)
sig2_b64 = get_image_base64(SIG2_PATH)
sig3_b64 = get_image_base64(SIG3_PATH)

# --- HEADER: DESIGN PROFESIONAL KHAZANAH KORPORAT RUMAH SAKIT ---
logo_img_tag = f"<img src='{logo_b64}' width='150'>" if logo_b64 else "<b style='color:#002B5B;'>MANDAYA HOSPITAL</b>"

st.markdown(f"""
<div class="header-card">
    <div style="display: flex; align-items: center; gap: 20px;">
        {logo_img_tag}
    </div>
    <div style="text-align: center; flex-grow: 1;">
        <h2 style="margin: 0; color: #002B5B; font-size: 24px; font-weight: 800; letter-spacing: -0.5px;">MANDAYA ROYAL HOSPITAL — RADIOLOGY QC SYSTEM</h2>
        <p style="margin: 6px 0 0 0; color: #64748B; font-size: 14px; font-weight: 500;">Departemen Radiologi Tahun 2026 | Enterprise Monitoring Portal</p>
    </div>
    <div style="text-align: right;">
        <span style="background: #E2E8F0; color: #002B5B; padding: 6px 14px; border-radius: 20px; font-size: 12px; font-weight: 700;">v2.5 Live</span>
    </div>
</div>
""", unsafe_allow_html=True)

# --- MENU UTAMA DI TENGAH ---
col_space1, col_nav, col_space2 = st.columns([1, 6, 1])
with col_nav:
    main_menu = st.radio(
        "Pilih Menu Utama:", 
        [
            "📝 Input Data", 
            "📅 Live Preview Bulanan", 
            "📂 Review Backdate", 
            "⚙️ Revisi / Hapus"
        ], 
        horizontal=True,
        label_visibility="collapsed"
    )

st.markdown("<br>", unsafe_allow_html=True)

master_html_start = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;700&display=swap');
        body {
            background-color: white;
            color: black;
            font-family: 'Plus Jakarta Sans', sans-serif;
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
            background: linear-gradient(135deg, #002B5B 0%, #004080 100%);
            color: white;
            border: none;
            padding: 12px 24px;
            font-family: 'Plus Jakarta Sans', sans-serif;
            font-weight: 700;
            font-size: 14px;
            border-radius: 8px;
            cursor: pointer;
            margin: 15px 0 20px 20px;
            box-shadow: 0 4px 12px rgba(0,43,91,0.2);
        }
        .print-btn:hover { background: linear-gradient(135deg, #001E3F 0%, #002B5B 100%); }
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
        <button class="print-btn" onclick="window.print()">🖨️ Download / Save as PDF (Print All)</button>
    </div>
"""
master_html_end = "</body></html>"

raw_suhu_sheets = [s for s in xls_global.sheet_names if str(s).endswith("Oct")] if xls_global else []
clean_room_list = [s.replace(" Oct", "") for s in raw_suhu_sheets]
qc_sheets_all = [s for s in xls_global.sheet_names if "QC" in s] if xls_global else []

# ==========================================
# 1. MENU: INPUT DATA
# ==========================================
if main_menu == "📝 Input Data":
    st.markdown("### 📝 Form Pengisian Data Harian (Bulanan & Terpadu)")
    st.markdown("<p style='color: #64748B !important;'>Pilih bulan target dan modalitas Anda. Checklist QC otomatis hanya muncul pada **Dinas Pagi (P)**.</p>", unsafe_allow_html=True)
    
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        target_month_input = st.selectbox("📅 Pilih Bulan Target Arsip:", ACTIVE_MONTHS, index=0)
    with col_m2:
        selected_room_clean = st.selectbox(
            "🏥 Pilih Modalitas / Ruangan Utama:", 
            clean_room_list
        )
    
    selected_room = f"{selected_room_clean} Oct"
    auto_matched_qc = ROOM_TO_QC_MAP.get(selected_room_clean, None)
    has_qc_room = auto_matched_qc is not None

    with st.form("clean_corporate_input_form"):
        st.markdown("<br><h4>📌 Step 2: Identitas & Waktu Pengisian</h4>", unsafe_allow_html=True)
        col_a, col_b, col_c = st.columns(3)
        with col_a:
            petugas_input = st.selectbox("Inisial Petugas", OFFICER_INITIALS)
        with col_b:
            tanggal_input = st.selectbox("Tanggal Pengisian (1-31)", list(range(1, 32)))
        with col_c:
            dinas_input = st.selectbox("Jadwal Dinas", ["P", "S", "M"])

        st.markdown(f"<br><h4>Step 3: 🌡️ Suhu & Kelembapan Ruangan ({selected_room_clean})</h4>", unsafe_allow_html=True)
        col_s, col_k = st.columns(2)
        with col_s:
            suhu_input = st.number_input("Suhu Ruangan (°C) [Target 18 - 23°C]", min_value=15.0, max_value=30.0, value=22.0, step=0.5)
        with col_k:
            kel_input = st.number_input("Kelembapan Ruangan (%) [Target 40 - 60%]", min_value=30.0, max_value=70.0, value=55.0, step=1.0)

        qc_responses = {}
        is_morning_shift = (dinas_input == "P")
        show_qc_section = has_qc_room and is_morning_shift

        if show_qc_section:
            st.markdown(f"<br><h4>Step 4: 📋 Daily Quality Control Spesifik — {auto_matched_qc.replace('QC ', '')} (Dinas Pagi)</h4>", unsafe_allow_html=True)
            include_qc = st.checkbox("Lakukan pengisian checklist QC harian?", value=True)
            
            qc_items_list = []
            if include_qc and auto_matched_qc and xls_global:
                df_qc_sheet = pd.read_excel(xls_global, sheet_name=auto_matched_qc, header=None)
                param_idx = 5
                for r_i in [8, 7, 6]:
                    if r_i < len(df_qc_sheet):
                        row_h = [str(x).upper() for x in df_qc_sheet.iloc[r_i].values]
                        for c_i, val in enumerate(row_h):
                            if "PARAMETER" in val:
                                param_idx = c_i
                                break
                        if param_idx != 5:
                            break

                curr_cat = ""
                for idx, row in df_qc_sheet.iterrows():
                    if idx >= 8:
                        r_vals = [str(x) if pd.notna(x) else "" for x in row.values]
                        r_text = " ".join(r_vals).upper()
                        if any(kwd in r_text for kwd in ["DISIAPKAN", "MENGETAHUI", "RHEINNER", "JOKO", "CHRISTOPHER"]):
                            continue
                        non_empty = [v.strip() for i, v in enumerate(r_vals[:param_idx]) if v.strip() != "" and v.strip() != "NO"]
                        if len(non_empty) == 1 and r_vals[0] != "" and not r_vals[0].isdigit():
                            curr_cat = non_empty[0]
                            continue
                        if len(r_vals) > 1 and r_vals[0].isdigit():
                            keg = r_vals[1].strip()
                            param = r_vals[param_idx].strip() if param_idx < len(r_vals) else ""
                            if not param:
                                for c in range(2, param_idx + 1):
                                    if c < len(r_vals) and r_vals[c].strip() != "" and r_vals[c].strip() not in ["✓", "X"]:
                                        param = r_vals[c].strip()
                                        break
                            qc_items_list.append((curr_cat, r_vals[0], keg, param))

            if include_qc and auto_matched_qc:
                active_category = ""
                row_idx_tracker = 0
                for cat, no, keg, param in qc_items_list:
                    if cat != active_category:
                        active_category = cat
                        st.markdown(f"<br><b>📌 {active_category}</b>", unsafe_allow_html=True)
                    
                    label = f"**{no}. {keg}** — *Parameter: {param}*" if param else f"**{no}. {keg}**"
                    
                    st.markdown(f"""
                    <div class="qc-question-card">
                        <div style="font-weight: 600; font-size: 14px; margin-bottom: 8px; color: #002B5B;">{label}</div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    qc_responses[row_idx_tracker] = st.radio(
                        f"Pilih status pertanyaan {no}", 
                        ["Berfungsi / Lengkap / Baik (✓)", "Tidak Berfungsi / Rusak (X)"], 
                        key=f"corp_qc_{target_month_input}_{auto_matched_qc}_{row_idx_tracker}",
                        label_visibility="collapsed"
                    )
                    row_idx_tracker += 1
        elif has_qc_room and not is_morning_shift:
            st.markdown(f"<br><div style='padding: 16px; background-color: #FEF3C7; border: 1px solid #F59E0B; border-radius: 12px; font-weight: 600; color: #92400E;'>ℹ️ Informasi: Anda memilih **Dinas {dinas_input}**. Checklist QC hanya diisi pada **Dinas Pagi (P)** (dilakukan sekali sehari). Form ini hanya akan menyimpan Suhu & Kelembapan.</div>", unsafe_allow_html=True)
        else:
            st.markdown(f"<br><div style='padding: 16px; background-color: #F1F5F9; border: 1px solid #CBD5E1; border-radius: 12px; font-weight: 600; color: #002B5B;'>ℹ️ Ruangan **{selected_room_clean}** hanya memerlukan pemantauan Suhu & Kelembapan (Tidak memiliki checklist QC).</div>", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        submitted_data = st.form_submit_button("🚀 Submit Data Harian")
        
        if submitted_data:
            local_suhu = load_local_db(DB_SUHU_FILE)
            file_key = target_month_input
            if file_key not in local_suhu:
                local_suhu[file_key] = {}
            if selected_room not in local_suhu[file_key]:
                local_suhu[file_key][selected_room] = []
                
            local_suhu[file_key][selected_room].append({
                "tanggal": tanggal_input,
                "dinas": dinas_input,
                "petugas": petugas_input,
                "suhu": suhu_input,
                "kelembapan": kel_input
            })
            save_local_db(DB_SUHU_FILE, local_suhu)
            
            if show_qc_section and include_qc and auto_matched_qc:
                local_qc = load_local_db(DB_QC_FILE)
                if file_key not in local_qc:
                    local_qc[file_key] = {}
                if auto_matched_qc not in local_qc[file_key]:
                    local_qc[file_key][auto_matched_qc] = {}
                    
                day_key = str(tanggal_input)
                if day_key not in local_qc[file_key][auto_matched_qc]:
                    local_qc[file_key][auto_matched_qc][day_key] = {}
                    
                local_qc[file_key][auto_matched_qc][day_key]['worker'] = petugas_input
                for r_idx, resp in qc_responses.items():
                    symbol = "✓" if "Baik" in resp else "X"
                    local_qc[file_key][auto_matched_qc][day_key][str(r_idx)] = {"status": symbol}
                    
                save_local_db(DB_QC_FILE, local_qc)
                
            st.success(f"✅ Data untuk **{selected_room_clean}** bulan **{target_month_input}** berhasil disimpan!")
            st.balloons()

# ==========================================
# 2. MENU: LIVE PREVIEW BULANAN (DENGAN DOWNLOAD ALL PADA SUHU & QC)
# ==========================================
elif main_menu == "📅 Live Preview Bulanan":
    st.markdown("### 📅 Live Preview Rekapitulasi Per Bulan")
    st.markdown("<p style='color: #64748B !important;'>Melihat rekapitulasi data real-time untuk bulan Oktober, November, dan Desember.</p>", unsafe_allow_html=True)
    
    col_live1, col_live2, col_live3 = st.columns(3)
    with col_live1:
        view_month = st.selectbox("Pilih Bulan Live:", ACTIVE_MONTHS, index=0)
    with col_live2:
        live_category = st.selectbox("Pilih Kategori:", ["🌡️ Suhu & Kelembapan", "📋 Daily Quality Control (QC)"])
    with col_live3:
        if live_category == "🌡️ Suhu & Kelembapan":
            room_selection_options = ["🌐 Pilih Semua Ruangan (Download All)"] + clean_room_list
            selected_room_preview = st.selectbox("Pilih Ruangan:", room_selection_options)
        else:
            qc_selection_options = ["🌐 Pilih Semua Modality (Download All)"] + qc_sheets_all
            selected_qc = st.selectbox("Pilih Modality QC:", qc_selection_options)

    st.markdown("<br>", unsafe_allow_html=True)
    if live_category == "🌡️ Suhu & Kelembapan":
        if selected_room_preview == "🌐 Pilih Semua Ruangan (Download All)":
            st.markdown("#### 📥 Mode Download All: Menampilkan Seluruh Ruangan (Suhu & Kelembapan) Sekaligus")
            combined_html = ""
            for r_clean in clean_room_list:
                full_sheet = f"{r_clean} Oct"
                combined_html += build_room_html(full_sheet, view_month, logo_b64) + "<div style='page-break-after: always;'></div>"
            
            components.html(master_html_start + combined_html + master_html_end, height=1200, scrolling=True)
        else:
            full_sheet = f"{selected_room_preview} Oct"
            room_html = build_room_html(full_sheet, view_month, logo_b64)
            components.html(master_html_start + room_html + master_html_end, height=850, scrolling=True)
    else:
        if selected_qc == "🌐 Pilih Semua Modality (Download All)":
            st.markdown("#### 📥 Mode Download All: Menampilkan Seluruh Modality QC Sekaligus")
            combined_qc_html = ""
            for qc_s in qc_sheets_all:
                df_qc = pd.read_excel(xls_global, sheet_name=qc_s, header=None) if xls_global else pd.DataFrame()
                combined_qc_html += build_qc_html(qc_s, view_month, df_qc, logo_b64, sig1_b64, sig2_b64, sig3_b64) + "<div style='page-break-after: always;'></div>"
            
            components.html(master_html_start + combined_qc_html + master_html_end, height=1200, scrolling=True)
        else:
            df_qc = pd.read_excel(xls_global, sheet_name=selected_qc, header=None) if xls_global else pd.DataFrame()
            qc_html = build_qc_html(selected_qc, view_month, df_qc, logo_b64, sig1_b64, sig2_b64, sig3_b64)
            components.html(master_html_start + qc_html + master_html_end, height=950, scrolling=True)

# ==========================================
# 3. MENU: REVIEW BACKDATE
# ==========================================
elif main_menu == "📂 Review Backdate":
    st.markdown("### 📂 Tinjauan Rekapitulasi Arsip Excel")
    
    col_filter1, col_filter2 = st.columns(2)
    with col_filter1:
        selected_file = st.selectbox("Pilih File Excel:", all_files)
    with col_filter2:
        review_category = st.selectbox("Pilih Kategori Tinjauan:", ["🌡️ Suhu & Kelembapan", "📋 Daily Quality Control (QC)"])
        
    xls = pd.ExcelFile(selected_file) if selected_file else None
    
    if xls:
        st.markdown("<br>", unsafe_allow_html=True)
        if review_category == "🌡️ Suhu & Kelembapan":
            raw_data_sheets = [s for s in xls.sheet_names if str(s).endswith("Oct")]
            selected_sheet = st.selectbox(
                "Pilih Ruangan:", 
                raw_data_sheets, 
                format_func=lambda x: str(x).replace(" Oct", "").strip()
            )
            room_html = build_room_html(selected_sheet, "Oktober", logo_b64)
            components.html(master_html_start + room_html + master_html_end, height=850, scrolling=True)
            
        else:
            qc_sheets = [s for s in xls.sheet_names if "QC" in s]
            selected_qc = st.selectbox("Pilih Modality QC:", qc_sheets)
            df_qc = pd.read_excel(xls, sheet_name=selected_qc, header=None)
            qc_html = build_qc_html(selected_qc, "Oktober", df_qc, logo_b64, sig1_b64, sig2_b64, sig3_b64)
            components.html(master_html_start + qc_html + master_html_end, height=950, scrolling=True)

# ==========================================
# 4. MENU: REVISI / HAPUS
# ==========================================
elif main_menu == "⚙️ Revisi / Hapus":
    st.markdown("### ⚙️ Menu Revisi & Penghapusan Data")
    st.markdown("<p style='color: #64748B !important;'>Kelola dan hapus data input harian jika terjadi kesalahan tanggal atau nilai.</p>", unsafe_allow_html=True)
    
    col_rev1, col_rev2 = st.columns(2)
    with col_rev1:
        selected_month_clean = st.selectbox("Pilih Bulan Data:", ACTIVE_MONTHS, index=0)
    with col_rev2:
        db_type_to_clean = st.selectbox("Pilih Kategori Data:", ["QC", "Suhu & Kelembapan"])
    
    st.markdown("---")
    if db_type_to_clean == "QC":
        current_qc_db = load_local_db(DB_QC_FILE)
        file_key = selected_month_clean
        if file_key in current_qc_db and current_qc_db[file_key]:
            st.write(f"Daftar input QC tersimpan untuk bulan **{file_key}**:")
            for modality, days in current_qc_db[file_key].items():
                for d_key in list(days.keys()):
                    col1, col2 = st.columns([3, 1])
                    with col1:
                        st.write(f"• **{modality}** — Tanggal {d_key}")
                    with col2:
                        if st.button("🗑️ Hapus", key=f"del_qc_{modality}_{d_key}"):
                            del current_qc_db[file_key][modality][d_key]
                            save_local_db(DB_QC_FILE, current_qc_db)
                            st.success(f"Data {modality} tanggal {d_key} berhasil dihapus!")
                            st.rerun()
        else:
            st.info("Belum ada data input QC lokal untuk bulan ini.")
    else:
        current_suhu_db = load_local_db(DB_SUHU_FILE)
        file_key = selected_month_clean
        if file_key in current_suhu_db:
            st.write(f"Daftar input Suhu & Kelembapan tersimpan untuk bulan **{file_key}**:")
            for room, entries in current_suhu_db[file_key].items():
                for idx, entry in enumerate(entries):
                    col1, col2 = st.columns([3, 1])
                    with col1:
                        st.write(f"• **{room.replace(' Oct', '')}** — Tgl {entry['tanggal']} (Dinas {entry['dinas']}, Petugas: {entry['petugas']})")
                    with col2:
                        if st.button("🗑️ Hapus", key=f"del_suhu_{room}_{idx}"):
                            current_suhu_db[file_key][room].pop(idx)
                            save_local_db(DB_SUHU_FILE, current_suhu_db)
                            st.success("Data suhu berhasil dihapus!")
                            st.rerun()
        else:
            st.info("Belum ada data input Suhu lokal untuk bulan ini.")
