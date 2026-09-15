import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os
from datetime import datetime

# Page Configuration
st.set_page_config(page_title="Mandaya Radiology QC & Monitoring", layout="wide")

EXCEL_FILE = "0. DATABASE QC DAN MONITOR SUHU SEPTEMBER 2026.xlsx"

# Load Data Function
@st.cache_data(ttl=5)
def load_data():
    if os.path.exists(EXCEL_FILE):
        xls = pd.ExcelFile(EXCEL_FILE)
        sheets = xls.sheet_names
        return xls, sheets
    return None, []

xls, sheet_names = load_data()

st.title("🏥 Mandaya Royal Hospital Puri - Radiology QC & Monitoring")
st.markdown("### Digital Daily Quality Control & Temperature Monitoring System")

if xls is None:
    st.error(f"Database file '{EXCEL_FILE}' not found in the repository! Please upload your Excel file to GitHub.")
else:
    # Sidebar Navigation
    st.sidebar.header("Navigation & Options")
    app_mode = st.sidebar.radio("Pilih Menu:", ["📊 View Dashboard & Graphs", "📝 Form Input QC (Scan QR)"])

    # Filter environment sheets vs QC sheets
    env_sheets = [s for s in sheet_names if "QC" not in s and "PILOT" not in s]
    
    if app_mode == "📊 View Dashboard & Graphs":
        st.subheader("Visualisasi Grafik Suhu & Kelembapan Bulanan")
        
        selected_sheet = st.selectbox("Pilih Modality / Ruangan:", env_sheets)
        
        if selected_sheet:
            df = pd.read_excel(EXCEL_FILE, sheet_name=selected_sheet)
            
            # Identify columns dynamically based on sheet structure
            suhu_col = [c for c in df.columns if 'Suhu' in c][0] if any('Suhu' in c for c in df.columns) else None
            kelembapan_col = [c for c in df.columns if 'Kelembapan' in c][0] if any('Kelembapan' in c for c in df.columns) else None
            
            if suhu_col and kelembapan_col:
                # Setup Timeline (Days 1 to 31, Shifts P, S, M)
                days = list(range(1, 32))
                shifts = ['P', 'S', 'M']
                full_timeline = [f"{day}-{shift}" for day in days for shift in shifts]
                
                plot_df = pd.DataFrame({'X_Axis': full_timeline})
                
                if 'Tanggal Pengisian' in df.columns and 'Jadwal Dinas' in df.columns:
                    df['Tanggal_Str'] = df['Tanggal Pengisian'].astype(str).str.split('/').str[0].str.strip()
                    df['X_Axis'] = df['Tanggal_Str'] + "-" + df['Jadwal Dinas'].astype(str).str.strip()
                    plot_df = pd.merge(plot_df, df[['X_Axis', suhu_col, kelembapan_col]], on='X_Axis', how='left')

                # Function to plot custom segmented lines
                def render_chart(data_series, title, y_min, y_max, target_min, target_max, unit):
                    fig, ax = plt.subplots(figsize=(18, 4.5))
                    x_coords = np.arange(len(plot_df))
                    y_vals = data_series.values if data_series is not None else np.array([np.nan]*len(full_timeline))
                    
                    valid_indices = np.where(~pd.isna(y_vals))[0]
                    valid_x = x_coords[valid_indices]
                    valid_y = y_vals[valid_indices]

                    # Segmented line coloring
                    for j in range(len(valid_x) - 1):
                        x0, x1 = valid_x[j], valid_x[j+1]
                        y0, y1 = valid_y[j], valid_y[j+1]
                        line_color = 'red' if (y0 < target_min or y0 > target_max) or (y1 < target_min or y1 > target_max) else 'black'
                        ax.plot([x0, x1], [y0, y1], color=line_color, linestyle='-', linewidth=2, zorder=2)

                    # Alternating column banding
                    for i in range(len(days)):
                        if i % 2 == 1:
                            start_idx = i * 3
                            ax.axvspan(start_idx - 0.5, start_idx + 2.5, facecolor='lightcyan', alpha=0.6, zorder=0)

                    ax.set_ylim(y_min, y_max)
                    ax.set_yticks(range(y_min, y_max + 1, 1 if y_max - y_min <= 15 else 5))
                    ax.set_xticks(x_coords)
                    ax.set_xticklabels([s for d in days for s in shifts], fontsize=7)
                    
                    # Day headers
                    for i, day in enumerate(days):
                        ax.text(i * 3 + 1, y_max + (y_max-y_min)*0.03, str(day), ha='center', va='bottom', fontsize=9, fontweight='bold')

                    ax.grid(True, which='both', linestyle='-', linewidth=0.5, color='gray')
                    ax.margins(x=0)
                    ax.set_title(f"{title} (Target: {target_min}-{target_max}{unit})", loc='left', fontweight='bold', fontsize=12)
                    return fig

                # Render Suhu Chart (Target 18-23°C)
                st.pyplot(render_chart(plot_df[suhu_col], f"SUHU - {selected_sheet}", 15, 30, 18, 23, "°C"))
                
                # Render Kelembapan Chart (Target 40-60%)
                st.pyplot(render_chart(plot_df[kelembapan_col], f"KELEMBAPAN - {selected_sheet}", 30, 65, 40, 60, "%"))
            else:
                st.warning("Kolom suhu atau kelembapan tidak ditemukan pada sheet ini.")

    elif app_mode == "📝 Form Input QC (Scan QR)":
        st.subheader("Formulir Input Harian Radiologi")
        st.markdown("Silakan isi form di bawah ini setelah melakukan pengecekan fisik di ruangan.")
        
        with st.form("qc_form"):
            col1, col2, col3 = st.columns(3)
            with col1:
                petugas = st.text_input("Nama Petugas (Inisial):", "RR")
                modality_input = st.selectbox("Pilih Modality:", env_sheets)
            with col2:
                tanggal = st.number_input("Tanggal:", min_value=1, max_value=31, value=datetime.now().day)
                bulan = st.selectbox("Bulan:", ["September", "October", "November", "December", "January", "February", "March", "April", "May", "June", "July", "August"])
            with col3:
                jadwal_dinas = st.selectbox("Jadwal Dinas (Shift):", ["P", "S", "M"])
                tahun = st.number_input("Tahun:", value=2026)

            st.markdown("---")
            st.markdown("#### Parameter Lingkungan")
            col_t, col_h = st.columns(2)
            with col_t:
                suhu_val = st.number_input("Suhu Ruangan (°C):", min_value=10.0, max_value=40.0, value=20.0, step=0.5)
            with col_h:
                kelembapan_val = st.number_input("Kelembapan Ruangan (%):", min_value=10.0, max_value=90.0, value=50.0, step=1.0)
            
            keterangan = st.text_input("Keterangan / Catatan Tambahan (Opsional):", "")
            
            submit_button = st.form_submit_button(label="💾 Simpan & Perbarui Data")
            
            if submit_button:
                # Appends data to Excel database dynamically
                new_row = {
                    'Stamp Amser': datetime.now(),
                    'Nama Petugas': petugas,
                    'Tanggal Pengisian': tanggal,
                    'Bulan': bulan,
                    'Tahun': tahun,
                    'Jadwal Dinas': jadwal_dinas,
                }
                
                try:
                    # Read existing excel, append, and save back
                    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                        # Logic to append to the specific sheet
                        df_existing = pd.read_excel(EXCEL_FILE, sheet_name=modality_input)
                        suhu_c = [c for c in df_existing.columns if 'Suhu' in c][0]
                        kelembapan_c = [c for c in df_existing.columns if 'Kelembapan' in c][0]
                        
                        new_row[suhu_c] = suhu_val
                        new_row[kelembapan_c] = kelembapan_val
                        new_row['Keterangan'] = keterangan
                        
                        df_updated = pd.concat([df_existing, pd.DataFrame([new_row])], ignore_index=True)
                        df_updated.to_excel(writer, sheet_name=modality_input, index=False)
                        
                    st.success("✅ Data berhasil disimpan! Grafik akan otomatis memperbarui dirinya.")
                except Exception as e:
                    st.error(f"Terjadi kesalahan saat menyimpan data: {e}")
