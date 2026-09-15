import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os
from datetime import datetime

st.set_page_config(page_title="Mandaya Radiology QC & Monitoring", layout="wide")

st.title("🏥 Mandaya Royal Hospital Puri - Radiology QC & Monitoring")
st.markdown("### Digital Daily Quality Control & Temperature Monitoring System")

# Automatically find all monthly Excel files in the repository
excel_files = [f for f in os.listdir() if f.endswith('.xlsx') and ("DATABASE" in f or any(m in f for m in ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember']))]

if not excel_files:
    st.error("⚠️ No monthly Excel database files found in your GitHub repository! Please upload your monthly Excel files.")
else:
    # Sidebar Navigation
    st.sidebar.header("Navigasi Aplikasi")
    selected_file = st.sidebar.selectbox("Pilih Bulan / File Database:", sorted(excel_files))
    
    app_mode = st.sidebar.radio("Pilih Menu:", ["📊 View Dashboard & Graphs", "📝 Form Input QC (Scan QR)"])

    # Load selected Excel file
    xls = pd.ExcelFile(selected_file)
    sheet_names = xls.sheet_names
    env_sheets = [s for s in sheet_names if "QC" not in s and "PILOT" not in s]

    if app_mode == "📊 View Dashboard & Graphs":
        st.subheader(f"📊 Visualisasi Grafik - {selected_file}")
        
        selected_sheet = st.selectbox("Pilih Modality / Ruangan:", env_sheets)
        
        if selected_sheet:
            df = pd.read_excel(selected_file, sheet_name=selected_sheet)
            
            suhu_col = [c for c in df.columns if 'Suhu' in c][0] if any('Suhu' in c for c in df.columns) else None
            kelembapan_col = [c for c in df.columns if 'Kelembapan' in c][0] if any('Kelembapan' in c for c in df.columns) else None
            
            if suhu_col and kelembapan_col:
                days = list(range(1, 32))
                shifts = ['P', 'S', 'M']
                full_timeline = [f"{day}-{shift}" for day in days for shift in shifts]
                
                plot_df = pd.DataFrame({'X_Axis': full_timeline})
                
                if 'Tanggal Pengisian' in df.columns and 'Jadwal Dinas' in df.columns:
                    df['Tanggal_Str'] = df['Tanggal Pengisian'].astype(str).str.split('/').str[0].str.strip()
                    df['X_Axis'] = df['Tanggal_Str'] + "-" + df['Jadwal Dinas'].astype(str).str.strip()
                    plot_df = pd.merge(plot_df, df[['X_Axis', suhu_col, kelembapan_col]], on='X_Axis', how='left')

                def render_chart(data_series, title, y_min, y_max, target_min, target_max, unit):
                    fig, ax = plt.subplots(figsize=(18, 4.5))
                    x_coords = np.arange(len(plot_df))
                    y_vals = data_series.values if data_series is not None else np.array([np.nan]*len(full_timeline))
                    
                    valid_indices = np.where(~pd.isna(y_vals))[0]
                    valid_x = x_coords[valid_indices]
                    valid_y = y_vals[valid_indices]

                    for j in range(len(valid_x) - 1):
                        x0, x1 = valid_x[j], valid_x[j+1]
                        y0, y1 = valid_y[j], valid_y[j+1]
                        line_color = 'red' if (y0 < target_min or y0 > target_max) or (y1 < target_min or y1 > target_max) else 'black'
                        ax.plot([x0, x1], [y0, y1], color=line_color, linestyle='-', linewidth=2, zorder=2)

                    for i in range(len(days)):
                        if i % 2 == 1:
                            start_idx = i * 3
                            ax.axvspan(start_idx - 0.5, start_idx + 2.5, facecolor='lightcyan', alpha=0.6, zorder=0)

                    ax.set_ylim(y_min, y_max)
                    ax.set_yticks(range(y_min, y_max + 1, 1 if y_max - y_min <= 15 else 5))
                    ax.set_xticks(x_coords)
                    ax.set_xticklabels([s for d in days for s in shifts], fontsize=7)
                    
                    for i, day in enumerate(days):
                        ax.text(i * 3 + 1, y_max + (y_max-y_min)*0.03, str(day), ha='center', va='bottom', fontsize=9, fontweight='bold')

                    ax.grid(True, which='both', linestyle='-', linewidth=0.5, color='gray')
                    ax.margins(x=0)
                    ax.set_title(f"{title} (Target: {target_min}-{target_max}{unit})", loc='left', fontweight='bold', fontsize=12)
                    return fig

                st.pyplot(render_chart(plot_df[suhu_col], f"SUHU - {selected_sheet}", 15, 30, 18, 23, "°C"))
                st.pyplot(render_chart(plot_df[kelembapan_col], f"KELEMBAPAN - {selected_sheet}", 30, 65, 40, 60, "%"))
            else:
                st.warning("Kolom suhu atau kelembapan tidak ditemukan pada sheet ini.")

    elif app_mode == "📝 Form Input QC (Scan QR)":
        st.subheader("📝 Formulir Input Harian Petugas")
        st.markdown("Silakan lengkapi data di bawah ini setelah melakukan pengecekan ruangan.")
        
        with st.form("qc_input_form"):
            col1, col2, col3 = st.columns(3)
            with col1:
                petugas = st.text_input("Nama Petugas (Inisial):", "RR")
                modality_input = st.selectbox("Pilih Modality / Ruangan:", env_sheets)
            with col2:
                tanggal = st.number_input("Tanggal:", min_value=1, max_value=31, value=datetime.now().day)
                bulan_input = st.selectbox("Bulan Target:", ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"])
            with col3:
                jadwal_dinas = st.selectbox("Shift Dinas:", ["P", "S", "M"])
                tahun = st.number_input("Tahun:", value=2026)

            st.markdown("---")
            st.markdown("#### Parameter Lingkungan")
            col_t, col_h = st.columns(2)
            with col_t:
                suhu_val = st.number_input("Suhu Ruangan (°C):", min_value=10.0, max_value=40.0, value=20.0, step=0.5)
            with col_h:
                kelembapan_val = st.number_input("Kelembapan Ruangan (%):", min_value=10.0, max_value=90.0, value=50.0, step=1.0)
            
            keterangan = st.text_input("Keterangan Tambahan:", "")
            
            submit_btn = st.form_submit_button(label="💾 Simpan Data ke Database")
            
            if submit_btn:
                try:
                    new_row = {
                        'Stamp Amser': datetime.now(),
                        'Nama Petugas': petugas,
                        'Tanggal Pengisian': tanggal,
                        'Bulan': bulan_input,
                        'Tahun': tahun,
                        'Jadwal Dinas': jadwal_dinas,
                    }
                    
                    with pd.ExcelWriter(selected_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                        df_existing = pd.read_excel(selected_file, sheet_name=modality_input)
                        suhu_c = [c for c in df_existing.columns if 'Suhu' in c][0]
                        kelembapan_c = [c for c in df_existing.columns if 'Kelembapan' in c][0]
                        
                        new_row[suhu_c] = suhu_val
                        new_row[kelembapan_c] = kelembapan_val
                        new_row['Keterangan'] = keterangan
                        
                        df_updated = pd.concat([df_existing, pd.DataFrame([new_row])], ignore_index=True)
                        df_updated.to_excel(writer, sheet_name=modality_input, index=False)
                        
                    st.success("✅ Data berhasil disimpan! Silakan cek menu grafik untuk melihat pembaruannya.")
                except Exception as e:
                    st.error(f"Gagal menyimpan data: {e}")
