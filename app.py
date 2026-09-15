import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="Radiology QC Dashboard", layout="wide")
st.title("Daily Quality Control & Environment Monitoring")

# Connect to Google Sheets (Assumes you have two tabs: 'SuhuData' and 'QCData')
conn = st.connection("gsheets", type=GSheetsConnection)
# df_env = conn.read(worksheet="SuhuData") 
# df_qc = conn.read(worksheet="QCData")

# --- MOCK DATA FOR TESTING (Replace with the conn.read above in production) ---
days = list(range(1, 32))
shifts = ['P', 'S', 'M']
full_timeline = [f"{day}-{shift}" for day in days for shift in shifts]
df_env = pd.DataFrame({'X_Axis': full_timeline})
df_env['Suhu'] = np.random.uniform(17, 24, size=len(full_timeline)) # Mock temp
df_env['Kelembapan'] = np.random.uniform(35, 65, size=len(full_timeline)) # Mock humidity
# -----------------------------------------------------------------------------

# Create Tabs for the two different outputs
tab1, tab2 = st.tabs(["Grafik Suhu & Kelembapan", "Daily QC Checklist"])

with tab1:
    st.header("Digital Monitoring Suhu dan Kelembapan")
    
    def create_segmented_graph(data_y, title, y_min, y_max, target_min, target_max):
        fig, ax = plt.subplots(figsize=(18, 5))
        x_coords = np.arange(len(df_env))
        y_vals = data_y.values
        
        valid_indices = np.where(~pd.isna(y_vals))[0]
        valid_x = x_coords[valid_indices]
        valid_y = y_vals[valid_indices]

        # Draw connecting line segments with conditional coloring
        for j in range(len(valid_x) - 1):
            x0, x1 = valid_x[j], valid_x[j+1]
            y0, y1 = valid_y[j], valid_y[j+1]
            
            # Red if outside target zone
            line_color = 'red' if (y0 < target_min or y0 > target_max) or (y1 < target_min or y1 > target_max) else 'black'
            ax.plot([x0, x1], [y0, y1], color=line_color, linestyle='-', linewidth=2, zorder=2)

        # Alternating background banding for days
        for i in range(len(days)):
            if i % 2 == 1:
                start_idx = i * 3
                ax.axvspan(start_idx - 0.5, start_idx + 2.5, facecolor='lightcyan', alpha=0.5, zorder=0)

        ax.set_ylim(y_min, y_max)
        ax.set_yticks(range(y_min, y_max + 1, 1 if y_max - y_min <= 15 else 5))
        ax.set_xticks(x_coords)
        ax.set_xticklabels([shift for day in days for shift in shifts], fontsize=8)
        
        # Add Day numbers
        for i, day in enumerate(days):
            ax.text(i * 3 + 1, y_max + (y_max-y_min)*0.02, str(day), ha='center', va='bottom', fontsize=10, fontweight='bold')

        ax.grid(True, which='both', linestyle='-', linewidth=0.5, color='black')
        ax.margins(x=0)
        ax.set_title(title, loc='left', fontweight='bold')
        return fig

    # Render Suhu Graph (Target 18-23C)
    fig_suhu = create_segmented_graph(df_env['Suhu'], "SUHU - Target Temperatur (18-23)°C", 15, 31, 18, 23)
    st.pyplot(fig_suhu)

    # Render Kelembapan Graph (Target 40-60%)
    fig_kelembapan = create_segmented_graph(df_env['Kelembapan'], "KELEMBAPAN - Target kelembapan (40-60%)", 30, 65, 40, 60)
    st.pyplot(fig_kelembapan)

with tab2:
    st.header("Form Digital Daily Quality Control")
    st.subheader("Cek Kelengkapan & Peralatan Penunjang")
    
    # In production, this data pulls from Google Forms/Sheets where radiographers check boxes
    # This creates a stylized table mapping day 1-31 with checkmarks
    qc_categories = [
        "Control box / panel listrik (Lampu indikator menyala)",
        "AC (Menyala / Berfungsi)",
        "Suhu Ruangan (18-22 C)",
        "Kelembapan Udara (40-60%)",
        "Kebersihan ruangan (Tidak terlihat sampah)"
    ]
    
    # Mock dataframe with checkmarks for demonstration
    qc_data = { "PARAMETER": qc_categories }
    for i in range(1, 32):
        qc_data[str(i)] = ["✓" if np.random.rand() > 0.1 else "X" for _ in range(len(qc_categories))]
        
    df_display_qc = pd.DataFrame(qc_data)
    
    # Render interactive table in the app
    st.dataframe(df_display_qc, use_container_width=True, hide_index=True)
