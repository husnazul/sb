import os
import pandas as pd
import streamlit as st
from datetime import datetime

# Folder setup
UPLOAD_FOLDER = 'uploads'
PROCESSED_FILE = 'processed_data.xlsx'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Set Streamlit theme
st.set_page_config(page_title="DASA Processor", layout="wide")
st.markdown("""
    <style>
        .main { background-color: #f0f2f6; }
        .stButton > button {
            background-color: #4CAF50;
            color: white;
        }
        .stFileUploader label {
            color: #333;
        }
    </style>
""", unsafe_allow_html=True)

# Name mapping dictionary
name_mapping = {
    'genting': 'Genting Sanyen Power Sdn. Bhd.',
    'jimaheast': 'Jimah East Power Sdn Bhd',
    'jimahenergy': 'Jimah Energy Venture Sdn Bhd',
    'kev': 'Kapar Energy Ventures Sdn. Bhd.',
    'pengerang': 'Pengerang Power Sdn Bhd',
    'prai': 'Prai Power Sdn Bhd',
    'cameron': 'S.J. Cameron Highlands',
    'gelugor': 'S.J. Gelugor',
    'huluterengganu': 'S.J. Hulu Terengganu',
    'kenyir': 'S.J. Kenyir',
    'pergau': 'S.J. Pergau',
    'sg perak': 'S.J. Sg Perak',
    'pd1': 'S.J. Tuanku Jaafar (PD1)',
    'pd2': 'S.J. Tuanku Jaafar (PD2)',
    'ulu jela': 'S.J. Ulu Jelai',
    'segari': 'Segari Energy Ventures Sdn. Bhd.',
    'tbe': 'Tanjung Bin Energy Sdn Bhd',
    'tbp': 'Tanjung Bin Power Sdn Bhd',
    'connaught': 'TNB Connaught Bridge Sdn Bhd',
    'janamanjung': 'TNB Janamanjung Sdn Bhd',
    'manjung4': 'TNB Janamanjung Sdn Bhd (Manjung4)',
    'manjung5': 'TNB Manjung Five Sdn Bhd',
    'spg': 'Southern Power Generation',
    'putrajaya': 'S.J. Putrajaya',
    'tuah': 'Edra Energy Sdn Bhd'
}

def process_excel(file, label):
    try:
        df = pd.read_excel(file, usecols=['Unit', 'Reason', 'Start Time', 'End Time'])
        df['Start Time'] = pd.to_datetime(df['Start Time'], errors='coerce', dayfirst=True)
        df['End Time'] = pd.to_datetime(df['End Time'], errors='coerce', dayfirst=True)
        df.dropna(subset=['Start Time', 'End Time'], inplace=True)
        df['Total Hours'] = (df['End Time'] - df['Start Time']).dt.total_seconds() / 3600

        # Aggregating PO (ACC04)
        po_df = df[df['Reason'] == 'ACC04'].groupby('Unit').agg({'Total Hours': 'sum'}).reset_index()
        po_df['Total Days'] = po_df['Total Hours'] / 24
        po_df['Reason'] = 'PO'

        # Aggregating FO (ACC03 + ACC05)
        fo_df = df[df['Reason'].isin(['ACC03', 'ACC05'])].groupby('Unit').agg({'Total Hours': 'sum'}).reset_index()
        fo_df['Total Days'] = fo_df['Total Hours'] / 24
        fo_df['Reason'] = 'FO'

        result_df = pd.concat([po_df, fo_df])
        cleaned_label = label.strip().lower()
        result_df['Source'] = name_mapping.get(cleaned_label, label)
        return result_df[['Source', 'Unit', 'Reason', 'Total Days']]

    except Exception as e:
        st.error(f"Error processing {label}: {e}")
        return pd.DataFrame()

# Streamlit App
st.title("📊 DASA Reason Code Processor")
st.markdown("""
### Welcome!
Upload your **DASA file(s)** for the month.

✅ Ensure your data starts on **12:00 AM, 1st of the month** and ends on **12:00 AM, 1st of the next month**.
""")

uploaded_files = st.file_uploader("📂 Upload Excel files", accept_multiple_files=True, type="xlsx")

if uploaded_files:
    all_dfs = []
    for uploaded_file in uploaded_files:
        label = os.path.splitext(uploaded_file.name)[0]
        df = process_excel(uploaded_file, label)
        if not df.empty:
            all_dfs.append((label, df))

    if all_dfs:
        with pd.ExcelWriter(PROCESSED_FILE) as writer:
            for label, df in all_dfs:
                df.to_excel(writer, sheet_name=label[:31], index=False)

        combined_df = pd.concat([df for _, df in all_dfs])
        st.success("✅ Processing complete! Preview below:")
        st.dataframe(combined_df, use_container_width=True)

        with open(PROCESSED_FILE, "rb") as f:
            st.download_button(
                label="⬇️ Download Processed Excel",
                data=f,
                file_name=PROCESSED_FILE,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("⚠️ No valid data found in the uploaded files.")
