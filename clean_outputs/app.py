import streamlit as st
import pandas as pd
import openpyxl
import io

st.set_page_config(page_title="Raventrack ETL Pipeline Dashboard", page_icon="📊", layout="wide")

st.title("📊 Raventrack ETL Pipeline Dashboard")
st.write("Upload your files below to automatically process data and download the cleaned end product.")

# Layout Columns for file uploads
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Raventrack Export")
    raw_file = st.file_uploader("Upload Raw Raventrack CSV", type=["csv"])

with col2:
    st.subheader("2. Affiliate Log Reference")
    aff_file = st.file_uploader("Upload Data Test File (Excel)", type=["xlsx"])

if raw_file and aff_file:
    if st.button("🚀 Run ETL Pipeline", type="primary"):
        try:
            # --- 1. PROCESS THE EXCEL REFERENCE FILE ---
            # Read Nicola's 'RT Tag Reference' sheet directly
            df_aff_raw = pd.read_excel(aff_file, sheet_name='RT Tag Reference')
            
            # Clean columns by stripping whitespace and dropping empty rows
            df_aff_raw.columns = df_aff_raw.columns.str.strip()
            df_aff_clean = df_aff_raw.dropna(subset=['Affiliate Profile Username']).copy()
            
            # Create a lowercase helper join column to avoid casing mismatches
            df_aff_clean['join_username'] = df_aff_clean['Affiliate Profile Username'].astype(str).str.strip().str.lower()
            
            # --- 2. PROCESS THE CSV RAW FILE ---
            # Preview first few lines to skip potential header trash dynamically
            bytes_data = raw_file.getvalue()
            lines = bytes_data.decode('utf-8', errors='ignore').split('\n')
            
            skip_rows = 0
            for i, line in enumerate(lines[:15]):
                if 'player' in line.lower() or 'username' in line.lower():
                    skip_rows = i
                    break
            
            raw_file.seek(0)
            df_raw = pd.read_csv(raw_file, skiprows=skip_rows, low_memory=False)
            df_raw.columns = df_raw.columns.str.strip()
            
            # Identify source columns using flexible matching rules
            player_id_source = next((c for c in df_raw.columns if 'player' in c.lower()), df_raw.columns[0])
            aff_user_source = next((c for c in df_raw.columns if 'user' in c.lower() or 'affiliate' in c.lower()), df_raw.columns[1])
            
            # Safely create standard lower casing join column on the CSV data
            df_raw['join_username'] = df_raw[aff_user_source].astype(str).str.strip().str.lower()
            
            # --- 3. ETL PIPELINE CLEANING AND MERGE ---
            # Extract just the Player/User ID and the Username to clean up the output
            df_clean = df_raw[[player_id_source, 'join_username']].copy()
            df_clean.columns = ['User ID', 'join_username']
            
            # Perform the Left Merge on our lowercase tracking key directly to Nicola's data
            df_final = pd.merge(df_clean, df_aff_clean, on='join_username', how='left')
            
            # Clean up tracking helper keys and reorganize layout columns
            if 'join_username' in df_final.columns:
                df_final = df_final.drop(columns=['join_username'])
                
            # Bring 'Affiliate Profile Username' forward if it exists post-merge
            cols = list(df_final.columns)
            if 'Affiliate Profile Username' in cols:
                cols.insert(1, cols.pop(cols.index('Affiliate Profile Username')))
                df_final = df_final[cols]
                df_final = df_final.sort_values(by='Affiliate Profile Username')
            
            # --- 4. SUCCESS & EXPORT GENERATION ---
            st.success("🎉 ETL Pipeline Completed Successfully!")
            st.dataframe(df_final.head(100), use_container_width=True)
            
            # Write out to memory buffer for immediate download button activation
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Cleaned_Output')
            processed_data = output.getvalue()
            
            st.download_button(
                label="📥 Download Cleaned Excel Product",
                data=processed_data,
                file_name="Cleaned_Raventrack_Data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"❌ An error occurred during processing: {str(e)}")
else:
    st.info("💡 Please upload both required input data files above to initialize processing.")