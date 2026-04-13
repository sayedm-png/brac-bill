# --- The template part of your code ---
template_name = "Template.xlsm"

try:
    # keep_vba=True is critical for .xlsm files
    wb = load_workbook(template_name, keep_vba=True)
    ws = wb.active 

    # Your updated cell mappings (double-check these in your Excel)
    ws['B5'] = str(visit_date) 
    ws['B6'] = area            
    # ... (rest of your cell updates)

    # Save to memory without losing macros
    output = BytesIO()
    wb.save(output)
    processed_data = output.getvalue()

    st.success("Official Bill Formatted Successfully!")
    
    st.download_button(
        label="📥 Download Final Bill",
        data=processed_data,
        file_name=f"Field_Bill_{area}.xlsm",
        mime="application/vnd.ms-excel.sheet.macroenabled.12" # Correct type for .xlsm
    )
except Exception as e:
    st.error(f"Error: {e}. Check if Template.xlsm is uploaded to GitHub.")
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import datetime
import re
from io import BytesIO

# --- 1. Data Fetching (Google Drive Allowances) ---
DRIVE_URL = "https://drive.google.com/file/d/1yUpT0N-8d_P1org-e5LwfffCYxco2sPb/view?ts=69dc8c08"

@st.cache_data(ttl=3600)
def fetch_allowance_data(url):
    try:
        file_id = re.search(r'[-\w]{25,}', url).group()
        direct_download = f'https://docs.google.com/spreadsheets/d/{file_id}/export?format=csv'
        return pd.read_csv(direct_download)
    except:
        return None

# --- 2. App Interface ---
st.set_page_config(page_title="Official Bill Generator", layout="centered")
st.title("BRAC Official Bill Generator")
st.write("This app populates the exact 2-page Excel template.")

allowance_df = fetch_allowance_data(DRIVE_URL)

with st.form("official_form"):
    col1, col2 = st.columns(2)
    with col1:
        visit_date = st.date_input("Date", datetime.date.today())
        if allowance_df is not None:
            area = st.selectbox("Area", options=allowance_df['Area'].unique())
            fixed_dist = allowance_df.loc[allowance_df['Area'] == area, 'Allowance'].values[0]
        else:
            area = st.text_input("Area")
            fixed_dist = st.number_input("Fixed Allowance", 0)
        
        ground_cost = st.number_input("Ground Travel", 0)

    with col2:
        b = st.number_input("Breakfast", 0)
        l = st.number_input("Lunch", 0)
        d = st.number_input("Dinner", 0)
        halt = st.number_input("Night Haltage", 0)
    
    submit = st.form_submit_button("Prepare Official Document")

if submit:
    # Calculations
    food_total = (b * 70) + (l * 140) + (d * 140)
    halt_total = halt * 150
    
    # --- 3. The Template Engine ---
    template_name = "Field Visit 8-9 March 2026 Cumilla (1).xlsm"
    
    try:
        # Load the official template
        wb = load_workbook(template_name, keep_vba=True)
        ws = wb.active # Assuming the bill is the first sheet

        # INJECTION LOGIC: Update these cell addresses to match your Excel
        # Example addresses based on typical BRAC forms:
        ws['B5'] = str(visit_date)  # Date Cell
        ws['B6'] = area             # Area Cell
        ws['E10'] = ground_cost     # Ground Travel Cell
        ws['E11'] = fixed_dist      # Fixed Distance Cell
        ws['E12'] = food_total      # Food Cell
        ws['E13'] = halt_total      # Haltage Cell

        # Save to memory
        output = BytesIO()
        wb.save(output)
        
        st.success("Format preserved. Document ready.")
        st.download_button(
            label="📥 Download Exact Format Bill",
            data=output.getvalue(),
            file_name=f"Official_Bill_{area}.xlsm",
            mime="application/vnd.ms-excel.sheet.macroenabled.12"
        )
    except Exception as e:
        st.error(f"Error loading template: {e}. Ensure the .xlsm file is in GitHub.")
