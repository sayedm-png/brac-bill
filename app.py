import streamlit as st
import pandas as pd
from io import BytesIO
import datetime
import re

# --- 1. Configuration & Data Fetching ---
# Your provided public Google Drive link
DRIVE_URL = "https://drive.google.com/file/d/1yUpT0N-8d_P1org-e5LwfffCYxco2sPb/view?ts=69dc8c08"

# Official BRAC Rates from your requirements
RATES = {
    "Breakfast": 70,
    "Lunch": 140,
    "Dinner": 140,
    "NightHalt": 150
}

@st.cache_data(ttl=3600)  # Refreshes data from Drive every hour
def fetch_allowance_data(url):
    try:
        # Convert view link to direct download link
        file_id = re.search(r'[-\w]{25,}', url).group()
        direct_download = f'https://docs.google.com/spreadsheets/d/{file_id}/export?format=csv'
        df = pd.read_csv(direct_download)
        # Cleaning column names just in case
        df.columns = [c.strip() for c in df.columns]
        return df
    except Exception as e:
        return None

# --- 2. App Interface ---
st.set_page_config(page_title="Field Bill Pro", layout="wide")
st.title("💸 Field Visit Billing System")
st.caption("Frontier Tech & Impact Innovation Unit | BRAC")

# Fetch data from Google Drive
allowance_df = fetch_allowance_data(DRIVE_URL)

with st.sidebar:
    st.header("1. Trip Details")
    visit_date = st.date_input("Date of Visit", datetime.date.today())
    
    if allowance_df is not None:
        # Assumes your file has columns 'Area' and 'Allowance'
        area_list = allowance_df['Area'].unique().tolist()
        selected_area = st.selectbox("Select Area of Visit", options=area_list)
        fixed_dist = allowance_df.loc[allowance_df['Area'] == selected_area, 'Allowance'].values[0]
    else:
        st.error("Could not load distance data. Please enter manually.")
        selected_area = st.text_input("Area Name")
        fixed_dist = st.number_input("Fixed Distance Allowance", min_value=0)

    travel_from = st.text_input("From", value="Dhaka")
    travel_to = st.text_input("To", value=selected_area)
    ground_cost = st.number_input("Ground Travel Cost (Actual)", min_value=0)

    st.header("2. Subsistence")
    b_count = st.number_input("Breakfasts", min_value=0)
    l_count = st.number_input("Lunches", min_value=0)
    d_count = st.number_input("Dinners", min_value=0)
    halt_nights = st.number_input("Night Haltage (Nights)", min_value=0)

# --- 3. Calculation Engine ---
food_total = (b_count * RATES["Breakfast"]) + (l_count * RATES["Lunch"]) + (d_count * RATES["Dinner"])
halt_total = halt_nights * RATES["NightHalt"]
grand_total = ground_cost + fixed_dist + food_total + halt_total

# --- 4. Live Preview (The "Xero" View) ---
col_preview, col_export = st.columns([2, 1])

with col_preview:
    st.subheader("Billing Preview")
    summary_data = {
        "Description": ["Ground Travel", "Distance Allowance", "Daily Allowance (Meals)", "Night Haltage"],
        "Breakdown": ["Actual Cost", f"Fixed for {selected_area}", f"{b_count}B, {l_count}L, {d_count}D", f"{halt_nights} Night(s)"],
        "Amount (BDT)": [ground_cost, fixed_dist, food_total, halt_total]
    }
    st.table(pd.DataFrame(summary_data))
    st.metric("Total Bill Amount", f"{grand_total} BDT")

# --- 5. Professional Excel Generation ---
def generate_excel():
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet('Bill')
        
        # Formats
        bold = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#f2f2f2'})
        center = workbook.add_format({'align': 'center', 'border': 1})
        money = workbook.add_format({'num_format': '#,##0', 'border': 1})
        border = workbook.add_format({'border': 1})

        # Header Info
        worksheet.write('A1', 'FIELD VISIT BILL REPORT', workbook.add_format({'bold': True, 'size': 14}))
        worksheet.write('A3', 'Date:', bold); worksheet.write('B3', str(visit_date), border)
        worksheet.write('A4', 'Area:', bold); worksheet.write('B4', selected_area, border)
        worksheet.write('A5', 'Route:', bold); worksheet.write('B5', f"{travel_from} to {travel_to}", border)

        # Table Headers
        cols = ['Description', 'Details', 'Amount (BDT)']
        for i, col in enumerate(cols):
            worksheet.write(7, i, col, bold)

        # Content
        content = [
            ["Ground Travel", "Local Conveyance", ground_cost],
            ["Distance Allowance", f"Fixed rate for {selected_area}", fixed_dist],
            ["Daily Allowance", f"Meals (B:{b_count}, L:{l_count}, D:{d_count})", food_total],
            ["Night Haltage", f"{halt_nights} Night(s) stay", halt_total],
            ["GRAND TOTAL", "", grand_total]
        ]
        
        for r, row in enumerate(content):
            worksheet.write(r + 8, 0, row[0], border)
            worksheet.write(r + 8, 1, row[1], border)
            worksheet.write(r + 8, 2, row[2], money)

        # Signatures
        worksheet.write('A18', '_____________________')
        worksheet.write('A19', 'Claimant Signature')
        worksheet.write('C18', '_____________________')
        worksheet.write('C19', 'Authorized Approval')
        
    return output.getvalue()

with col_export:
    st.subheader("Export")
    st.download_button(
        label="📥 Download Final Bill (Excel)",
        data=generate_excel(),
        file_name=f"Bill_{selected_area}_{visit_date}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
