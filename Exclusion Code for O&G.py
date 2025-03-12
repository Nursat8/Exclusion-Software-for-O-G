import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

def filter_companies_by_revenue(uploaded_file, sector_exclusions, total_thresholds):
    if uploaded_file is None:
        return None, None
    
    # Load the Excel file
    xls = pd.ExcelFile(uploaded_file)
    df = xls.parse("All Companies", header=[3, 4])
    
    # Flatten multi-level columns
    df.columns = [' '.join(map(str, col)).strip() for col in df.columns]
    
    # Column Mapping
    column_mapping = {
        "Company Unnamed: 11_level_1": "Company",
        "Bloomberg BB Ticker": "BB Ticker",
        "ISIN Codes ISIN equity": "ISIN equity",
        "LEI LEI": "LEI",
        "Unconventionals Fracking": "Fracking Revenue",
        "Unconventionals Tar Sands": "Tar Sand Revenue",
        "Unconventionals Coalbed Methane": "Coalbed Methane Revenue",
        "Unconventionals Extra Heavy Oil": "Extra Heavy Oil Revenue",
        "Unconventionals Ultra Deepwater": "Ultra Deepwater Revenue",
        "Unconventionals Arctic": "Arctic Revenue",
        "Unconventional Production Unnamed: 25_level_1": "Unconventional Production Revenue"
    }
    
    df.rename(columns=column_mapping, inplace=True, errors='ignore')
    
    # Keep only required columns
    required_columns = list(column_mapping.values()) + ["Exclusion Reason"]
    df = df[list(column_mapping.values())]
    
    # Separate companies with no data
    companies_with_no_data = df[df[list(column_mapping.values())[4:]].isnull().all(axis=1)]
    df = df.dropna(subset=list(column_mapping.values())[4:], how='all')
    
    revenue_columns = list(column_mapping.values())[4:]
    for col in revenue_columns:
        df[col] = df[col].astype(str).str.replace('%', '', regex=True).str.replace(',', '', regex=True)
        df[col] = pd.to_numeric(df[col], errors='coerce')
    
    df[revenue_columns] = df[revenue_columns].fillna(0)
    if df[revenue_columns].max().max() <= 1:
        df[revenue_columns] = df[revenue_columns] * 100
    
    # Calculate total exclusion revenues for selected sectors
    for key, threshold_data in total_thresholds.items():
        selected_sectors = threshold_data["sectors"]
        threshold_value = threshold_data["threshold"]
        valid_sectors = [sector for sector in selected_sectors if sector in df.columns]
        if valid_sectors:
            df[key] = df[valid_sectors].sum(axis=1)
    
    # Apply exclusion logic per sector
    excluded_reasons = []
    for index, row in df.iterrows():
        reasons = []
        for sector, (exclude, threshold) in sector_exclusions.items():
            if exclude and (threshold == "" or row[sector] > float(threshold)):
                reasons.append(f"{sector} Revenue Exceeded")
        for key, threshold_data in total_thresholds.items():
            threshold_value = threshold_data["threshold"]
            if key in df.columns and row[key] > float(threshold_value):
                reasons.append(f"{key} Revenue Exceeded")
        excluded_reasons.append(", ".join(reasons) if reasons else "")
    
    df["Exclusion Reason"] = excluded_reasons
    retained_companies = df[df["Exclusion Reason"] == ""]
    excluded_companies = df[df["Exclusion Reason"] != ""]
    
    # Ensure companies_with_no_data has all required columns
    for col in required_columns:
        if col not in companies_with_no_data.columns:
            companies_with_no_data[col] = np.nan
    
    # Remove unnecessary columns from output
    retained_companies = retained_companies[required_columns]
    excluded_companies = excluded_companies[required_columns]
    companies_with_no_data = companies_with_no_data[required_columns]
    
    # Save to Excel in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        retained_companies.to_excel(writer, sheet_name="Retained Companies", index=False)
        excluded_companies.to_excel(writer, sheet_name="Excluded Companies", index=False)
        companies_with_no_data.to_excel(writer, sheet_name="No Data Companies", index=False)
    output.seek(0)
    
    return output, {
        "Total Companies": len(df) + len(companies_with_no_data),
        "Retained Companies": len(retained_companies),
        "Excluded Companies": len(excluded_companies),
        "Companies with No Data": len(companies_with_no_data)
    }

# Streamlit UI
st.title("Company Revenue Filter")
st.write("Upload an Excel file and set exclusion thresholds.")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

st.sidebar.header("Set Exclusion Criteria")

def sector_exclusion_input(sector_name):
    exclude = st.sidebar.checkbox(f"Exclude {sector_name}", value=False)
    threshold = ""
    if exclude:
        threshold = st.sidebar.text_input(f"{sector_name} Revenue Threshold (%)", "")
    return sector_name, (exclude, threshold)

sector_exclusions = dict([
    sector_exclusion_input("Fracking Revenue"),
    sector_exclusion_input("Tar Sand Revenue"),
    sector_exclusion_input("Coalbed Methane Revenue"),
    sector_exclusion_input("Extra Heavy Oil Revenue"),
    sector_exclusion_input("Ultra Deepwater Revenue"),
    sector_exclusion_input("Arctic Revenue"),
    sector_exclusion_input("Unconventional Production Revenue")
])

st.sidebar.header("Set Multiple Custom Total Revenue Thresholds")
total_thresholds = {}
num_custom_thresholds = st.sidebar.number_input("Number of Custom Total Thresholds", min_value=1, max_value=5, value=1)
for i in range(num_custom_thresholds):
    selected_sectors = st.sidebar.multiselect(f"Select Sectors for Custom Threshold {i+1}", list(sector_exclusions.keys()), key=f"sectors_{i}")
    total_threshold = st.sidebar.text_input(f"Total Revenue Threshold {i+1} (%)", "", key=f"threshold_{i}")
    if selected_sectors and total_threshold:
        total_thresholds[f"Custom Total Revenue {i+1}"] = {"sectors": selected_sectors, "threshold": total_threshold}

if st.sidebar.button("Run Filtering Process"):
    if uploaded_file:
        filtered_output, stats = filter_companies_by_revenue(uploaded_file, sector_exclusions, total_thresholds)
        
        if filtered_output:
            st.success("File processed successfully!")
            
            # Display statistics
            st.subheader("Processing Statistics")
            for key, value in stats.items():
                st.write(f"**{key}:** {value}")
            
            # Download button
            st.download_button(
                label="Download Filtered Excel",
                data=filtered_output,
                file_name="filtered_companies.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
