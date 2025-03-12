import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

def filter_companies_by_revenue(uploaded_file, tar_sand_threshold, arctic_threshold, coalbed_threshold, total_threshold):
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
        "Unconventionals Tar Sands": "Tar Sand Revenue",
        "Unconventionals Arctic": "Arctic Revenue",
        "Unconventionals Coalbed Methane": "Coalbed Methane Revenue"
    }
    
    df.rename(columns=column_mapping, inplace=True, errors='ignore')
    
    # Keep only required columns
    required_columns = list(column_mapping.values()) + ["Exclusion Reason"]
    df = df[list(column_mapping.values())]
    
    # Remove companies with no data (empty cells in exclusion-related columns)
    df = df.dropna(subset=["Tar Sand Revenue", "Arctic Revenue", "Coalbed Methane Revenue"], how='all')
    
    revenue_columns = ["Tar Sand Revenue", "Arctic Revenue", "Coalbed Methane Revenue"]
    for col in revenue_columns:
        df[col] = df[col].astype(str).str.replace('%', '', regex=True).str.replace(',', '', regex=True)
        df[col] = pd.to_numeric(df[col], errors='coerce')
    
    df[revenue_columns] = df[revenue_columns].fillna(0)
    if df[revenue_columns].max().max() <= 1:
        df[revenue_columns] = df[revenue_columns] * 100
    
    df["Total Exclusion Revenue"] = df[revenue_columns].sum(axis=1)
    
    # Apply separate thresholds for each sector
    excluded_reasons = []
    for index, row in df.iterrows():
        reasons = []
        if row["Tar Sand Revenue"] > tar_sand_threshold:
            reasons.append("Tar Sand Revenue Exceeded")
        if row["Arctic Revenue"] > arctic_threshold:
            reasons.append("Arctic Revenue Exceeded")
        if row["Coalbed Methane Revenue"] > coalbed_threshold:
            reasons.append("Coalbed Methane Revenue Exceeded")
        if row["Total Exclusion Revenue"] > total_threshold:
            reasons.append("Total Exclusion Revenue Exceeded")
        excluded_reasons.append(", ".join(reasons) if reasons else "")
    
    df["Exclusion Reason"] = excluded_reasons
    retained_companies = df[df["Exclusion Reason"] == ""]
    excluded_companies = df[df["Exclusion Reason"] != ""]
    
    # Remove unnecessary columns from output
    retained_companies = retained_companies[required_columns]
    excluded_companies = excluded_companies[required_columns]
    
    # Save to Excel in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        retained_companies.to_excel(writer, sheet_name="Retained Companies", index=False)
        excluded_companies.to_excel(writer, sheet_name="Excluded Companies", index=False)
    output.seek(0)
    
    return output, {
        "Total Companies": len(df),
        "Retained Companies": len(retained_companies),
        "Excluded Companies": len(excluded_companies)
    }

# Streamlit UI
st.title("Company Revenue Filter")
st.write("Upload an Excel file and set exclusion thresholds.")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

st.sidebar.header("Set Exclusion Thresholds")
tar_sand_threshold = st.sidebar.text_input("Tar Sand Revenue Threshold (%)", "20")
arctic_threshold = st.sidebar.text_input("Arctic Revenue Threshold (%)", "20")
coalbed_threshold = st.sidebar.text_input("Coalbed Methane Revenue Threshold (%)", "20")
total_threshold = st.sidebar.text_input("Total Revenue Threshold (%)", "20")

# Convert inputs to numeric values
tar_sand_threshold = float(tar_sand_threshold) if tar_sand_threshold else 20
arctic_threshold = float(arctic_threshold) if arctic_threshold else 20
coalbed_threshold = float(coalbed_threshold) if coalbed_threshold else 20
total_threshold = float(total_threshold) if total_threshold else 20

if st.sidebar.button("Run Filtering Process"):
    if uploaded_file:
        filtered_output, stats = filter_companies_by_revenue(uploaded_file, tar_sand_threshold, arctic_threshold, coalbed_threshold, total_threshold)
        
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
