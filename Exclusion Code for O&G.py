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
        "GOGEL ID Unnamed: 51_level_1": "GOGEL ID",
        "GCEL ID Unnamed: 52_level_1": "GCEL ID",
        "Bloomberg BB Company ID": "BB Company ID",
        "Bloomberg BB ID": "BB ID",
        "Bloomberg BB FIGI": "BB FIGI",
        "ISIN Codes ISIN equity": "ISIN equity",
        "ISIN Codes ISINs bonds": "ISINs bonds",
        "LEI LEI": "LEI",
        "NACE NACE Classification": "NACE Classification",
        "NACE NACE Code": "NACE Code",
        "Unconventionals Tar Sands": "Tar Sand Revenue",
        "Unconventionals Arctic": "Arctic Revenue",
        "Unconventionals Coalbed Methane": "Coalbed Methane Revenue",
        "Primary Business Sectors Unnamed: 14_level_1": "Primary Business Sector",
        "Pipelines Length of Pipelines under Development": "Pipeline Expansion",
        "LNG Terminals Total Capacity under Development": "LNG Terminal Expansion"
    }
    
    df.rename(columns=column_mapping, inplace=True, errors='ignore')
    
    # Keep only required columns
    required_columns = list(column_mapping.values())
    df = df[required_columns]
    
    revenue_columns = ["Tar Sand Revenue", "Arctic Revenue", "Coalbed Methane Revenue"]
    for col in revenue_columns:
        df[col] = df[col].astype(str).str.replace('%', '', regex=True).str.replace(',', '', regex=True)
        df[col] = pd.to_numeric(df[col], errors='coerce')
    
    df[revenue_columns] = df[revenue_columns].fillna(0)
    if df[revenue_columns].max().max() <= 1:
        df[revenue_columns] = df[revenue_columns] * 100
    
    df["Total Exclusion Revenue"] = df[revenue_columns].sum(axis=1)
    
    # Apply separate thresholds for each sector
    sector_excluded = (
        (df["Tar Sand Revenue"] > tar_sand_threshold) |
        (df["Arctic Revenue"] > arctic_threshold) |
        (df["Coalbed Methane Revenue"] > coalbed_threshold) |
        (df["Total Exclusion Revenue"] > total_threshold)
    )
    
    retained_companies = df[~sector_excluded]
    excluded_companies = df[sector_excluded]
    
    # Level 2 Exclusion
    upstream_exclusion_keywords = ["Conventional Oil & Gas", "Unconventional Oil & Gas"]
    level2_excluded = retained_companies[
        retained_companies["Primary Business Sector"].astype(str).str.contains('|'.join(upstream_exclusion_keywords), case=False, na=False) |
        retained_companies["Pipeline Expansion"].notna() |
        retained_companies["LNG Terminal Expansion"].notna()
    ]
    level2_retained = retained_companies.drop(level2_excluded.index)
    
    # Statistics
    stats = {
        "Total Companies": len(df),
        "Retained Companies (Level 1)": len(retained_companies),
        "Excluded Companies (Level 1)": len(excluded_companies),
        "Excluded Companies (Level 2)": len(level2_excluded),
        "Retained Companies (Final)": len(level2_retained)
    }
    
    # Save to Excel in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        retained_companies.to_excel(writer, sheet_name="Retained Level 1", index=False)
        excluded_companies.to_excel(writer, sheet_name="Excluded Level 1", index=False)
        level2_excluded.to_excel(writer, sheet_name="Excluded Level 2", index=False)
        level2_retained.to_excel(writer, sheet_name="Retained Final", index=False)
    output.seek(0)
    
    return output, stats

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
