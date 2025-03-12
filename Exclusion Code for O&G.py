import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

def filter_companies_by_revenue(uploaded_file, tar_sand_exclude, tar_sand_threshold,
                                arctic_exclude, arctic_threshold,
                                coalbed_exclude, coalbed_threshold,
                                total_exclude, total_threshold):
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
    
    # Separate companies with no data
    companies_with_no_data = df[df[["Tar Sand Revenue", "Arctic Revenue", "Coalbed Methane Revenue"]].isnull().all(axis=1)]
    df = df.dropna(subset=["Tar Sand Revenue", "Arctic Revenue", "Coalbed Methane Revenue"], how='all')
    
    revenue_columns = ["Tar Sand Revenue", "Arctic Revenue", "Coalbed Methane Revenue"]
    for col in revenue_columns:
        df[col] = df[col].astype(str).str.replace('%', '', regex=True).str.replace(',', '', regex=True)
        df[col] = pd.to_numeric(df[col], errors='coerce')
    
    df[revenue_columns] = df[revenue_columns].fillna(0)
    if df[revenue_columns].max().max() <= 1:
        df[revenue_columns] = df[revenue_columns] * 100
    
    df["Total Exclusion Revenue"] = df[revenue_columns].sum(axis=1)
    
    # Apply exclusion logic per sector
    excluded_reasons = []
    for index, row in df.iterrows():
        reasons = []
        if tar_sand_exclude and (tar_sand_threshold == "" or row["Tar Sand Revenue"] > float(tar_sand_threshold)):
            reasons.append("Tar Sand Revenue Exceeded")
        if arctic_exclude and (arctic_threshold == "" or row["Arctic Revenue"] > float(arctic_threshold)):
            reasons.append("Arctic Revenue Exceeded")
        if coalbed_exclude and (coalbed_threshold == "" or row["Coalbed Methane Revenue"] > float(coalbed_threshold)):
            reasons.append("Coalbed Methane Revenue Exceeded")
        if total_exclude and (total_threshold == "" or row["Total Exclusion Revenue"] > float(total_threshold)):
            reasons.append("Total Exclusion Revenue Exceeded")
        excluded_reasons.append(", ".join(reasons) if reasons else "")
    
    df["Exclusion Reason"] = excluded_reasons
    retained_companies = df[df["Exclusion Reason"] == ""]
    excluded_companies = df[df["Exclusion Reason"] != ""]
    
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
    return exclude, threshold

tar_sand_exclude, tar_sand_threshold = sector_exclusion_input("Tar Sand")
arctic_exclude, arctic_threshold = sector_exclusion_input("Arctic")
coalbed_exclude, coalbed_threshold = sector_exclusion_input("Coalbed Methane")
total_exclude, total_threshold = sector_exclusion_input("Total Revenue")

if st.sidebar.button("Run Filtering Process"):
    if uploaded_file:
        filtered_output, stats = filter_companies_by_revenue(
            uploaded_file, tar_sand_exclude, tar_sand_threshold,
            arctic_exclude, arctic_threshold,
            coalbed_exclude, coalbed_threshold,
            total_exclude, total_threshold)
        
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
