import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import io

############################
# BEGIN FIRST SNIPPET CODE #
############################

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
    
    return (retained_companies, excluded_companies, companies_with_no_data), {
        "Total Companies": len(df) + len(companies_with_no_data),
        "Retained Companies": len(retained_companies),
        "Excluded Companies": len(excluded_companies),
        "Companies with No Data": len(companies_with_no_data)
    }

# Streamlit UI for FIRST SNIPPET
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
    selected_sectors = st.sidebar.multiselect(
        f"Select Sectors for Custom Threshold {i+1}",
        list(sector_exclusions.keys()),
        key=f"sectors_{i}"
    )
    total_threshold = st.sidebar.text_input(f"Total Revenue Threshold {i+1} (%)", "", key=f"threshold_{i}")
    if selected_sectors and total_threshold:
        total_thresholds[f"Custom Total Revenue {i+1}"] = {"sectors": selected_sectors, "threshold": total_threshold}

############################
# BEGIN SECOND SNIPPET CODE
############################

def load_data(file, sheet_name, header_row):
    # We keep this function exactly as is, adjusting row based on snippet requirement
    return pd.read_excel(file, sheet_name=sheet_name, header=4)

def filter_exclusions(upstream_df, midstream_df):
    # Select correct columns using index positions
    upstream_df = upstream_df.iloc[:, [5, 27, 41, 42, 46]]  # Company, AB, AP, AQ, AU
    upstream_df.columns = ["Company", "Fossil Fuel Share of Revenue", "BB Ticker", "ISIN Equity", "LEI"]
    
    midstream_df = midstream_df.iloc[:, [5, 8, 9, 10, 11]]  # Company, I, J, K, L
    midstream_df.columns = [
        "Company",
        "Length of Pipelines under Development",
        "Liquefaction Capacity (Export)",
        "Regasification Capacity (Import)",
        "Total Capacity under Development"
    ]
    
    # Convert Fossil Fuel Share of Revenue to numeric, handling errors
    upstream_df["Fossil Fuel Share of Revenue"] = pd.to_numeric(
        upstream_df["Fossil Fuel Share of Revenue"].astype(str).str.replace('%', ''),
        errors='coerce'
    ).fillna(0)  # Replace NaN with 0
    
    # Identify exclusion criteria
    upstream_exclusion = upstream_df["Fossil Fuel Share of Revenue"] > 0
    midstream_exclusion = midstream_df.iloc[:, 1:].notna().any(axis=1)  # Check if any midstream column has a value
    
    # Create exclusion reason
    upstream_df["Exclusion Reason"] = ""
    upstream_df.loc[upstream_exclusion, "Exclusion Reason"] = "Upstream - Fossil Fuel Revenue > 0%"
    midstream_df["Exclusion Reason"] = ""
    midstream_df.loc[midstream_exclusion, "Exclusion Reason"] = "Midstream Expansion - Capacity in Development"
    
    # Combine data
    excluded_companies = pd.concat([
        upstream_df.loc[upstream_exclusion, [
            "Company", "BB Ticker", "ISIN Equity", "LEI",
            "Fossil Fuel Share of Revenue", "Exclusion Reason"
        ]],
        midstream_df.loc[midstream_exclusion, [
            "Company", "Exclusion Reason", "Length of Pipelines under Development",
            "Liquefaction Capacity (Export)", "Regasification Capacity (Import)",
            "Total Capacity under Development"
        ]]
    ], ignore_index=True)
    
    return excluded_companies

########################################################
# NEW FUNCTION: FULL MERGE FOR LEVEL 2 RETAINED vs. EXCL
########################################################

def process_level2_exclusions(upstream_df, midstream_df):
    """
    This function merges Upstream & Midstream data by [Company, BB Ticker, ISIN Equity, LEI],
    determines if each company is excluded or retained, and returns two DataFrames:
        1. level2_excluded
        2. level2_retained
    based on the logic that:
    - Upstream is excluded if 'Fossil Fuel Share of Revenue' > 0.
    - Midstream is excluded if any capacity columns have a non-0, non-null value.
    - 'Exclusion Reason' merges both possible reasons if they apply.
    """
    # Upstream with consistent columns
    up = upstream_df.copy()
    up.columns = ["Company", "Fossil Fuel Share of Revenue", "BB Ticker", "ISIN Equity", "LEI"]
    
    # Midstream with consistent columns
    mid = midstream_df.copy()
    mid.columns = [
        "Company",
        "Length of Pipelines under Development",
        "Liquefaction Capacity (Export)",
        "Regasification Capacity (Import)",
        "Total Capacity under Development"
    ]
    
    # Merge on [Company, BB Ticker, ISIN Equity, LEI] if possible, but the original snippet
    # only has these columns in Upstream. We'll do left merges carefully.
    # Upstream has [Company, Fossil Fuel..., BB Ticker, ISIN Equity, LEI]
    # Midstream has [Company, Length..., Liquefaction..., Regasification..., Total...]
    # We unify them by 'Company' only, or if you truly have BB Ticker, ISIN Equity, LEI in Midstream as well.
    # The snippet doesn't show that midstream has BB Ticker, ISIN Equity, LEI columns.
    # => We'll merge only on "Company".
    # Adjust if your real data also has matching tickers/ISINs:
    merged = pd.merge(
        up,
        mid,
        on="Company",
        how="outer",
        suffixes=("_up", "_mid")
    )
    
    # If you do have columns 'BB Ticker','ISIN Equity','LEI' in midstream in real data, you can do:
    # how="outer", left_on=["Company","BB Ticker","ISIN Equity","LEI"], ...
    # For demonstration, we'll keep the snippet's structure and rely on 'Company' alone.
    
    # Convert the capacity columns to numeric so we can check > 0
    capacity_cols = [
        "Length of Pipelines under Development",
        "Liquefaction Capacity (Export)",
        "Regasification Capacity (Import)",
        "Total Capacity under Development"
    ]
    for c in capacity_cols:
        merged[c] = pd.to_numeric(merged[c], errors="coerce").fillna(0)
    
    # Convert Fossil Fuel Share of Revenue to numeric
    merged["Fossil Fuel Share of Revenue"] = pd.to_numeric(
        merged["Fossil Fuel Share of Revenue"],
        errors="coerce"
    ).fillna(0)
    
    # Exclusion conditions
    upstream_condition = merged["Fossil Fuel Share of Revenue"] > 0
    midstream_condition = (
        (merged["Length of Pipelines under Development"] != 0) |
        (merged["Liquefaction Capacity (Export)"] != 0) |
        (merged["Regasification Capacity (Import)"] != 0) |
        (merged["Total Capacity under Development"] != 0)
    )
    
    # Build Exclusion Reason
    reasons = []
    for idx, row in merged.iterrows():
        reason_list = []
        if row["Fossil Fuel Share of Revenue"] > 0:
            reason_list.append("Upstream - Fossil Fuel Revenue > 0%")
        # Check if any midstream expansions:
        if any(row[col] != 0 for col in capacity_cols):
            reason_list.append("Midstream Expansion - Capacity in Development")
        reasons.append(", ".join(reason_list) if reason_list else "")
    
    merged["Exclusion Reason"] = reasons
    
    # Keep columns: Company, BB Ticker, ISIN Equity, LEI,
    # Fossil Fuel Share of Revenue, capacity columns, Exclusion Reason
    # We'll fill BB Ticker, ISIN Equity, LEI from the Upstream side if midstream doesn't have them
    final_cols = [
        "Company", "BB Ticker", "ISIN Equity", "LEI",
        "Fossil Fuel Share of Revenue"
    ] + capacity_cols + ["Exclusion Reason"]
    
    # If 'BB Ticker', 'ISIN Equity', 'LEI' don't exist in midstream, that's fine. Already in 'up'.
    for c in ["BB Ticker", "ISIN Equity", "LEI"]:
        if c not in merged.columns:
            merged[c] = np.nan
    
    level2_full = merged[final_cols].copy()
    
    excluded_mask = upstream_condition | midstream_condition
    level2_excluded = level2_full[excluded_mask].copy()
    level2_retained = level2_full[~excluded_mask].copy()
    
    return level2_excluded, level2_retained

############################
# SINGLE RUN BUTTON LOGIC  #
############################

if st.sidebar.button("Run Filtering Process"):
    if uploaded_file:
        # -- LEVEL 1 FILTERS (Unconventionals) --
        (retained_companies, excluded_companies, companies_with_no_data), stats = filter_companies_by_revenue(
            uploaded_file, sector_exclusions, total_thresholds
        )
        
        # Display stats in the main UI
        st.success("File processed for Level 1!")
        st.subheader("Level 1 Processing Statistics")
        for key, value in stats.items():
            st.write(f"**{key}:** {value}")
        
        # -- Show DataFrames (Level 1) in the UI (optional) --
        st.subheader("Level 1: Retained Companies")
        st.dataframe(retained_companies)
        st.subheader("Level 1: Excluded Companies")
        st.dataframe(excluded_companies)
        st.subheader("Level 1: No Data Companies")
        st.dataframe(companies_with_no_data)
        
        # -- LEVEL 2 FILTERS (Upstream + Midstream) using original snippet logic --
        # We'll show them, but also produce the new logic with Retained vs. Excluded
        upstream_df = load_data(uploaded_file, sheet_name="Upstream", header_row=4)
        midstream_df = load_data(uploaded_file, sheet_name="Midstream Expansion", header_row=4)
        
        old_excluded_data = filter_exclusions(upstream_df, midstream_df)
        st.subheader("Level 2 Excluded Companies (Original Snippet Logic)")
        st.dataframe(old_excluded_data)
        
        # -- NEW LOGIC for FULL MERGE & RETAINED vs EXCLUDED (Level 2) --
        level2_excluded_data, level2_retained_data = process_level2_exclusions(upstream_df, midstream_df)
        
        st.subheader("Level 2 Excluded (Improved Logic)")
        st.dataframe(level2_excluded_data)
        
        st.subheader("Level 2 Retained (Improved Logic)")
        st.dataframe(level2_retained_data)
        
        # -- COMBINE ALL RESULTS INTO A SINGLE EXCEL --
        output_combined = BytesIO()
        with pd.ExcelWriter(output_combined, engine='xlsxwriter') as writer:
            # From Level 1
            retained_companies.to_excel(writer, sheet_name="Retained Companies", index=False)
            excluded_companies.to_excel(writer, sheet_name="Excluded Companies", index=False)
            companies_with_no_data.to_excel(writer, sheet_name="No Data Companies", index=False)
            
            # From Level 2 (New Logic)
            level2_excluded_data.to_excel(writer, sheet_name="Level 2 Excluded", index=False)
            level2_retained_data.to_excel(writer, sheet_name="Level 2 Retained", index=False)
        
        output_combined.seek(0)
        
        # Single Download Button for the Combined File
        st.download_button(
            label="Download Combined Excel (Level 1 + 2)",
            data=output_combined,
            file_name="all_exclusions.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Please upload an Excel file before running the filtering process.")
