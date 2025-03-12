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
    
    # Return dataframes, plus stats
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
    # Original snippet function
    return pd.read_excel(file, sheet_name=sheet_name, header=4)

def filter_exclusions(upstream_df, midstream_df):
    # Check shape to prevent ValueError if columns are missing
    if upstream_df.shape[1] < 47:
        raise ValueError(
            f"Upstream sheet has only {upstream_df.shape[1]} columns. "
            "Expected at least 47 so we can index [5, 27, 41, 42, 46]."
        )
    if midstream_df.shape[1] < 12:
        raise ValueError(
            f"Midstream sheet has only {midstream_df.shape[1]} columns. "
            "Expected at least 12 so we can index [5, 8, 9, 10, 11]."
        )
    
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

##########################################
# EXTENDED LOGIC: LEVEL 2 RETAINED + STATS
##########################################

def filter_exclusions_with_retained(upstream_df, midstream_df):
    """
    Same logic as snippet's filter_exclusions, but also determine which
    companies are retained. Then return stats for level 2.
    """
    # Check shape first
    if upstream_df.shape[1] < 47:
        raise ValueError(
            f"Upstream sheet has only {upstream_df.shape[1]} columns. "
            "Expected at least 47 so we can index [5, 27, 41, 42, 46]."
        )
    if midstream_df.shape[1] < 12:
        raise ValueError(
            f"Midstream sheet has only {midstream_df.shape[1]} columns. "
            "Expected at least 12 so we can index [5, 8, 9, 10, 11]."
        )
    
    # Snippet lines
    upstream_df = upstream_df.iloc[:, [5, 27, 41, 42, 46]]
    upstream_df.columns = ["Company", "Fossil Fuel Share of Revenue", "BB Ticker", "ISIN Equity", "LEI"]
    
    midstream_df = midstream_df.iloc[:, [5, 8, 9, 10, 11]]
    midstream_df.columns = [
        "Company",
        "Length of Pipelines under Development",
        "Liquefaction Capacity (Export)",
        "Regasification Capacity (Import)",
        "Total Capacity under Development"
    ]
    
    # Convert numeric
    upstream_df["Fossil Fuel Share of Revenue"] = pd.to_numeric(
        upstream_df["Fossil Fuel Share of Revenue"].astype(str).str.replace('%', ''),
        errors='coerce'
    ).fillna(0)
    
    # Upstream exclusion
    upstream_exclusion = upstream_df["Fossil Fuel Share of Revenue"] > 0
    
    # Midstream exclusion
    # If any of [Length..., Liquefaction..., Regasification..., Total Capacity...]
    # is non-null (notna), we exclude. For numeric safety, convert them to numeric:
    for c in [
        "Length of Pipelines under Development",
        "Liquefaction Capacity (Export)",
        "Regasification Capacity (Import)",
        "Total Capacity under Development"
    ]:
        midstream_df[c] = pd.to_numeric(midstream_df[c], errors='coerce').fillna(0)
    
    midstream_exclusion = midstream_df.iloc[:, 1:].apply(lambda row: any(row != 0), axis=1)
    
    # Mark exclusion reason
    upstream_df["Exclusion Reason"] = ""
    upstream_df.loc[upstream_exclusion, "Exclusion Reason"] = "Upstream - Fossil Fuel Revenue > 0%"
    
    midstream_df["Exclusion Reason"] = ""
    midstream_df.loc[midstream_exclusion, "Exclusion Reason"] = "Midstream Expansion - Capacity in Development"
    
    # Merge to find full union of companies (by Company).
    # If you do have unique IDs like [Company, BB Ticker, ISIN, LEI], you can do multi-column merges.
    # We'll do a left outer merge on 'Company' from the upstream side for demonstration:
    merged_df = pd.merge(
        upstream_df,
        midstream_df,
        on="Company",
        how="outer",
        suffixes=("_up", "_mid")
    )
    # Build final columns (filling if they don't exist):
    for c in ["BB Ticker", "ISIN Equity", "LEI"]:
        if c+"_up" in merged_df.columns:
            merged_df[c] = merged_df[c+"_up"]
        elif c+"_mid" in merged_df.columns:
            merged_df[c] = merged_df[c+"_mid"]
        else:
            merged_df[c] = np.nan
    
    # Combine Exclusion Reason from upstream + midstream if both exist

    def combine_reasons(r1, r2):
        """Safely combine reason strings, handling NaN or None by using empty strings."""
        # Convert None/NaN to empty string
        r1 = str(r1) if pd.notna(r1) else ""
        r2 = str(r2) if pd.notna(r2) else ""
    
        if r1 and r2:
            return r1 + ", " + r2
    else:
        # Return whichever one isn't blank (or both blank if both are blank)
        return r1 or r2
    
    merged_df["Exclusion Reason"] = merged_df.apply(
        lambda row: combine_reasons(row["Exclusion Reason_up"], row["Exclusion Reason_mid"]), axis=1
    )
    
    # Keep numeric columns from each side:
    # Upstream: Fossil Fuel Share of Revenue
    # Midstream: Length..., Liquefaction..., Regasification..., Total...
    # We'll rename them nicely:
    merged_df["Fossil Fuel Share of Revenue"] = pd.to_numeric(
        merged_df["Fossil Fuel Share of Revenue_up"], errors='coerce'
    ).fillna(0)
    
    # Already numeric from above
    merged_df["Length of Pipelines under Development"] = merged_df["Length of Pipelines under Development"]
    merged_df["Liquefaction Capacity (Export)"] = merged_df["Liquefaction Capacity (Export)"]
    merged_df["Regasification Capacity (Import)"] = merged_df["Regasification Capacity (Import)"]
    merged_df["Total Capacity under Development"] = merged_df["Total Capacity under Development"]
    
    # Define final columns
    final_cols = [
        "Company", "BB Ticker", "ISIN Equity", "LEI",
        "Fossil Fuel Share of Revenue",
        "Length of Pipelines under Development",
        "Liquefaction Capacity (Export)",
        "Regasification Capacity (Import)",
        "Total Capacity under Development",
        "Exclusion Reason"
    ]
    
    # We rely on the new columns we set or re-labeled
    final_df = merged_df[final_cols].copy()
    
    # Excluded => if Exclusion Reason is not empty
    excluded_mask = final_df["Exclusion Reason"].str.strip() != ""
    level2_excluded = final_df[excluded_mask].copy()
    level2_retained = final_df[~excluded_mask].copy()
    
    # Build stats
    stats_level2 = {
        "Total Companies (Level 2)": len(final_df),
        "Excluded (Level 2)": len(level2_excluded),
        "Retained (Level 2)": len(level2_retained)
    }
    return level2_excluded, level2_retained, stats_level2


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
        
        # Show DataFrames (Level 1) in the UI (optional)
        st.subheader("Level 1: Retained Companies")
        st.dataframe(retained_companies)
        st.subheader("Level 1: Excluded Companies")
        st.dataframe(excluded_companies)
        st.subheader("Level 1: No Data Companies")
        st.dataframe(companies_with_no_data)
        
        # -- LEVEL 2 FILTERS (Upstream + Midstream) using snippet logic
        try:
            upstream_df = load_data(uploaded_file, sheet_name="Upstream", header_row=4)
            midstream_df = load_data(uploaded_file, sheet_name="Midstream Expansion", header_row=4)
            
            # Original snippet code: shows just the Excluded
            old_excluded_data = filter_exclusions(upstream_df.copy(), midstream_df.copy())
            st.subheader("Level 2 Excluded Companies (Original Snippet Logic)")
            st.dataframe(old_excluded_data)
            
            # NEW code: Excluded + Retained + Stats
            level2_excluded_data, level2_retained_data, stats2 = filter_exclusions_with_retained(
                upstream_df.copy(),
                midstream_df.copy()
            )
            
            st.subheader("Level 2 Processing Statistics")
            for k, v in stats2.items():
                st.write(f"**{k}:** {v}")
            
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
        except ValueError as ve:
            st.error(f"Error with Level 2 data format: {ve}")
        
    else:
        st.error("Please upload an Excel file before running the filtering process.")
