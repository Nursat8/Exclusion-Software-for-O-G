import re
import pandas as pd
import numpy as np
import io
from io import BytesIO
import streamlit as st

##############################################
# Combined Renaming for Both Level 1 and Level 2
##############################################
def flatten_multilevel_columns(df):
    """Flatten multi-level column headers into single strings."""
    df.columns = [" ".join(map(str, col)).strip() for col in df.columns]
    return df

def find_column(df, possible_matches, how="exact", required=True):
    """
    Searches df.columns for the first column that matches any of the possible_matches.
    """
    for col in df.columns:
        for pattern in possible_matches:
            if how == "exact":
                if col.strip().lower() == pattern.strip().lower():
                    return col
            elif how == "partial":
                if pattern.strip().lower() in col.lower():
                    return col
            elif how == "regex":
                if re.search(pattern, col, flags=re.IGNORECASE):
                    return col
    if required:
        raise ValueError(
            f"Could not find a required column. Tried {possible_matches} in columns: {df.columns.tolist()}"
        )
    return None

def rename_columns(df, rename_map, how="exact"):
    """
    Given a dictionary { new_col_name: [list of possible appearances] },
    search & rename them in the DataFrame if found.
    """
    for new_col_name, patterns in rename_map.items():
        old_name = find_column(df, patterns, how=how, required=False)
        if old_name:
            df.rename(columns={old_name: new_col_name}, inplace=True)
    return df

##############################################
# Combined Filter Function (Levels 1 & 2)
##############################################
def filter_companies_combined(uploaded_file, sector_exclusions, total_thresholds, level2_thresholds):
    """
    Reads the Excel file, applies Level 1 (revenue) and Level 2 (pipeline/capacity) exclusion logic,
    and returns an Excel file (in memory) plus summary stats.
    
    - sector_exclusions and total_thresholds: used for revenue (Level 1) thresholds.
    - level2_thresholds: a dict containing:
         "apply_upstream": (bool) exclude if 'GOGEL Tab' contains "upstream"
         "Length of Pipelines under Development": numeric threshold (e.g., 10)
         "Liquefaction Capacity (Export)": numeric threshold
         "Regasification Capacity (Import)": numeric threshold
         "Total Capacity under Development": numeric threshold
    """
    if uploaded_file is None:
        return None, None

    # --- 1) Read and Flatten ---
    xls = pd.ExcelFile(uploaded_file)
    df = xls.parse("All Companies", header=[3,4])
    df = flatten_multilevel_columns(df)
    # (If needed, shift header row up; here we assume row 4 is header after flattening)
    # df = df.iloc[1:].reset_index(drop=True)

    # --- 2) Combined Renaming ---
    # Create one combined rename map that includes both revenue and pipeline/capacity columns.
    rename_map = {
        # Company Info
        "Company": ["company name", "company"],
        "BB Ticker": ["bloomberg bb ticker", "bb ticker"],
        "ISIN equity": ["isin codes isin equity", "isin equity"],
        "LEI": ["lei lei", "lei", "legal entity identifier"],
        # Level 1 (Revenue) Columns
        "Hydrocarbons Production (%)": ["hydrocarbons production", "hydrocarbons"],
        "Fracking Revenue": ["fracking", "fracking revenue"],
        "Tar Sand Revenue": ["tar sands", "tar sand revenue"],
        "Coalbed Methane Revenue": ["coalbed methane", "cbm revenue"],
        "Extra Heavy Oil Revenue": ["extra heavy oil", "extra heavy oil revenue"],
        "Ultra Deepwater Revenue": ["ultra deepwater", "ultra deepwater revenue"],
        "Arctic Revenue": ["arctic", "arctic revenue"],
        "Unconventional Production Revenue": ["unconventional production", "unconventional production revenue"],
        # Level 2 (Pipeline/Capacity) Columns
        "GOGEL Tab": ["GOGEL Tab"],
        "Length of Pipelines under Development": ["length of pipelines", "pipeline under dev"],
        "Liquefaction Capacity (Export)": ["liquefaction capacity (export)", "lng export capacity", "Liquefaction Capacity Export"],
        "Regasification Capacity (Import)": ["regasification capacity (import)", "lng import capacity", "Regasification Capacity Import"],
        "Total Capacity under Development": ["total capacity under development", "total dev capacity"]
    }
    df = rename_columns(df, rename_map, how="partial")

    # --- 3) Ensure Key Columns Exist ---
    # For revenue thresholds (Level 1), we require these columns:
    revenue_cols = [
        "Hydrocarbons Production (%)",
        "Fracking Revenue",
        "Tar Sand Revenue",
        "Coalbed Methane Revenue",
        "Extra Heavy Oil Revenue",
        "Ultra Deepwater Revenue",
        "Arctic Revenue",
        "Unconventional Production Revenue"
    ]
    # For Level 2, we need these:
    level2_cols = [
        "GOGEL Tab",
        "Length of Pipelines under Development",
        "Liquefaction Capacity (Export)",
        "Regasification Capacity (Import)",
        "Total Capacity under Development"
    ]
    # Also ensure company info exists:
    info_cols = ["Company", "BB Ticker", "ISIN equity", "LEI"]

    # Fill missing columns with NaN (or 0 for numeric fields)
    for col in (info_cols + revenue_cols + level2_cols):
        if col not in df.columns:
            if col in info_cols:
                df[col] = ""
            else:
                df[col] = np.nan

    # Create an Exclusion Reason column (will combine Level 1 & 2 reasons)
    df["Exclusion Reason"] = ""

    # --- 4) Identify "No Data" Rows (for revenue data) ---
    companies_with_no_data = df[df[revenue_cols].isnull().all(axis=1)].copy()
    df = df.dropna(subset=revenue_cols, how="all")

    # --- 5) Process Revenue Data (Level 1) ---
    # Clean and convert revenue columns (assumed to be in decimals)
    for col in revenue_cols:
        df[col] = (
            df[col]
            .astype(str)
            .str.replace("%", "", regex=True)
            .str.replace(",", "", regex=True)
        )
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    
    # Calculate custom total thresholds if provided (for revenue)
    for key, threshold_data in total_thresholds.items():
        selected_sectors = threshold_data["sectors"]
        valid_sectors = [sector for sector in selected_sectors if sector in df.columns]
        if valid_sectors:
            df[key] = df[valid_sectors].sum(axis=1)
        else:
            df[key] = 0.0

    # Level 1 Exclusion Logic: iterate per row over sector_exclusions and custom revenue thresholds.
    level1_reasons = []
    for _, row in df.iterrows():
        reasons = []
        for sector, (exclude_flag, threshold_str) in sector_exclusions.items():
            if exclude_flag and threshold_str.strip():
                try:
                    user_threshold_decimal = float(threshold_str) / 100.0
                    if row[sector] > user_threshold_decimal:
                        reasons.append(
                            f"{sector} Revenue Exceeded: {row[sector]*100:.2f}% > {float(threshold_str):.2f}%"
                        )
                except ValueError:
                    pass
        for key, threshold_data in total_thresholds.items():
            try:
                threshold_value_decimal = float(threshold_data["threshold"]) / 100.0
                if row[key] > threshold_value_decimal:
                    reasons.append(
                        f"{key} Exceeded: {row[key]*100:.2f}% > {float(threshold_data['threshold']):.2f}%"
                    )
            except ValueError:
                pass
        level1_reasons.append(", ".join(reasons))
    df["Exclusion Reason Level 1"] = level1_reasons

    # --- 6) Process Pipeline/Capacity Data (Level 2) ---
    # Convert Level 2 numeric columns
    for col in level2_cols[1:]:
        df[col] = (
            df[col].astype(str)
            .str.replace("%", "", regex=True)
            .str.replace(",", "", regex=True)
        )
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    
    level2_reasons = []
    for _, row in df.iterrows():
        reasons = []
        # Upstream check based on GOGEL Tab (if user opted in)
        if level2_thresholds.get("apply_upstream", False):
            if isinstance(row["GOGEL Tab"], str) and "upstream" in row["GOGEL Tab"].lower():
                reasons.append("Upstream in GOGEL Tab")
        # Check each pipeline/capacity column against its threshold
        for col in level2_cols[1:]:
            thresh = level2_thresholds.get(col, 0)
            try:
                if row[col] > thresh:
                    reasons.append(f"{col} {row[col]:.2f} > {thresh:.2f}")
            except Exception:
                pass
        level2_reasons.append(", ".join(reasons))
    df["Exclusion Reason Level 2"] = level2_reasons

    # --- 7) Combine Exclusion Reasons ---
    # (If a company meets either Level 1 or Level 2 criteria, mark it as excluded.)
    combined_reasons = []
    for r1, r2 in zip(df["Exclusion Reason Level 1"], df["Exclusion Reason Level 2"]):
        combined = "; ".join([s for s in [r1, r2] if s])
        combined_reasons.append(combined)
    df["Exclusion Reason"] = combined_reasons

    # --- 8) Split Data into Excluded vs. Retained ---
    retained_companies = df[df["Exclusion Reason"] == ""].copy()
    excluded_companies = df[df["Exclusion Reason"] != ""].copy()
    # Make sure companies_with_no_data has all columns
    for col in df.columns:
        if col not in companies_with_no_data.columns:
            companies_with_no_data[col] = np.nan
    companies_with_no_data = companies_with_no_data[df.columns]

    # --- 9) Clean BB Ticker in the Output Only ---
    def remove_equity_from_bb_ticker(df):
        df = df.copy()
        if "BB Ticker" in df.columns:
            df["BB Ticker"] = (
                df["BB Ticker"]
                .astype(str)
                .str.replace(r"\u00A0", " ", regex=True)
                .str.replace(r"(?i)\bEquity\b", "", regex=True)
                .str.replace(r"\s+", " ", regex=True)
                .str.strip()
            )
        return df

    retained_clean = remove_equity_from_bb_ticker(retained_companies)
    excluded_clean = remove_equity_from_bb_ticker(excluded_companies)
    no_data_clean = remove_equity_from_bb_ticker(companies_with_no_data)

    # --- 10) Write to Excel in Memory ---
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        retained_clean.to_excel(writer, sheet_name="Retained Companies", index=False)
        excluded_clean.to_excel(writer, sheet_name="Excluded Companies", index=False)
        no_data_clean.to_excel(writer, sheet_name="No Data Companies", index=False)
        # Optionally, you could write a summary sheet
    output.seek(0)

    stats = {
        "Total Companies": len(df) + len(companies_with_no_data),
        "Retained Companies": len(retained_companies),
        "Excluded Companies": len(excluded_companies),
        "Companies with No Data": len(companies_with_no_data)
    }
    return output, stats

##############################################
# STREAMLIT APP
##############################################
def main():
    st.title("O&G Exclusion Filter (Combined Level 1 & Level 2 Thresholds)")
    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

    if not uploaded_file:
        st.warning("Please upload an Excel file.")
        return

    st.sidebar.header("Level 1 Exclusion Criteria (Revenue in %)")
    def sector_exclusion_input(sector_name):
        exclude = st.sidebar.checkbox(f"Exclude {sector_name}", value=False)
        threshold = ""
        if exclude:
            threshold = st.sidebar.text_input(f"{sector_name} Revenue Threshold (%)", "")
        return sector_name, (exclude, threshold)
    # Define revenue sectors (Level 1)
    sector_list = [
        "Hydrocarbons Production (%)",
        "Fracking Revenue (%)",
        "Tar Sand Revenue (%)",
        "Coalbed Methane Revenue (%)",
        "Extra Heavy Oil Revenue (%)",
        "Ultra Deepwater Revenue (%)",
        "Arctic Revenue (%)",
        "Unconventional Production Revenue (%)",
    ]
    sector_exclusions = dict(sector_exclusion_input(s) for s in sector_list)

    st.sidebar.header("Custom Total Revenue Thresholds (Level 1, in %)")
    total_thresholds = {}
    num_custom_thresholds = st.sidebar.number_input("Number of Custom Total Thresholds", min_value=1, max_value=5, value=1)
    for i in range(num_custom_thresholds):
        selected_sectors = st.sidebar.multiselect(f"Select Sectors for Custom Threshold {i+1}", sector_list, key=f"sectors_{i}")
        total_threshold = st.sidebar.text_input(f"Total Revenue Threshold {i+1} (%)", "", key=f"threshold_{i}")
        if selected_sectors and total_threshold:
            total_thresholds[f"Custom Total Revenue {i+1}"] = {"sectors": selected_sectors, "threshold": total_threshold}

    st.sidebar.header("Level 2 Exclusion Criteria (Pipeline/Capacity)")
    apply_upstream = st.sidebar.checkbox("Exclude companies with 'upstream' in GOGEL Tab", value=True)
    pipeline_threshold = st.sidebar.number_input("Length of Pipelines under Development Threshold", value=0.0)
    liq_threshold = st.sidebar.number_input("Liquefaction Capacity (Export) Threshold", value=0.0)
    rega_threshold = st.sidebar.number_input("Regasification Capacity (Import) Threshold", value=0.0)
    total_cap_threshold = st.sidebar.number_input("Total Capacity under Development Threshold", value=0.0)
    level2_thresholds = {
        "apply_upstream": apply_upstream,
        "Length of Pipelines under Development": pipeline_threshold,
        "Liquefaction Capacity (Export)": liq_threshold,
        "Regasification Capacity (Import)": rega_threshold,
        "Total Capacity under Development": total_cap_threshold
    }

    if st.sidebar.button("Run Combined Exclusion"):
        output_file, stats = filter_companies_combined(uploaded_file, sector_exclusions, total_thresholds, level2_thresholds)
        if output_file:
            st.success("Processing complete!")
            st.subheader("Processing Statistics")
            for key, value in stats.items():
                st.write(f"**{key}:** {value}")
            st.download_button(
                label="Download Filtered Excel",
                data=output_file,
                file_name="O&G_Companies_Combined_Exclusion.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
