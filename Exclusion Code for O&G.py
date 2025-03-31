import re
import pandas as pd
import numpy as np
import io
from io import BytesIO
import streamlit as st

# -------------------------- Utility Functions --------------------------
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

# -------------------------- Level 1 Exclusion Logic --------------------------
def filter_companies_by_revenue(uploaded_file, sector_exclusions, total_thresholds):
    """
    Reads the Excel file, applies revenue-based (Level 1) exclusion logic,
    and returns an Excel file (in memory) plus summary stats.
    Output file has three sheets: Retained, Excluded, and No Data Companies.
    """
    if uploaded_file is None:
        return None, None
    
    # 1) Read the Excel file
    xls = pd.ExcelFile(uploaded_file)
    df = xls.parse("All Companies", header=[3, 4])
    
    # 2) Flatten multi-level columns
    df.columns = [" ".join(map(str, col)).strip() for col in df.columns]

    # 3) "Equity" string remover in BB Ticker column    
    def remove_equity_from_bb_ticker(df):
        df = df.copy()
        if "BB Ticker" in df.columns:
            df["BB Ticker"] = (
                df["BB Ticker"]
                .astype(str)
                .str.replace(r"\u00A0", " ", regex=True)  # Replace non-breaking spaces
                .str.replace(r"(?i)\s*Equity\s*", "", regex=True)  # Remove 'Equity'
                .str.strip()
            )
        return df

    # 4) Dynamically rename columns if needed
    rename_map = {
        "Company": ["company name", "company"],
        "BB Ticker": ["bloomberg bb ticker", "bb ticker"],
        "ISIN equity": ["isin codes isin equity", "isin equity"],
        "LEI": ["lei lei", "lei", "legal entity identifier"],
        "Hydrocarbons Production (%)": ["hydrocarbons production", "hydrocarbons"],
        "Fracking Revenue": ["fracking", "fracking revenue"],
        "Tar Sand Revenue": ["tar sands", "tar sand revenue"],
        "Coalbed Methane Revenue": ["coalbed methane", "cbm revenue"],
        "Extra Heavy Oil Revenue": ["extra heavy oil", "extra heavy oil revenue"],
        "Ultra Deepwater Revenue": ["ultra deepwater", "ultra deepwater revenue"],
        "Arctic Revenue": ["arctic", "arctic revenue"],
        "Unconventional Production Revenue": ["unconventional production", "unconventional production revenue"]
    }
    df = rename_columns(df, rename_map, how="partial")

    # 5) Ensure we have all columns. If missing, fill with NaN
    needed_cols = list(rename_map.keys())
    for col in needed_cols:
        if col not in df.columns:
            df[col] = np.nan

    # Create an Exclusion Reason column
    df["Exclusion Reason"] = ""
    # Keep only relevant columns
    all_cols = needed_cols + ["Exclusion Reason"]
    df = df[all_cols]
    
    # 6) Identify "No Data" rows (all revenue columns are NaN or empty)
    revenue_cols = needed_cols[4:]  # columns from index 4 onward are revenue columns
    companies_with_no_data = df[df[revenue_cols].isnull().all(axis=1)].copy()
    df = df.dropna(subset=revenue_cols, how='all')

    # 7) Clean & convert revenue columns to numeric
    for col in revenue_cols:
        df[col] = (
            df[col]
            .astype(str)
            .str.replace("%", "", regex=True)
            .str.replace(",", "", regex=True)
        )
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # 8) Calculate any custom total thresholds (sum of selected revenue sectors)
    for key, threshold_data in total_thresholds.items():
        selected_sectors = threshold_data["sectors"]
        valid_sectors = [sector for sector in selected_sectors if sector in df.columns]
        if valid_sectors:
            df[key] = df[valid_sectors].sum(axis=1)
        else:
            df[key] = 0.0

    # ---------- 9) Apply Level 1 exclusion logic ----------
    excluded_reasons = []
    for _, row in df.iterrows():
        reasons = []
        # Process each revenue indicator from sector_exclusions.
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
        # Process custom total thresholds
        for key, threshold_data in total_thresholds.items():
            try:
                threshold_value_decimal = float(threshold_data["threshold"]) / 100.0
                if row[key] > threshold_value_decimal:
                    reasons.append(
                        f"{key} Exceeded: {row[key]*100:.2f}% > {float(threshold_data['threshold']):.2f}%"
                    )
            except ValueError:
                pass
        excluded_reasons.append(", ".join(reasons))
    df["Exclusion Reason"] = excluded_reasons

    # 10) Split data into Retained vs. Excluded
    retained_companies = df[df["Exclusion Reason"] == ""].copy()
    excluded_companies = df[df["Exclusion Reason"] != ""].copy()
    for col in df.columns:
        if col not in companies_with_no_data.columns:
            companies_with_no_data[col] = np.nan
    companies_with_no_data = companies_with_no_data[df.columns]

    # ---------- 11) Write output to Excel in memory (Level 1) ----------
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        retained_clean = remove_equity_from_bb_ticker(retained_companies)
        excluded_clean = remove_equity_from_bb_ticker(excluded_companies)
        no_data_clean = remove_equity_from_bb_ticker(companies_with_no_data)
    
        retained_clean.to_excel(writer, sheet_name="Retained Companies", index=False)
        excluded_clean.to_excel(writer, sheet_name="Excluded Companies", index=False)
        no_data_clean.to_excel(writer, sheet_name="No Data Companies", index=False)
        stats = {
            "Total Companies": len(df) + len(companies_with_no_data),
            "Retained Companies": len(retained_companies),
            "Excluded Companies": len(excluded_companies),
            "Companies with No Data": len(companies_with_no_data)
        }
    output.seek(0)
    return output, stats

# -------------------------- Level 2 Exclusion Logic --------------------------
def filter_companies_by_revenue_level2(uploaded_file, sector_exclusions, total_thresholds):
    """
    Reads the Excel file, applies revenue-based exclusion logic (Level 1) AND
    adds an extra Level 2 criterion: if the "GOGEL Tab" column contains "upstream"
    (case-insensitive), an exclusion reason "Upstream in GOGEL Tab" is added.
    The output file has the same structure as Level 1.
    """
    if uploaded_file is None:
        return None, None
    
    # 1) Read the Excel file
    xls = pd.ExcelFile(uploaded_file)
    df = xls.parse("All Companies", header=[3, 4])
    
    # 2) Flatten multi-level columns
    df.columns = [" ".join(map(str, col)).strip() for col in df.columns]

    # 3) "Equity" remover for BB Ticker (applied only in output)
    def remove_equity_from_bb_ticker(df):
        df = df.copy()
        if "BB Ticker" in df.columns:
            df["BB Ticker"] = (
                df["BB Ticker"]
                .astype(str)
                .str.replace(r"\u00A0", " ", regex=True)
                .str.replace(r"(?i)\s*Equity\s*", "", regex=True)
                .str.strip()
            )
        return df

    # 4) Dynamically rename columns (include "GOGEL Tab" for Level 2)
    rename_map = {
        "Company": ["company name", "company"],
        "GOGEL Tab": ["gogel tab", "gogel"],  # Level 2: include this column
        "BB Ticker": ["bloomberg bb ticker", "bb ticker"],
        "ISIN equity": ["isin codes isin equity", "isin equity"],
        "LEI": ["lei lei", "lei", "legal entity identifier"],
        "Hydrocarbons Production (%)": ["hydrocarbons production", "hydrocarbons"],
        "Fracking Revenue": ["fracking", "fracking revenue"],
        "Tar Sand Revenue": ["tar sands", "tar sand revenue"],
        "Coalbed Methane Revenue": ["coalbed methane", "cbm revenue"],
        "Extra Heavy Oil Revenue": ["extra heavy oil", "extra heavy oil revenue"],
        "Ultra Deepwater Revenue": ["ultra deepwater", "ultra deepwater revenue"],
        "Arctic Revenue": ["arctic", "arctic revenue"],
        "Unconventional Production Revenue": ["unconventional production", "unconventional production revenue"]
    }
    df = rename_columns(df, rename_map, how="partial")

    # 5) Ensure key columns exist (include "GOGEL Tab")
    needed_cols = list(rename_map.keys())
    for col in needed_cols:
        if col not in df.columns:
            df[col] = np.nan

    # Create an Exclusion Reason column
    df["Exclusion Reason"] = ""
    # Keep relevant columns (for Level 2 include "GOGEL Tab")
    all_cols = needed_cols + ["Exclusion Reason"]
    df = df[all_cols]
    
    # 6) Identify "No Data" rows (for revenue data)
    revenue_cols = [col for col in needed_cols if col not in ["Company", "GOGEL Tab", "BB Ticker", "ISIN equity", "LEI"]]
    companies_with_no_data = df[df[revenue_cols].isnull().all(axis=1)].copy()
    df = df.dropna(subset=revenue_cols, how='all')

    # 7) Clean & convert revenue columns to numeric
    for col in revenue_cols:
        df[col] = (
            df[col]
            .astype(str)
            .str.replace("%", "", regex=True)
            .str.replace(",", "", regex=True)
        )
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # 8) Calculate any custom total thresholds (for revenue)
    for key, threshold_data in total_thresholds.items():
        selected_sectors = threshold_data["sectors"]
        valid_sectors = [sector for sector in selected_sectors if sector in df.columns]
        if valid_sectors:
            df[key] = df[valid_sectors].sum(axis=1)
        else:
            df[key] = 0.0

    # ---------- 9) Apply Level 1 exclusion logic (as before) ----------
    excluded_reasons = []
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
        excluded_reasons.append(", ".join(reasons))
    df["Exclusion Reason"] = excluded_reasons

    # ---------- 10) Apply additional Level 2 criterion ----------
    # If "GOGEL Tab" contains "upstream" (case-insensitive), append an exclusion reason.
    if "GOGEL Tab" in df.columns:
        df["Exclusion Reason"] = df.apply(
            lambda row: (row["Exclusion Reason"] + "; Upstream in GOGEL Tab").strip("; ")
            if (isinstance(row["GOGEL Tab"], str) and "upstream" in row["GOGEL Tab"].lower())
            else row["Exclusion Reason"],
            axis=1
        )

    # 11) Split data into Retained vs. Excluded
    retained_companies = df[df["Exclusion Reason"] == ""].copy()
    excluded_companies = df[df["Exclusion Reason"] != ""].copy()
    for col in df.columns:
        if col not in companies_with_no_data.columns:
            companies_with_no_data[col] = np.nan
    companies_with_no_data = companies_with_no_data[df.columns]

    # ---------- 12) Write output to Excel in memory (Level 2) ----------
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        retained_clean = remove_equity_from_bb_ticker(retained_companies)
        excluded_clean = remove_equity_from_bb_ticker(excluded_companies)
        no_data_clean = remove_equity_from_bb_ticker(companies_with_no_data)
    
        retained_clean.to_excel(writer, sheet_name="Retained Companies", index=False)
        excluded_clean.to_excel(writer, sheet_name="Excluded Companies", index=False)
        no_data_clean.to_excel(writer, sheet_name="No Data Companies", index=False)
        stats = {
            "Total Companies": len(df) + len(companies_with_no_data),
            "Retained Companies": len(retained_companies),
            "Excluded Companies": len(excluded_companies),
            "Companies with No Data": len(companies_with_no_data)
        }
    output.seek(0)
    return output, stats

# -------------------------- STREAMLIT APP --------------------------
def main():
    st.title("O&G Exclusion Filter")
    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

    st.sidebar.header("Set Exclusion Criteria (Thresholds in %)")
    def sector_exclusion_input(sector_name):
        """
        Returns a tuple: (sector_name, (exclude_checkbox, threshold_string))
        Example: ("Fracking Revenue", (True, "10")) => exclude if > 10%
        """
        exclude = st.sidebar.checkbox(f"Exclude {sector_name}", value=False)
        threshold = ""
        if exclude:
            threshold = st.sidebar.text_input(f"{sector_name} Revenue Threshold (%)", "")
        return sector_name, (exclude, threshold)

    # Define the sectors for Level 1 user input
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

    st.sidebar.header("Set Multiple Custom Total Revenue Thresholds (in %)")
    total_thresholds = {}
    num_custom_thresholds = st.sidebar.number_input("Number of Custom Total Thresholds", min_value=1, max_value=5, value=1)
    for i in range(num_custom_thresholds):
        selected_sectors = st.sidebar.multiselect(f"Select Sectors for Custom Threshold {i+1}", sector_list, key=f"sectors_{i}")
        total_threshold = st.sidebar.text_input(f"Total Revenue Threshold {i+1} (%)", "", key=f"threshold_{i}")
        if selected_sectors and total_threshold:
            total_thresholds[f"Custom Total Revenue {i+1}"] = {"sectors": selected_sectors, "threshold": total_threshold}

    st.sidebar.header("Choose Exclusion Level")
    level = st.sidebar.radio("Select Level", options=["Level 1", "Level 2"])

    if st.sidebar.button("Run Exclusion"):
        if uploaded_file:
            if level == "Level 1":
                output_file, stats = filter_companies_by_revenue(uploaded_file, sector_exclusions, total_thresholds)
                filename = "O&G_Companies_Level1_Exclusion.xlsx"
            else:  # Level 2
                output_file, stats = filter_companies_by_revenue_level2(uploaded_file, sector_exclusions, total_thresholds)
                filename = "O&G_Companies_Level2_Exclusion.xlsx"
            if output_file:
                st.success("File processed successfully!")
                st.subheader("Processing Statistics")
                for key, value in stats.items():
                    st.write(f"**{key}:** {value}")
                st.download_button(
                    label="Download Filtered Excel",
                    data=output_file,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.warning("Please upload an Excel file first.")

if __name__ == "__main__":
    main()
