import re
import pandas as pd

def find_column(df, possible_matches, how="exact", required=True):
    """
    Searches df.columns for the first column that matches any of the possible_matches.
    
    Parameters
    ----------
    df : pd.DataFrame
        The DataFrame in which to search for columns.
    possible_matches : list of str
        A list of potential column names or patterns to look for.
    how : str, optional
        Matching mode:
         - "exact"  => requires exact match
         - "partial" => checks if `possible_match` is a substring of the column name
         - "regex"   => interprets `possible_match` as a regex
    required : bool, optional
        If True, raises an error if no column is found; otherwise returns None.

    Returns
    -------
    str or None
        The actual column name in df.columns that was matched, or None if not found
        (and required=False).
    """
    df_cols = list(df.columns)
    for col in df_cols:
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

    Parameters
    ----------
    df : pd.DataFrame
        The DataFrame whose columns will be renamed in-place.
    rename_map : dict
        Keys = new/standardized column name,
        Values = list of possible matches for that column in the DF.
    how : str
        "exact", "partial", or "regex".

    Returns
    -------
    df : pd.DataFrame
        Same DataFrame with renamed columns.
    """
    for new_col_name, patterns in rename_map.items():
        old_name = find_column(df, patterns, how=how, required=False)
        if old_name:
            df.rename(columns={old_name: new_col_name}, inplace=True)
    return df

import streamlit as st
import pandas as pd
import numpy as np
import io
from io import BytesIO

# If you placed the utilities in the same file, just use them directly
# Otherwise, uncomment and import:
# from column_utils import rename_columns

def filter_companies_by_revenue(uploaded_file, sector_exclusions, total_thresholds):
    if uploaded_file is None:
        return None, None
    
    # ---------- 1) Read the Excel file ----------
    xls = pd.ExcelFile(uploaded_file)
    # Adjust headers if your real file differs
    df = xls.parse("All Companies", header=[3, 4])
    
    # Flatten multi-level columns
    df.columns = [' '.join(map(str, col)).strip() for col in df.columns]
    
    # ---------- 2) Dynamically rename columns ----------
    rename_map = {
        "Company":                 ["company name", "company"],
        "BB Ticker":               ["bloomberg bb ticker", "bb ticker"],
        "ISIN equity":             ["isin codes isin equity", "isin equity"],
        "LEI":                     ["lei lei", "lei", "legal entity identifier"],
        "Fracking Revenue":        ["fracking", "fracking revenue"],
        "Tar Sand Revenue":        ["tar sands", "tar sand revenue"],
        "Coalbed Methane Revenue": ["coalbed methane", "cbm revenue"],
        "Extra Heavy Oil Revenue": ["extra heavy oil", "extra heavy oil revenue"],
        "Ultra Deepwater Revenue": ["ultra deepwater", "ultra deepwater revenue"],
        "Arctic Revenue":          ["arctic", "arctic revenue"],
        "Unconventional Production Revenue": ["unconventional production", "unconventional production revenue"]
    }
    # ---- BEGIN EXTRA LINE #1: Add Hydrocarbons Production to rename_map
    rename_map["Hydrocarbons Production"] = ["hydrocarbons production", "hydrocarbons"]
    # ---- END EXTRA LINE #1

    df = rename_columns(df, rename_map, how="partial")

    # Ensure we have all columns. If missing, fill with NaN
    needed_cols = list(rename_map.keys())
    for col in needed_cols:
        if col not in df.columns:
            df[col] = np.nan

    # Create an Exclusion Reason column
    df["Exclusion Reason"] = ""
    
    # Keep only relevant columns
    all_cols = needed_cols + ["Exclusion Reason"]
    df = df[all_cols]
    
    # ---------- 3) Identify "No Data" rows ----------
    revenue_cols = needed_cols[4:]  # everything after the first 4 is "revenue"
    companies_with_no_data = df[df[revenue_cols].isnull().all(axis=1)].copy()
    
    # Drop rows that have all null revenues
    df = df.dropna(subset=revenue_cols, how='all')

    # ---------- 4) Clean & convert to numeric ----------
    for col in revenue_cols:
        df[col] = (
            df[col]
            .astype(str)
            .str.replace('%', '', regex=True)
            .str.replace(',', '', regex=True)
        )
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # If values are 0 <= x <= 1, multiply by 100
    if df[revenue_cols].max().max() <= 1:
        df[revenue_cols] = df[revenue_cols] * 100

    # ---------- 5) Calculate total thresholds (optional) ----------
    for key, threshold_data in total_thresholds.items():
        selected_sectors = threshold_data["sectors"]
        threshold_value = threshold_data["threshold"]
        valid_sectors = [sector for sector in selected_sectors if sector in df.columns]
        if valid_sectors:
            df[key] = df[valid_sectors].sum(axis=1)

    # ---------- 6) Apply exclusion logic ----------
    excluded_reasons = []
    for _, row in df.iterrows():
        reasons = []
        # sector_exclusions is a dict like:
        # { "Fracking Revenue": (True, "10"), "Arctic Revenue": (False, ""), ... }
        for sector, (exclude_flag, threshold_str) in sector_exclusions.items():
            if exclude_flag:
                th = float(threshold_str) if threshold_str else 0.0
                if row[sector] > th:
                    reasons.append(f"{sector} Revenue Exceeded")

        # Check each custom total threshold
        for key, threshold_data in total_thresholds.items():
            threshold_value = float(threshold_data["threshold"])
            if key in df.columns and row[key] > threshold_value:
                reasons.append(f"{key} Revenue Exceeded")

        excluded_reasons.append(", ".join(reasons))

    df["Exclusion Reason"] = excluded_reasons

    # ---------- 7) Split data into retained vs excluded ----------
    retained_companies = df[df["Exclusion Reason"] == ""].copy()
    excluded_companies = df[df["Exclusion Reason"] != ""].copy()

    # Make sure companies_with_no_data has all columns
    for col in df.columns:
        if col not in companies_with_no_data.columns:
            companies_with_no_data[col] = np.nan

    # Reorder them the same way
    companies_with_no_data = companies_with_no_data[df.columns]

    # ---------- 8) Write output to Excel in memory ----------
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        retained_companies.to_excel(writer, sheet_name="Retained Companies", index=False)
        excluded_companies.to_excel(writer, sheet_name="Excluded Companies", index=False)
        companies_with_no_data.to_excel(writer, sheet_name="No Data Companies", index=False)
    output.seek(0)
    
    stats = {
        "Total Companies": len(df) + len(companies_with_no_data),
        "Retained Companies": len(retained_companies),
        "Excluded Companies": len(excluded_companies),
        "Companies with No Data": len(companies_with_no_data)
    }
    return output, stats

# -------------------------- STREAMLIT APP --------------------------
def main():
    st.title("Level 1 Exclusion Filter (All Companies - Dynamic Column Search)")
    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

    st.sidebar.header("Set Exclusion Criteria")

    def sector_exclusion_input(sector_name):
        exclude = st.sidebar.checkbox(f"Exclude {sector_name}", value=False)
        threshold = ""
        if exclude:
            threshold = st.sidebar.text_input(f"{sector_name} Revenue Threshold (%)", "")
        return sector_name, (exclude, threshold)

    # For example, the 7 unconventionals:
    sector_exclusions = dict([
        sector_exclusion_input("Fracking Revenue"),
        sector_exclusion_input("Tar Sand Revenue"),
        sector_exclusion_input("Coalbed Methane Revenue"),
        sector_exclusion_input("Extra Heavy Oil Revenue"),
        sector_exclusion_input("Ultra Deepwater Revenue"),
        sector_exclusion_input("Arctic Revenue"),
        sector_exclusion_input("Unconventional Production Revenue"),
        # ---- BEGIN EXTRA LINE #2: Also allow “Hydrocarbons Production” to be excluded
        sector_exclusion_input("Hydrocarbons Production")
        # ---- END EXTRA LINE #2
    ])

    st.sidebar.header("Set Multiple Custom Total Revenue Thresholds")
    total_thresholds = {}
    num_custom_thresholds = st.sidebar.number_input(
        "Number of Custom Total Thresholds",
        min_value=1, max_value=5, value=1
    )
    for i in range(num_custom_thresholds):
        selected_sectors = st.sidebar.multiselect(
            f"Select Sectors for Custom Threshold {i+1}",
            list(sector_exclusions.keys()),
            key=f"sectors_{i}"
        )
        total_threshold = st.sidebar.text_input(
            f"Total Revenue Threshold {i+1} (%)",
            "",
            key=f"threshold_{i}"
        )
        if selected_sectors and total_threshold:
            total_thresholds[f"Custom Total Revenue {i+1}"] = {
                "sectors": selected_sectors,
                "threshold": total_threshold
            }

    if st.sidebar.button("Run Level 1 Exclusion"):
        if uploaded_file:
            output_file, stats = filter_companies_by_revenue(
                uploaded_file, sector_exclusions, total_thresholds
            )
            if output_file:
                st.success("File processed successfully!")
                st.subheader("Processing Statistics")
                for key, value in stats.items():
                    st.write(f"**{key}:** {value}")

                st.download_button(
                    label="Download Filtered Excel",
                    data=output_file,
                    file_name="O&G Companies Level 1 Exclusion.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main()

