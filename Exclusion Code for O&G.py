import re
import pandas as pd
import numpy as np
from io import BytesIO

def find_column(df, possible_matches, how="exact", required=True):
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
    for new_col_name, patterns in rename_map.items():
        old_name = find_column(df, patterns, how=how, required=False)
        if old_name:
            df.rename(columns={old_name: new_col_name}, inplace=True)
    return df


def filter_companies_by_revenue(uploaded_file, sector_exclusions, total_thresholds):
    if uploaded_file is None:
        return None, None
    
    # ---------- 1) Read the Excel file ----------
    xls = pd.ExcelFile(uploaded_file)
    df = xls.parse("All Companies", header=[3, 4])
    
    # Flatten multi-level columns
    df.columns = [' '.join(map(str, col)).strip() for col in df.columns]
    
    # ---------- 2) Dynamically rename columns ----------
    rename_map = {
        "Company":                 ["company name", "company"],
        "BB Ticker":               ["bloomberg bb ticker", "bb ticker"],
        "ISIN equity":             ["isin codes isin equity", "isin equity"],
        "LEI":                     ["lei lei", "lei", "legal entity identifier"],
        "Hydrocarbons Production": ["hydrocarbons production", "hydrocarbons"],
        "Fracking Revenue":        ["fracking", "fracking revenue"],
        "Tar Sand Revenue":        ["tar sands", "tar sand revenue"],
        "Coalbed Methane Revenue": ["coalbed methane", "cbm revenue"],
        "Extra Heavy Oil Revenue": ["extra heavy oil", "extra heavy oil revenue"],
        "Ultra Deepwater Revenue": ["ultra deepwater", "ultra deepwater revenue"],
        "Arctic Revenue":          ["arctic", "arctic revenue"],
        "Unconventional Production Revenue": ["unconventional production", "unconventional production revenue"]
    }
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
        df[col] = (df[col]
                   .astype(str)
                   .str.replace('%', '', regex=True)
                   .str.replace(',', '', regex=True))
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # If columns appear to be in [0,1], convert them to [0,100].
    # This turns 0.20 => 20.0 for example.
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
        
        # For each sector, compare row[sector] against the user threshold directly
        for sector, (exclude_flag, threshold_str) in sector_exclusions.items():
            if exclude_flag:
                if not threshold_str.strip():
                    continue  # no threshold typed
                th = float(threshold_str)  # interpret “10” as 10%  
                if row[sector] > th:
                    reasons.append(f"{sector} Revenue Exceeded ({row[sector]:.1f} > {th})")

        # Check each custom total threshold
        for key, threshold_data in total_thresholds.items():
            if key not in df.columns:
                continue
            if not threshold_data["threshold"].strip():
                continue
            threshold_value = float(threshold_data["threshold"])
            if row[key] > threshold_value:
                reasons.append(f"{key} Revenue Exceeded ({row[key]:.1f} > {threshold_value})")

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
