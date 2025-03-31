import re
import pandas as pd
import numpy as np
import io
from io import BytesIO
import streamlit as st

#############################
# LEVEL 1 FUNCTIONS (O&G Revenue Exclusion)
#############################

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

def filter_companies_by_revenue(uploaded_file, sector_exclusions, total_thresholds):
    """
    Reads the Excel file, applies revenue-based exclusion logic, and returns an
    Excel file (in memory) plus summary stats.
    """
    if uploaded_file is None:
        return None, None

    # 1) Read the Excel file
    xls = pd.ExcelFile(uploaded_file)
    # Adjust headers if your real file differs
    df = xls.parse("All Companies", header=[3, 4])
    
    # 2) Flatten multi-level columns
    df.columns = [" ".join(map(str, col)).strip() for col in df.columns]

    # 3) "Equity" string remover in BB Ticker column (applied only in the output)
    def remove_equity_from_bb_ticker(df):
        df = df.copy()
        if "BB Ticker" in df.columns:
            df["BB Ticker"] = (
                df["BB Ticker"]
                .astype(str)
                .str.replace(r"\u00A0", " ", regex=True)  # Replace non-breaking spaces
                .str.replace(r"(?i)\s*Equity\s*", "", regex=True)  # Remove 'Equity' with surrounding spaces
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
    revenue_cols = needed_cols[4:]  # everything after the first 4 is "revenue" columns
    companies_with_no_data = df[df[revenue_cols].isnull().all(axis=1)].copy()
    
    # Drop rows that have all-null revenues
    df = df.dropna(subset=revenue_cols, how='all')

    # 7) Clean & convert to numeric
    for col in revenue_cols:
        df[col] = (
            df[col]
            .astype(str)
            .str.replace("%", "", regex=True)
            .str.replace(",", "", regex=True)
        )
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # 8) Calculate any custom total thresholds (sums of certain sectors)
    for key, threshold_data in total_thresholds.items():
        selected_sectors = threshold_data["sectors"]
        valid_sectors = [sector for sector in selected_sectors if sector in df.columns]
        if valid_sectors:
            df[key] = df[valid_sectors].sum(axis=1)
        else:
            df[key] = 0.0

    # ---------- 9) Apply exclusion logic ----------
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

    # 10) Split data into retained vs excluded
    retained_companies = df[df["Exclusion Reason"] == ""].copy()
    excluded_companies = df[df["Exclusion Reason"] != ""].copy()
    for col in df.columns:
        if col not in companies_with_no_data.columns:
            companies_with_no_data[col] = np.nan
    companies_with_no_data = companies_with_no_data[df.columns]

    # ---------- 11) Write output to Excel in memory ----------
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

#############################
# LEVEL 2 FUNCTIONS (Pipeline/Capacity Exclusion)
#############################

def flatten_multilevel_columns_lvl2(df):
    """Flatten multi-level column headers into single strings."""
    df.columns = [
        " ".join(str(level) for level in col).strip()
        for col in df.columns
    ]
    return df

def find_column_lvl2(df, possible_matches, required=True):
    """Finds the first column matching any item in possible_matches."""
    for col in df.columns:
        col_lower = col.strip().lower().replace("\n", " ")
        for pattern in possible_matches:
            pat_lower = pattern.strip().lower().replace("\n", " ")
            if pat_lower in col_lower:
                return col
    if required:
        raise ValueError(
            f"Could not find a required column among {possible_matches}\n"
            f"Available columns: {list(df.columns)}"
        )
    return None

def rename_columns_lvl2(df):
    """
    Flatten multi-level headers and ensure correct column detection for Level 2.
    """
    df = flatten_multilevel_columns_lvl2(df)
    # Shift header row up by 1 row
    df = df.iloc[1:].reset_index(drop=True)
    rename_map = {
        "Company": ["company"],  
        "GOGEL Tab": ["GOGEL Tab"],  
        "BB Ticker": ["bb ticker", "bloomberg ticker"],
        "ISIN Equity": ["isin equity", "isin code"],
        "LEI": ["lei"],
        "Length of Pipelines under Development": ["length of pipelines", "pipeline under dev"],
        "Liquefaction Capacity (Export)": ["liquefaction capacity (export)", "lng export capacity", "Liquefaction Capacity Export"],
        "Regasification Capacity (Import)": ["regasification capacity (import)", "lng import capacity", "Regasification Capacity Import"],
        "Total Capacity under Development": ["total capacity under development", "total dev capacity"]
    }
    for new_col, patterns in rename_map.items():
        old_col = find_column_lvl2(df, patterns, required=False)
        if old_col and old_col != new_col:
            df.rename(columns={old_col: new_col}, inplace=True)
    return df

def filter_all_companies_lvl2(df):
    """Parses 'All Companies' sheet, applies exclusion logic, and splits into categories for Level 2."""
    df = rename_columns_lvl2(df)
    required_columns = [
        "Company", "GOGEL Tab", "BB Ticker", "ISIN Equity", "LEI",
        "Length of Pipelines under Development",
        "Liquefaction Capacity (Export)",
        "Regasification Capacity (Import)",
        "Total Capacity under Development"
    ]
    for col in required_columns:
        if col not in df.columns:
            df[col] = None if col in ["Company", "GOGEL Tab", "BB Ticker", "ISIN Equity", "LEI"] else 0

    numeric_cols = [
        "Length of Pipelines under Development",
        "Liquefaction Capacity (Export)",
        "Regasification Capacity (Import)",
        "Total Capacity under Development"
    ]
    for c in numeric_cols:
        df[c] = (
            df[c].astype(str)
            .str.replace("%", "", regex=True)
            .str.replace(",", "", regex=True)
        )
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    df["Upstream_Exclusion_Flag"] = df["GOGEL Tab"].str.contains("upstream", case=False, na=False)
    df["Midstream_Exclusion_Flag"] = (
        (df["Length of Pipelines under Development"] > 0) |
        (df["Liquefaction Capacity (Export)"] > 0) |
        (df["Regasification Capacity (Import)"] > 0) |
        (df["Total Capacity under Development"] > 0)
    )
    df["Excluded"] = df["Upstream_Exclusion_Flag"] | df["Midstream_Exclusion_Flag"]

    def get_exclusion_reason(row):
        reasons = []
        if row["Upstream_Exclusion_Flag"]:
            reasons.append("Upstream in GOGEL Tab")
        if row["Midstream_Exclusion_Flag"]:
            reasons.append("Midstream Expansion > 0")
        return "; ".join(reasons)
    
    df["Exclusion Reason"] = df.apply(get_exclusion_reason, axis=1)

    retained_df = df[~df["Excluded"]].copy()
    excluded_df = df[df["Excluded"]].copy()
    final_cols = [
        "Company", "BB Ticker", "ISIN Equity", "LEI",
        "GOGEL Tab",
        "Length of Pipelines under Development",
        "Liquefaction Capacity (Export)",
        "Regasification Capacity (Import)",
        "Total Capacity under Development",
        "Exclusion Reason"
    ]
    for c in final_cols:
        for d in [excluded_df, retained_df]:
            if c not in d.columns:
                d[c] = None
    return excluded_df[final_cols], retained_df[final_cols]

#############################
# STREAMLIT APP: SELECT LEVEL
#############################

def main():
    st.title("O&G Exclusion Filter")
    level = st.sidebar.radio("Select Exclusion Level", options=["Level 1", "Level 2"])

    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

    if not uploaded_file:
        st.warning("Please upload an Excel file.")
        return

    # LEVEL 1 UI and Processing
    if level == "Level 1":
        st.subheader("Level 1 Exclusion Filter for O&G")
        st.sidebar.header("Set Exclusion Criteria (Thresholds in %)")
        
        def sector_exclusion_input(sector_name):
            exclude = st.sidebar.checkbox(f"Exclude {sector_name}", value=False)
            threshold = ""
            if exclude:
                threshold = st.sidebar.text_input(f"{sector_name} Revenue Threshold (%)", "")
            return sector_name, (exclude, threshold)
        
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
        
        output_file, stats = filter_companies_by_revenue(uploaded_file, sector_exclusions, total_thresholds)
        if output_file:
            st.success("Level 1 processing complete!")
            st.subheader("Processing Statistics")
            for key, value in stats.items():
                st.write(f"**{key}:** {value}")
            st.download_button(
                label="Download Filtered Excel",
                data=output_file,
                file_name="O&G_Companies_Level1_Exclusion.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    # LEVEL 2 UI and Processing
    elif level == "Level 2":
        st.subheader("Level 2 Exclusion Filter for O&G")
        xls = pd.ExcelFile(uploaded_file)
        if "All Companies" not in xls.sheet_names:
            st.error("No sheet named 'All Companies'.")
            return
        df_all = pd.read_excel(uploaded_file, sheet_name="All Companies", header=[3,4])
        excluded, retained = filter_all_companies_lvl2(df_all)
        
        # Remove "Equity" from BB Ticker for output
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
        excluded = remove_equity_from_bb_ticker(excluded)
        retained = remove_equity_from_bb_ticker(retained)
        
        total_companies = len(excluded) + len(retained)
        st.subheader("Summary Statistics")
        st.write(f"**Total Companies Processed:** {total_companies}")
        st.write(f"**Excluded Companies (Upstream & Midstream):** {len(excluded)}")
        st.write(f"**Retained Companies:** {len(retained)}")
        
        st.subheader("Excluded Companies")
        st.dataframe(excluded)
        st.subheader("Retained Companies")
        st.dataframe(retained)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            excluded.to_excel(writer, sheet_name="Excluded", index=False)
            retained.to_excel(writer, sheet_name="Retained", index=False)
        output.seek(0)
        st.download_button(
            "Download Processed File",
            output,
            "O&G_Companies_Level2_Exclusion.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
