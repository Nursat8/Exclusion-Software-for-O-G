import streamlit as st
import pandas as pd
import io
from io import BytesIO

##################################
# 1) HELPER FUNCTIONS
##################################

def find_column(df, possible_matches, how="partial", required=True):
    """
    Searches df.columns for the first column that matches any item in `possible_matches`.
    Returns the actual column name, or None if not found and required=False.
    """
    for col in df.columns:
        for pattern in possible_matches:
            col_lower = col.strip().lower()
            pat_lower = pattern.strip().lower()
            if how == "exact":
                if col_lower == pat_lower:
                    return col
            elif how == "partial":
                if pat_lower in col_lower:
                    return col

    if required:
        raise ValueError(
            f"Could not find a required column among: {possible_matches}\n"
            f"Available columns: {list(df.columns)}"
        )
    return None

def rename_columns(df, rename_map, how="partial"):
    """
    Rename columns in-place according to rename_map:
        { new_col_name: [list_of_possible_matches], ... }
    """
    for new_col_name, possible_names in rename_map.items():
        old_name = find_column(df, possible_names, how=how, required=False)
        if old_name:
            df.rename(columns={old_name: new_col_name}, inplace=True)
    return df

def load_data(file, sheet_name, header_row=0):
    """Helper to load a given sheet from Excel with a specified header row."""
    return pd.read_excel(file, sheet_name=sheet_name, header=header_row)

##################################
# 2) CORE LOGIC
##################################

def filter_exclusions_and_retained(upstream_df, midstream_df):
    """
    1) Dynamically rename columns in Upstream for 'Fossil Fuel Share of Revenue'
    2) Dynamically rename columns in Midstream for pipeline/capacity
    3) Exclude if Upstream share > 0 or any Midstream capacity > 0
    4) Drop or reclassify blank 'Company' rows so they don't appear in 'Excluded'
    5) Split final into (Excluded, Retained, No Data)
    """

    # ------------------ (A) RENAME & PREP UPSTREAM ------------------
    upstream_rename_map = {
        "Company": ["company", "company name"],
        "Fossil Fuel Share of Revenue": [
            "fossil fuel share of revenue",
            "fossil fuel share",
            "fossil fuel revenue"
        ],
        "BB Ticker": ["bb ticker", "bloomberg ticker"],
        "ISIN Equity": ["isin equity", "isin code"],
        "LEI": ["lei"]
    }
    rename_columns(upstream_df, upstream_rename_map, how="partial")

    # Ensure required columns exist
    for col in upstream_rename_map.keys():
        if col not in upstream_df.columns:
            upstream_df[col] = None

    # Subset relevant
    upstream_subset = upstream_df[
        ["Company", "Fossil Fuel Share of Revenue", "BB Ticker", "ISIN Equity", "LEI"]
    ].copy()

    # Clean & convert fossil-fuel share
    upstream_subset["Fossil Fuel Share of Revenue"] = (
        upstream_subset["Fossil Fuel Share of Revenue"]
        .astype(str)
        .str.replace("%", "", regex=True)
    )
    upstream_subset["Fossil Fuel Share of Revenue"] = pd.to_numeric(
        upstream_subset["Fossil Fuel Share of Revenue"], errors="coerce"
    ).fillna(0)

    # Upstream exclusion if share > 0
    upstream_subset["Upstream_Exclusion_Flag"] = (
        upstream_subset["Fossil Fuel Share of Revenue"] > 0
    )

    # ------------------ (B) RENAME & PREP MIDSTREAM ------------------
    midstream_rename_map = {
        "Company": ["company", "company name"],
        "Length of Pipelines under Development": ["length of pipelines", "pipeline under dev"],
        "Liquefaction Capacity (Export)": ["liquefaction capacity", "lng export capacity"],
        "Regasification Capacity (Import)": ["regasification capacity", "lng import capacity"],
        "Total Capacity under Development": ["total capacity under development", "total dev capacity"]
    }
    rename_columns(midstream_df, midstream_rename_map, how="partial")

    # Fill missing numeric columns with 0
    for col in midstream_rename_map.keys():
        if col not in midstream_df.columns:
            midstream_df[col] = 0

    midstream_subset = midstream_df[
        [
            "Company",
            "Length of Pipelines under Development",
            "Liquefaction Capacity (Export)",
            "Regasification Capacity (Import)",
            "Total Capacity under Development"
        ]
    ].copy()

    # Convert to numeric
    numeric_cols = [
        "Length of Pipelines under Development",
        "Liquefaction Capacity (Export)",
        "Regasification Capacity (Import)",
        "Total Capacity under Development"
    ]
    for col in numeric_cols:
        midstream_subset[col] = pd.to_numeric(midstream_subset[col], errors='coerce').fillna(0)

    # Aggregate by company
    midstream_grouped = (
        midstream_subset
        .groupby("Company", dropna=False)
        .agg({
            "Length of Pipelines under Development": "max",
            "Liquefaction Capacity (Export)": "max",
            "Regasification Capacity (Import)": "max",
            "Total Capacity under Development": "max"
        })
        .reset_index()
    )

    # Midstream exclusion if any capacity > 0
    midstream_grouped["Midstream_Exclusion_Flag"] = (
        (midstream_grouped["Length of Pipelines under Development"] > 0)
        | (midstream_grouped["Liquefaction Capacity (Export)"] > 0)
        | (midstream_grouped["Regasification Capacity (Import)"] > 0)
        | (midstream_grouped["Total Capacity under Development"] > 0)
    )

    # ------------------ (C) COMBINE & CLEAN ------------------
    combined = pd.merge(
        upstream_subset,
        midstream_grouped,
        on="Company",
        how="outer"
    )

    # A) Convert Company to string for consistency
    combined["Company"] = combined["Company"].astype(str).str.strip()

    # B) Define your "junk" or "nonsense" set for Company
    junk_values = {"", ".", "n.a.", "na", "0"}

    def is_nonsense_company(val: str) -> bool:
        """Return True if 'val' is in junk_values or looks obviously junky."""
        return val.lower() in junk_values

    # C) Filter out nonsense
    combined = combined[~combined["Company"].apply(is_nonsense_company)].copy()

    # D) Now proceed with Upstream_Exclusion_Flag, Midstream_Exclusion_Flag, etc.
    combined["Upstream_Exclusion_Flag"] = combined["Upstream_Exclusion_Flag"].fillna(False).astype(bool)
    combined["Midstream_Exclusion_Flag"] = combined["Midstream_Exclusion_Flag"].fillna(False).astype(bool)
    combined["Excluded"] = combined["Upstream_Exclusion_Flag"] | combined["Midstream_Exclusion_Flag"]

    # Build reason
    reasons = []
    for _, row in combined.iterrows():
        r = []
        if row["Upstream_Exclusion_Flag"]:
            r.append("Upstream - Fossil Fuel Share > 0%")
        if row["Midstream_Exclusion_Flag"]:
            r.append("Midstream Expansion > 0")
        reasons.append("; ".join(r))
    combined["Exclusion Reason"] = reasons

    # ------------------ (D) SPLIT INTO EXCLUDED / RETAINED / NO DATA ------------------
    def is_empty_string_or_nan(val):
        return pd.isna(val) or str(val).strip() == ""

    no_data_cond = (
        (~combined["Excluded"])
        & combined["BB Ticker"].apply(is_empty_string_or_nan)
        & combined["ISIN Equity"].apply(is_empty_string_or_nan)
        & combined["LEI"].apply(is_empty_string_or_nan)
        & (combined.get("Length of Pipelines under Development", 0) == 0)
        & (combined.get("Liquefaction Capacity (Export)", 0) == 0)
        & (combined.get("Regasification Capacity (Import)", 0) == 0)
        & (combined.get("Total Capacity under Development", 0) == 0)
    )

    no_data_companies = combined[no_data_cond].copy()
    excluded_companies = combined[combined["Excluded"]].copy()
    retained_companies = combined[~combined["Excluded"] & ~no_data_cond].copy()

    return excluded_companies, retained_companies, no_data_companies

##################################
# 3) STREAMLIT APP
##################################

def main():
    st.title("Level 2 Exclusion Filter (Upstream & Midstream) - Blank Row Fix")
    uploaded_file = st.file_uploader("Upload the Excel file", type=["xlsx"])

    if uploaded_file:
        # Check that the needed sheets exist
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names

        if "Upstream" not in sheet_names:
            st.error(f"Could not find a sheet named 'Upstream'. Found sheets: {sheet_names}")
            return
        if "Midstream Expansion" not in sheet_names:
            st.error(f"Could not find a sheet named 'Midstream Expansion'. Found sheets: {sheet_names}")
            return

        # Load them. Adjust header_row if needed
        upstream_df = load_data(uploaded_file, sheet_name="Upstream", header_row=4)
        midstream_df = load_data(uploaded_file, sheet_name="Midstream Expansion", header_row=4)

        # Run core logic
        excluded_data, retained_data, no_data_data = filter_exclusions_and_retained(
            upstream_df, midstream_df
        )

        # --- Stats ---
        excluded_count = len(excluded_data)
        retained_count = len(retained_data)
        no_data_count = len(no_data_data)
        total_count = excluded_count + retained_count + no_data_count

        st.markdown("### Statistics")
        st.write(f"**Total companies:** {total_count}")
        st.write(f"**Excluded:** {excluded_count}")
        st.write(f"**Retained:** {retained_count}")
        st.write(f"**No Data:** {no_data_count}")

        # Show dataframes
        st.subheader("Excluded Companies")
        st.dataframe(excluded_data)

        st.subheader("Retained Companies")
        st.dataframe(retained_data)

        st.subheader("No Data Companies")
        st.dataframe(no_data_data)

        # Export to Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            excluded_data.to_excel(writer, index=False, sheet_name='Exclusions')
            retained_data.to_excel(writer, index=False, sheet_name='Retained')
            no_data_data.to_excel(writer, index=False, sheet_name='No Data')
        output.seek(0)

        st.download_button(
            "Download Exclusion & Retention & NoData List",
            output,
            "O&G_companies_Level_2_Exclusion_Fixed.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
