import streamlit as st
import pandas as pd
import io
from io import BytesIO

def find_column(df, possible_matches, how="partial", required=True):
    for col in df.columns:
        col_clean = col.strip().lower()
        for pattern in possible_matches:
            pat_clean = pattern.strip().lower()
            if how == "exact":
                if col_clean == pat_clean:
                    return col
            elif how == "partial":
                if pat_clean in col_clean:
                    return col
    if required:
        raise ValueError(
            f"Could not find a required column among {possible_matches}\n"
            f"Available columns: {list(df.columns)}"
        )
    return None

def rename_columns(df, rename_map, how="partial"):
    for new_col_name, possible_names in rename_map.items():
        old_name = find_column(df, possible_names, how=how, required=False)
        if old_name:
            df.rename(columns={old_name: new_col_name}, inplace=True)
    return df

def load_data(file, sheet_name, header_row=0):
    return pd.read_excel(file, sheet_name=sheet_name, header=header_row)

def filter_exclusions_and_retained(upstream_df, midstream_df):
    """
    Rename only 'Company' + 'Fossil Fuel Share of Revenue' in Upstream,
    and 'Company' + pipeline/capacity fields in Midstream,
    then do an outer merge on 'Company' so we see all rows that have a recognized 'Company' column.
    """

    ########## 1) Show columns before rename (for debugging) ##########
    st.write("## Columns in Upstream (before rename):", list(upstream_df.columns))
    st.write("## Columns in Midstream (before rename):", list(midstream_df.columns))

    ########## 2) Rename Upstream ##########
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

    st.write("## Columns in Upstream (after rename):", list(upstream_df.columns))

    # Ensure columns exist
    for col in upstream_rename_map.keys():
        if col not in upstream_df.columns:
            upstream_df[col] = None

    # Subset
    upstream_subset = upstream_df[
        ["Company", "Fossil Fuel Share of Revenue", "BB Ticker", "ISIN Equity", "LEI"]
    ].copy()

    # Convert fossil-fuel share
    upstream_subset["Fossil Fuel Share of Revenue"] = (
        upstream_subset["Fossil Fuel Share of Revenue"].astype(str)
        .str.replace("%", "", regex=True)
    )
    upstream_subset["Fossil Fuel Share of Revenue"] = pd.to_numeric(
        upstream_subset["Fossil Fuel Share of Revenue"], errors="coerce"
    ).fillna(0)
    upstream_subset["Upstream_Exclusion_Flag"] = upstream_subset["Fossil Fuel Share of Revenue"] > 0

    ########## 3) Rename Midstream ##########
    midstream_rename_map = {
        "Company": ["company", "company name"],
        "Length of Pipelines under Development": ["length of pipelines", "pipeline under dev"],
        "Liquefaction Capacity (Export)": ["liquefaction capacity", "lng export capacity"],
        "Regasification Capacity (Import)": ["regasification capacity", "lng import capacity"],
        "Total Capacity under Development": ["total capacity under development", "total dev capacity"]
    }
    rename_columns(midstream_df, midstream_rename_map, how="partial")

    st.write("## Columns in Midstream (after rename):", list(midstream_df.columns))

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

    # Convert numeric
    numeric_cols = [
        "Length of Pipelines under Development",
        "Liquefaction Capacity (Export)",
        "Regasification Capacity (Import)",
        "Total Capacity under Development"
    ]
    for col in numeric_cols:
        midstream_subset[col] = pd.to_numeric(midstream_subset[col], errors='coerce').fillna(0)

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
    midstream_grouped["Midstream_Exclusion_Flag"] = (
        (midstream_grouped["Length of Pipelines under Development"] > 0)
        | (midstream_grouped["Liquefaction Capacity (Export)"] > 0)
        | (midstream_grouped["Regasification Capacity (Import)"] > 0)
        | (midstream_grouped["Total Capacity under Development"] > 0)
    )

    ########## 4) Merge on 'Company' ##########
    combined = pd.merge(
        upstream_subset,
        midstream_grouped,
        on="Company",
        how="outer"
    )

    # Convert flags
    combined["Upstream_Exclusion_Flag"] = combined["Upstream_Exclusion_Flag"].fillna(False).astype(bool)
    combined["Midstream_Exclusion_Flag"] = combined["Midstream_Exclusion_Flag"].fillna(False).astype(bool)
    combined["Excluded"] = combined["Upstream_Exclusion_Flag"] | combined["Midstream_Exclusion_Flag"]

    # Build reason
    reason_list = []
    for _, row in combined.iterrows():
        r = []
        if row["Upstream_Exclusion_Flag"]:
            r.append("Upstream - Fossil Fuel Share > 0%")
        if row["Midstream_Exclusion_Flag"]:
            r.append("Midstream Expansion > 0")
        reason_list.append("; ".join(r))
    combined["Exclusion Reason"] = reason_list

    # Split
    def is_empty_string_or_nan(val):
        return pd.isna(val) or str(val).strip() == ""

    no_data_cond = (
        (~combined["Excluded"])
        & combined["BB Ticker"].apply(is_empty_string_or_nan)
        & combined["ISIN Equity"].apply(is_empty_string_or_nan)
        & combined["LEI"].apply(is_empty_string_or_nan)
        & (combined["Length of Pipelines under Development"] == 0)
        & (combined["Liquefaction Capacity (Export)"] == 0)
        & (combined["Regasification Capacity (Import)"] == 0)
        & (combined["Total Capacity under Development"] == 0)
    )

    no_data_companies = combined[no_data_cond].copy()
    excluded_companies = combined[combined["Excluded"]].copy()
    retained_companies = combined[~combined["Excluded"] & ~no_data_cond].copy()
    return excluded_companies, retained_companies, no_data_companies, combined

def main():
    st.title("Debug: Which Columns Are Detected as 'Company'?")
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

    header_row = st.number_input("Select header row index", value=4, min_value=0, max_value=30)

    if uploaded_file:
        xls = pd.ExcelFile(uploaded_file)
        sheets = xls.sheet_names
        st.write("### Sheets found:", sheets)

        if ("Upstream" in sheets) and ("Midstream Expansion" in sheets):
            upstream_df = load_data(uploaded_file, sheet_name="Upstream", header_row=header_row)
            midstream_df = load_data(uploaded_file, sheet_name="Midstream Expansion", header_row=header_row)

            ex, re_, nd, combined = filter_exclusions_and_retained(upstream_df, midstream_df)

            # Stats
            excluded_count = len(ex)
            retained_count = len(re_)
            no_data_count = len(nd)
            total_count = excluded_count + retained_count + no_data_count
            st.write("### Summary")
            st.write(f"Total: {total_count}, Excluded: {excluded_count}, Retained: {retained_count}, NoData: {no_data_count}")

            # Show combined preview
            st.subheader("Combined Data (merge on 'Company') - first 50 rows")
            st.dataframe(combined.head(50))

            st.subheader("Excluded Companies")
            st.dataframe(ex)

            st.subheader("Retained Companies")
            st.dataframe(re_)

            st.subheader("No Data Companies")
            st.dataframe(nd)

        else:
            st.error("Could not find BOTH 'Upstream' and 'Midstream Expansion' sheets.")
    else:
        st.info("Please upload an Excel file first.")

if __name__ == "__main__":
    main()
