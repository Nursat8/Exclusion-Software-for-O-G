import streamlit as st
import pandas as pd
import io
from io import BytesIO

################################
# 1) HELPER FUNCTIONS
################################

def find_column(df, possible_matches, how="partial", required=True):
    """
    Searches df.columns for the first column that matches any in `possible_matches`.
    Returns the actual column name, or None if not found (and required=False).
    """
    for col in df.columns:
        col_lower = col.strip().lower()
        for pattern in possible_matches:
            pat_lower = pattern.strip().lower()
            if how == "exact":
                if col_lower == pat_lower:
                    return col
            elif how == "partial":
                if pat_lower in col_lower:
                    return col
    if required:
        raise ValueError(
            f"Could not find a required column among {possible_matches}\n"
            f"Available columns: {list(df.columns)}"
        )
    return None

def rename_columns(df, rename_map, how="partial"):
    """
    Given a dict like:
        {
          "Company": ["company", "company name"],
          "Fossil Fuel Share of Revenue": ["fossil fuel share", "fossil fuel revenue"]
        }
    we look for each pattern and rename in-place if found.
    """
    for new_col_name, patterns in rename_map.items():
        old_name = find_column(df, patterns, how=how, required=False)
        if old_name:
            df.rename(columns={old_name: new_col_name}, inplace=True)
    return df

################################
# 2) CORE EXCLUSION LOGIC
################################

def filter_all_companies(df):
    """
    Takes a single DataFrame (`df`) from "All Companies" sheet, which must have:
     - "Company"
     - "Fossil Fuel Share of Revenue"
     - Midstream columns:
         "Length of Pipelines under Development"
         "Liquefaction Capacity (Export)"
         "Regasification Capacity (Import)"
         "Total Capacity under Development"

    Then:
      1) Upstream_Exclusion_Flag = (Fossil Fuel Share of Revenue > 0)
      2) Midstream_Exclusion_Flag = any midstream capacity > 0
      3) Excluded if either upstream or midstream
      4) Retained if not excluded
      5) "No Data" if all relevant numeric columns are 0 (and not excluded)
    """

    # 1) RENAME columns we care about
    rename_map = {
        "Company": ["company", "company name"],
        "Fossil Fuel Share of Revenue": [
            "fossil fuel share of revenue",
            "fossil fuel share",
            "fossil fuel revenue"
        ],
        "Length of Pipelines under Development": [
            "length of pipelines under development",
            "length of pipelines"
        ],
        "Liquefaction Capacity (Export)": [
            "liquefaction capacity (export)",
            "lng export capacity"
        ],
        "Regasification Capacity (Import)": [
            "regasification capacity (import)",
            "lng import capacity"
        ],
        "Total Capacity under Development": [
            "total capacity under development",
            "total dev capacity"
        ],
        "BB Ticker": ["bb ticker", "bloomberg ticker"],
        "ISIN Equity": ["isin equity", "isin code"],
        "LEI": ["lei"]
    }
    rename_columns(df, rename_map, how="partial")

    # 2) Ensure numeric columns exist or fill with zero
    needed_for_numeric = [
        "Fossil Fuel Share of Revenue",
        "Length of Pipelines under Development",
        "Liquefaction Capacity (Export)",
        "Regasification Capacity (Import)",
        "Total Capacity under Development"
    ]
    for col in needed_for_numeric:
        if col not in df.columns:
            df[col] = 0

    # Ensure "Company", "BB Ticker", "ISIN Equity", "LEI" exist
    for col in ["Company", "BB Ticker", "ISIN Equity", "LEI"]:
        if col not in df.columns:
            df[col] = None

    # 3) Convert numeric columns
    for col in needed_for_numeric:
        df[col] = (
            df[col]
            .astype(str)
            .str.replace("%", "", regex=True)
            .str.replace(",", "", regex=True)
        )
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # 4) Set flags
    df["Upstream_Exclusion_Flag"] = df["Fossil Fuel Share of Revenue"] > 0
    df["Midstream_Exclusion_Flag"] = (
        (df["Length of Pipelines under Development"] > 0)
        | (df["Liquefaction Capacity (Export)"] > 0)
        | (df["Regasification Capacity (Import)"] > 0)
        | (df["Total Capacity under Development"] > 0)
    )
    df["Excluded"] = df["Upstream_Exclusion_Flag"] | df["Midstream_Exclusion_Flag"]

    # 5) Build reason
    reasons = []
    for _, row in df.iterrows():
        r = []
        if row["Upstream_Exclusion_Flag"]:
            r.append("Upstream - Fossil Fuel Share > 0%")
        if row["Midstream_Exclusion_Flag"]:
            r.append("Midstream Expansion > 0")
        reasons.append("; ".join(r))
    df["Exclusion Reason"] = reasons

    # 6) "No Data" => all numeric = 0 and not excluded
    def is_no_data(row):
        if (
            (row["Fossil Fuel Share of Revenue"] == 0)
            and (row["Length of Pipelines under Development"] == 0)
            and (row["Liquefaction Capacity (Export)"] == 0)
            and (row["Regasification Capacity (Import)"] == 0)
            and (row["Total Capacity under Development"] == 0)
            and (not row["Excluded"])
        ):
            return True
        return False

    no_data_mask = df.apply(is_no_data, axis=1)

    # 7) Split out
    excluded_companies = df[df["Excluded"]].copy()
    no_data_companies = df[no_data_mask].copy()
    retained_companies = df[~df["Excluded"] & ~no_data_mask].copy()

    # 8) Keep only certain columns in the final output
    final_cols = [
        "Company",
        "BB Ticker",
        "ISIN Equity",
        "LEI",
        "Fossil Fuel Share of Revenue",
        "Length of Pipelines under Development",
        "Liquefaction Capacity (Export)",
        "Regasification Capacity (Import)",
        "Total Capacity under Development",
        "Exclusion Reason"
    ]
    # If any of these columns are missing, fill them so the indexing won't fail
    for c in final_cols:
        if c not in excluded_companies.columns:
            excluded_companies[c] = None
        if c not in retained_companies.columns:
            retained_companies[c] = None
        if c not in no_data_companies.columns:
            no_data_companies[c] = None

    excluded_companies = excluded_companies[final_cols]
    retained_companies = retained_companies[final_cols]
    no_data_companies = no_data_companies[final_cols]

    return excluded_companies, retained_companies, no_data_companies

################################
# 3) STREAMLIT APP
################################

def main():
    st.title("All Companies Exclusion Filter (Minimal)")

    uploaded_file = st.file_uploader(
        "Upload Excel file with a sheet named 'All Companies'", 
        type=["xlsx"]
    )

    # We'll just fix the header row to 4 (adjust if needed).
    # If your columns start on row 1 in Excel, set header_row=0, etc.
    header_row = 4  

    if uploaded_file:
        # Load the single sheet
        xls = pd.ExcelFile(uploaded_file)
        if "All Companies" not in xls.sheet_names:
            st.error("No sheet called 'All Companies' found!")
            return

        df_all = pd.read_excel(uploaded_file, sheet_name="All Companies", header=header_row)

        # Run the logic
        excluded_data, retained_data, no_data_data = filter_all_companies(df_all)

        # Stats
        excl_count = len(excluded_data)
        ret_count = len(retained_data)
        nodata_count = len(no_data_data)
        total_count = excl_count + ret_count + nodata_count

        st.subheader("Summary")
        st.write(f"**Total:** {total_count}")
        st.write(f"**Excluded:** {excl_count}")
        st.write(f"**Retained:** {ret_count}")
        st.write(f"**No Data:** {nodata_count}")

        st.subheader("Excluded Companies")
        st.dataframe(excluded_data)

        st.subheader("Retained Companies")
        st.dataframe(retained_data)

        st.subheader("No Data Companies")
        st.dataframe(no_data_data)

        # Export to Excel in memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            excluded_data.to_excel(writer, sheet_name="Excluded", index=False)
            retained_data.to_excel(writer, sheet_name="Retained", index=False)
            no_data_data.to_excel(writer, sheet_name="No Data", index=False)
        output.seek(0)

        st.download_button(
            label="Download Exclusion Results",
            data=output,
            file_name="All_Companies_Exclusion_Results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
