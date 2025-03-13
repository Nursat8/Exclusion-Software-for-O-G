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

def load_data(file, sheet_name, header_row=0):
    """Helper to load the specified sheet from Excel, using the given header row."""
    return pd.read_excel(file, sheet_name=sheet_name, header=header_row)


################################
# 2) CORE EXCLUSION LOGIC
################################

def filter_all_companies(df):
    """
    Takes a single DataFrame (`df`) from "All Companies" sheet, which must have:
     - "Company"
     - "Fossil Fuel Share of Revenue"
     - Possibly midstream columns:
         "Length of Pipelines under Development"
         "Liquefaction Capacity (Export)"
         "Regasification Capacity (Import)"
         "Total Capacity under Development"

    Then:
      1) Upstream_Exclusion_Flag = (Fossil Fuel Share of Revenue > 0)
      2) Midstream_Exclusion_Flag = any midstream capacity > 0
      3) Excluded if either flag is true
      4) Retained if not excluded
      5) "No Data" if all relevant numeric columns are 0 or blank
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
        ]
    }
    rename_columns(df, rename_map, how="partial")

    # 2) Make sure we at least have columns for them (fill with 0 or None)
    needed_for_numeric = [
        "Fossil Fuel Share of Revenue",
        "Length of Pipelines under Development",
        "Liquefaction Capacity (Export)",
        "Regasification Capacity (Import)",
        "Total Capacity under Development"
    ]
    for col in needed_for_numeric:
        if col not in df.columns:
            df[col] = 0  # no data => 0

    if "Company" not in df.columns:
        df["Company"] = None  # fallback, though it means we have no company name

    # 3) Convert to numeric
    numeric_cols = needed_for_numeric
    for col in numeric_cols:
        df[col] = (
            df[col]
            .astype(str)
            .str.replace("%", "", regex=True)
            .str.replace(",", "", regex=True)
        )
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # 4) Build flags
    df["Upstream_Exclusion_Flag"] = df["Fossil Fuel Share of Revenue"] > 0
    df["Midstream_Exclusion_Flag"] = (
        (df["Length of Pipelines under Development"] > 0)
        | (df["Liquefaction Capacity (Export)"] > 0)
        | (df["Regasification Capacity (Import)"] > 0)
        | (df["Total Capacity under Development"] > 0)
    )

    # 5) Exclude if either upstream or midstream
    df["Excluded"] = df["Upstream_Exclusion_Flag"] | df["Midstream_Exclusion_Flag"]

    # 6) Build Exclusion Reason
    reasons = []
    for _, row in df.iterrows():
        r = []
        if row["Upstream_Exclusion_Flag"]:
            r.append("Upstream - Fossil Fuel Share > 0%")
        if row["Midstream_Exclusion_Flag"]:
            r.append("Midstream Expansion > 0")
        reasons.append("; ".join(r))
    df["Exclusion Reason"] = reasons

    # 7) "No Data" means all relevant numeric columns = 0 (and not excluded)
    def is_no_data(row):
        # If everything is zero, no fossil share, no midstream dev
        # plus no Ticker or LEI, you can define logic or skip
        cond_numeric = (
            (row["Fossil Fuel Share of Revenue"] == 0)
            & (row["Length of Pipelines under Development"] == 0)
            & (row["Liquefaction Capacity (Export)"] == 0)
            & (row["Regasification Capacity (Import)"] == 0)
            & (row["Total Capacity under Development"] == 0)
        )
        return (cond_numeric and not row["Excluded"])

    no_data_mask = df.apply(is_no_data, axis=1)

    # 8) Subset into final data
    no_data_companies = df[no_data_mask].copy()
    excluded_companies = df[df["Excluded"]].copy()
    retained_companies = df[~df["Excluded"] & ~no_data_mask].copy()

    return excluded_companies, retained_companies, no_data_companies


################################
# 3) STREAMLIT APP
################################

def main():
    st.title("Exclusion from a Single 'All Companies' Sheet")

    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

    header_row = st.number_input("Header row index (0-based)", min_value=0, max_value=50, value=0,
                                 help="If your columns start on row 2 in Excel, use header_row=1, etc.")

    if uploaded_file:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        st.write("Available sheets:", sheet_names)

        if "All Companies" not in sheet_names:
            st.error("Sheet named 'All Companies' not found. Please check your Excel file or rename the sheet.")
            return

        # Load the single sheet
        df_all = load_data(uploaded_file, sheet_name="All Companies", header_row=header_row)
        st.write("Dataframe shape:", df_all.shape)

        # Apply the logic
        excluded_data, retained_data, no_data_data = filter_all_companies(df_all)

        # Stats
        excl_count = len(excluded_data)
        ret_count = len(retained_data)
        nodata_count = len(no_data_data)
        total_count = excl_count + ret_count + nodata_count

        st.markdown("### Statistics")
        st.write(f"Total: {total_count}")
        st.write(f"Excluded: {excl_count}")
        st.write(f"Retained: {ret_count}")
        st.write(f"No Data: {nodata_count}")

        st.subheader("Excluded Companies")
        st.dataframe(excluded_data)

        st.subheader("Retained Companies")
        st.dataframe(retained_data)

        st.subheader("No Data Companies")
        st.dataframe(no_data_data)

        # Export to Excel in memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            excluded_data.to_excel(writer, index=False, sheet_name="Excluded")
            retained_data.to_excel(writer, index=False, sheet_name="Retained")
            no_data_data.to_excel(writer, index=False, sheet_name="No Data")
        output.seek(0)

        st.download_button(
            label="Download Exclusion Results",
            data=output,
            file_name="All_Companies_Exclusion_Results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("Please upload an Excel file.")

if __name__ == "__main__":
    main()
