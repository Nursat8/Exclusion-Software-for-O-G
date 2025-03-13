import streamlit as st
import pandas as pd
import io
from io import BytesIO

################################
# 1) HELPER FUNCTIONS
################################

def find_column(df, possible_matches, how="partial", required=True):
    """Finds the best-matching column from a list of possible names."""
    for col in df.columns:
        col_lower = col.strip().lower()
        for pattern in possible_matches:
            pat_lower = pattern.strip().lower()
            if how == "exact" and col_lower == pat_lower:
                return col
            elif how == "partial" and pat_lower in col_lower:
                return col
    if required:
        raise ValueError(f"Could not find a required column among {possible_matches}\nAvailable columns: {list(df.columns)}")
    return None

def rename_columns(df, rename_map, how="partial"):
    """Renames columns based on a dictionary of possible matches."""
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
    - Reads 'All Companies' sheet
    - Uses **only the 'Company' column** (not 'Company Name in Bloomberg')
    - Applies exclusion criteria
    - Keeps all rows, even if they are missing data
    """

    ########## 1) RENAME COLUMNS ##########
    rename_map = {
        "Company": ["company"],  # Ensuring we use the correct "Company" column!
        "Fossil Fuel Share of Revenue": ["fossil fuel share of revenue", "fossil fuel revenue"],
        "Length of Pipelines under Development": ["length of pipelines under development", "length of pipelines"],
        "Liquefaction Capacity (Export)": ["liquefaction capacity (export)", "lng export capacity"],
        "Regasification Capacity (Import)": ["regasification capacity (import)", "lng import capacity"],
        "Total Capacity under Development": ["total capacity under development", "total dev capacity"],
        "BB Ticker": ["bb ticker", "bloomberg ticker"],
        "ISIN Equity": ["isin equity", "isin code"],
        "LEI": ["lei"]
    }
    rename_columns(df, rename_map, how="partial")

    ########## 2) ENSURE REQUIRED COLUMNS EXIST ##########
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

    for col in ["Company", "BB Ticker", "ISIN Equity", "LEI"]:
        if col not in df.columns:
            df[col] = None

    ########## 3) CONVERT NUMERIC COLUMNS ##########
    for col in needed_for_numeric:
        df[col] = (
            df[col]
            .astype(str)
            .str.replace("%", "", regex=True)
            .str.replace(",", "", regex=True)
        )
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    ########## 4) APPLY EXCLUSION CRITERIA ##########
    df["Upstream_Exclusion_Flag"] = df["Fossil Fuel Share of Revenue"] > 0
    df["Midstream_Exclusion_Flag"] = (
        (df["Length of Pipelines under Development"] > 0)
        | (df["Liquefaction Capacity (Export)"] > 0)
        | (df["Regasification Capacity (Import)"] > 0)
        | (df["Total Capacity under Development"] > 0)
    )
    df["Excluded"] = df["Upstream_Exclusion_Flag"] | df["Midstream_Exclusion_Flag"]

    ########## 5) BUILD EXCLUSION REASON ##########
    reasons = []
    for _, row in df.iterrows():
        r = []
        if row["Upstream_Exclusion_Flag"]:
            r.append("Upstream - Fossil Fuel Share > 0%")
        if row["Midstream_Exclusion_Flag"]:
            r.append("Midstream Expansion > 0")
        reasons.append("; ".join(r))
    df["Exclusion Reason"] = reasons

    ########## 6) IDENTIFY 'NO DATA' COMPANIES ##########
    def is_no_data(row):
        numeric_zero = (
            (row["Fossil Fuel Share of Revenue"] == 0)
            and (row["Length of Pipelines under Development"] == 0)
            and (row["Liquefaction Capacity (Export)"] == 0)
            and (row["Regasification Capacity (Import)"] == 0)
            and (row["Total Capacity under Development"] == 0)
            and (not row["Excluded"])
        )
        return numeric_zero

    no_data_mask = df.apply(is_no_data, axis=1)

    ########## 7) SPLIT INTO FINAL CATEGORIES ##########
    excluded_companies = df[df["Excluded"]].copy()
    no_data_companies = df[no_data_mask].copy()
    retained_companies = df[~df["Excluded"] & ~no_data_mask].copy()

    ########## 8) KEEP ONLY REQUIRED COLUMNS ##########
    final_cols = [
        "Company", "BB Ticker", "ISIN Equity", "LEI",
        "Fossil Fuel Share of Revenue",
        "Length of Pipelines under Development",
        "Liquefaction Capacity (Export)",
        "Regasification Capacity (Import)",
        "Total Capacity under Development",
        "Exclusion Reason"
    ]

    for c in final_cols:
        if c not in excluded_companies.columns:
            excluded_companies[c] = None
        if c not in retained_companies.columns:
            retained_companies[c] = None
        if c not in no_data_companies.columns:
            no_data_companies[c] = None

    return (
        excluded_companies[final_cols],
        retained_companies[final_cols],
        no_data_companies[final_cols]
    )

################################
# 3) STREAMLIT APP
################################

def main():
    st.title("All Companies Exclusion Filter (Now Using Correct 'Company' Column)")

    uploaded_file = st.file_uploader("Upload Excel file with a sheet named 'All Companies'", type=["xlsx"])

    header_row = 4  # Change if needed

    if uploaded_file:
        xls = pd.ExcelFile(uploaded_file)
        if "All Companies" not in xls.sheet_names:
            st.error("No sheet called 'All Companies' found!")
            return

        df_all = pd.read_excel(uploaded_file, sheet_name="All Companies", header=header_row)

        # Run filtering logic
        excluded_data, retained_data, no_data_data = filter_all_companies(df_all)

        # Stats
        total_count = len(excluded_data) + len(retained_data) + len(no_data_data)

        st.subheader("Summary")
        st.write(f"**Total Companies Processed:** {total_count}")
        st.write(f"**Excluded:** {len(excluded_data)}")
        st.write(f"**Retained:** {len(retained_data)}")
        st.write(f"**No Data:** {len(no_data_data)}")

        # Export to Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            excluded_data.to_excel(writer, sheet_name="Excluded", index=False)
            retained_data.to_excel(writer, sheet_name="Retained", index=False)
            no_data_data.to_excel(writer, sheet_name="No Data", index=False)
        output.seek(0)

        st.download_button("Download Results", output, "Filtered_Companies.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
