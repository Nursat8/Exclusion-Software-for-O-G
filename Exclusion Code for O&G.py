import streamlit as st
import pandas as pd
import io
from io import BytesIO

################################
# 1) HELPER FUNCTIONS
################################

def find_column_exact(df, target):
    """
    Returns the actual column whose .strip().lower() equals target.lower().
    Otherwise returns None.
    """
    target_lower = target.strip().lower()
    for col in df.columns:
        if col.strip().lower() == target_lower:
            return col
    return None

def find_column_partial(df, patterns):
    """
    Returns the first column that partially matches any item in `patterns`.
    Matching is case-insensitive, substring-based.
    E.g., if patterns = ["fossil fuel share"], any column containing
    "fossil fuel share" in its name (ignoring case) is matched.
    """
    for col in df.columns:
        col_lower = col.strip().lower()
        for pat in patterns:
            pat_lower = pat.strip().lower()
            if pat_lower in col_lower:
                return col
    return None

def rename_columns(df):
    """
    1) EXACT match for 'Company' => rename that column to 'Company' if found
       (ignoring columns like 'Company Name in Bloomberg').
    2) PARTIAL match for Upstream & Midstream fields.
    """

    # Strip trailing spaces from all columns
    df.columns = df.columns.str.strip()

    # ---------- EXACT match for "Company" ----------
    found_company = find_column_exact(df, "Company")
    if found_company and found_company != "Company":
        df.rename(columns={found_company: "Company"}, inplace=True)

    # If not found at all, create a blank "Company" so code won't crash
    if "Company" not in df.columns:
        df["Company"] = None

    # ---------- PARTIAL match for Upstream/Midstream ----------
    partial_map = {
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

    for new_col, patterns in partial_map.items():
        old_col = find_column_partial(df, patterns)
        if old_col and old_col != new_col:
            df.rename(columns={old_col: new_col}, inplace=True)

################################
# 2) CORE EXCLUSION LOGIC
################################

def filter_all_companies(df):
    """
    Reads the 'All Companies' sheet:
      - uses only the 'Company' column for company names (no 'BB Company' columns).
      - partial-match for Upstream & Midstream columns.
      - no row dropping.
    Splits into Excluded / Retained / No Data.
    """

    # 1) Rename columns
    rename_columns(df)

    # 2) Make sure numeric columns exist or fill with 0 if missing
    numeric_cols = [
        "Fossil Fuel Share of Revenue",
        "Length of Pipelines under Development",
        "Liquefaction Capacity (Export)",
        "Regasification Capacity (Import)",
        "Total Capacity under Development"
    ]
    for c in numeric_cols:
        if c not in df.columns:
            df[c] = 0

    # 3) Also ensure 'Company', 'BB Ticker', 'ISIN Equity', 'LEI' exist
    for c in ["Company", "BB Ticker", "ISIN Equity", "LEI"]:
        if c not in df.columns:
            df[c] = None

    # 4) Convert numeric columns
    for c in numeric_cols:
        df[c] = (
            df[c].astype(str)
            .str.replace("%", "", regex=True)
            .str.replace(",", "", regex=True)
        )
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    # 5) Upstream Exclusion: Fossil Fuel Share > 0
    df["Upstream_Exclusion_Flag"] = df["Fossil Fuel Share of Revenue"] > 0

    # 6) Midstream Exclusion: any pipeline/capacity > 0
    df["Midstream_Exclusion_Flag"] = (
        (df["Length of Pipelines under Development"] > 0)
        | (df["Liquefaction Capacity (Export)"] > 0)
        | (df["Regasification Capacity (Import)"] > 0)
        | (df["Total Capacity under Development"] > 0)
    )

    df["Excluded"] = df["Upstream_Exclusion_Flag"] | df["Midstream_Exclusion_Flag"]

    # 7) Build Exclusion Reason
    def build_reason(row):
        r = []
        if row["Upstream_Exclusion_Flag"]:
            r.append("Upstream - Fossil Fuel Share > 0%")
        if row["Midstream_Exclusion_Flag"]:
            r.append("Midstream Expansion > 0")
        return "; ".join(r)

    df["Exclusion Reason"] = df.apply(build_reason, axis=1)

    # 8) Define "No Data" => all numeric=0 and not Excluded
    def is_no_data(r):
        numeric_zero = (
            (r["Fossil Fuel Share of Revenue"] == 0)
            and (r["Length of Pipelines under Development"] == 0)
            and (r["Liquefaction Capacity (Export)"] == 0)
            and (r["Regasification Capacity (Import)"] == 0)
            and (r["Total Capacity under Development"] == 0)
        )
        return numeric_zero and not r["Excluded"]

    no_data_mask = df.apply(is_no_data, axis=1)

    excluded_df = df[df["Excluded"]].copy()
    no_data_df = df[no_data_mask].copy()
    retained_df = df[~df["Excluded"] & ~no_data_mask].copy()

    # 9) Final columns
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

    for c in final_cols:
        if c not in excluded_df.columns:
            excluded_df[c] = None
        if c not in no_data_df.columns:
            no_data_df[c] = None
        if c not in retained_df.columns:
            retained_df[c] = None

    return excluded_df[final_cols], retained_df[final_cols], no_data_df[final_cols]

################################
# 3) STREAMLIT APP
################################

def main():
    st.title("All Companies Exclusion Filter – Only 'Company' Column")

    uploaded_file = st.file_uploader("Upload an Excel file with 'All Companies' sheet", type=["xlsx"])

    # If your actual header row is different, adjust here
    header_row = 4  

    if uploaded_file:
        xls = pd.ExcelFile(uploaded_file)
        if "All Companies" not in xls.sheet_names:
            st.error("No sheet called 'All Companies' found!")
            return

        df_all = pd.read_excel(uploaded_file, sheet_name="All Companies", header=header_row)

        # Filter logic
        excluded, retained, no_data = filter_all_companies(df_all)

        total_count = len(excluded) + len(retained) + len(no_data)

        st.subheader("Summary")
        st.write(f"**Total:** {total_count}")
        st.write(f"**Excluded:** {len(excluded)}")
        st.write(f"**Retained:** {len(retained)}")
        st.write(f"**No Data:** {len(no_data)}")

        st.subheader("Excluded Companies")
        st.dataframe(excluded)

        st.subheader("Retained Companies")
        st.dataframe(retained)

        st.subheader("No Data Companies")
        st.dataframe(no_data)

        # Export to Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            excluded.to_excel(writer, sheet_name="Excluded", index=False)
            retained.to_excel(writer, sheet_name="Retained", index=False)
            no_data.to_excel(writer, sheet_name="No Data", index=False)
        output.seek(0)

        st.download_button(
            label="Download Exclusion Results",
            data=output,
            file_name="All_Companies_Exclusion_Results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
