import streamlit as st
import pandas as pd
import io
from io import BytesIO

################################
# 1) HELPER FUNCTIONS
################################

def find_column_exact(df, target_name):
    """
    Searches df.columns for an EXACT case-insensitive match to `target_name`.
    Returns the actual column name, or None if not found.
    """
    target_lower = target_name.strip().lower()
    for col in df.columns:
        if col.strip().lower() == target_lower:
            return col
    return None

def find_column_partial(df, possible_matches):
    """
    Searches df.columns for the first partial (case-insensitive) match
    to any pattern in `possible_matches`.
    Returns the actual column name, or None if not found.
    """
    for col in df.columns:
        col_lower = col.strip().lower()
        for pattern in possible_matches:
            pat_lower = pattern.strip().lower()
            if pat_lower in col_lower:
                return col
    return None

def rename_columns(df):
    """
    1) EXACT-match "Company" => rename to "Company".
    2) PARTIAL-match for everything else (fossil fuel share, midstream capacities, etc.).
    """
    # 1) EXACT match for "Company"
    #    If your real file’s column is exactly "Company" in Excel, this step may do nothing.
    #    But if it was e.g. "COMPANY" or "Company " (with trailing space), we fix it.
    company_col = find_column_exact(df, "Company")
    if company_col and company_col != "Company":
        df.rename(columns={company_col: "Company"}, inplace=True)

    # 2) PARTIAL match for the rest
    partial_map = {
        "Fossil Fuel Share of Revenue": [
            "fossil fuel share of revenue",
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

    for new_col_name, patterns in partial_map.items():
        old_col = find_column_partial(df, patterns)
        if old_col and old_col != new_col_name:
            df.rename(columns={old_col: new_col_name}, inplace=True)

################################
# 2) CORE EXCLUSION LOGIC
################################

def filter_all_companies(df):
    """
    - Reads from 'All Companies' sheet
    - Uses EXACT 'Company' (not 'Company Name in Bloomberg')
    - Applies exclusion criteria
    - Keeps all rows (no removing blanks)
    """

    # Rename columns
    rename_columns(df)

    # Ensure numeric columns exist or fill with zero
    numeric_cols = [
        "Fossil Fuel Share of Revenue",
        "Length of Pipelines under Development",
        "Liquefaction Capacity (Export)",
        "Regasification Capacity (Import)",
        "Total Capacity under Development"
    ]
    for col in numeric_cols:
        if col not in df.columns:
            df[col] = 0

    # Ensure these columns exist
    for col in ["Company", "BB Ticker", "ISIN Equity", "LEI"]:
        if col not in df.columns:
            df[col] = None

    # Convert numeric
    for col in numeric_cols:
        df[col] = (
            df[col]
            .astype(str)
            .str.replace("%", "", regex=True)
            .str.replace(",", "", regex=True)
        )
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # Upstream / Midstream flags
    df["Upstream_Exclusion_Flag"] = df["Fossil Fuel Share of Revenue"] > 0
    df["Midstream_Exclusion_Flag"] = (
        (df["Length of Pipelines under Development"] > 0)
        | (df["Liquefaction Capacity (Export)"] > 0)
        | (df["Regasification Capacity (Import)"] > 0)
        | (df["Total Capacity under Development"] > 0)
    )
    df["Excluded"] = df["Upstream_Exclusion_Flag"] | df["Midstream_Exclusion_Flag"]

    # Build Exclusion Reason
    reasons = []
    for _, row in df.iterrows():
        r = []
        if row["Upstream_Exclusion_Flag"]:
            r.append("Upstream - Fossil Fuel Share > 0%")
        if row["Midstream_Exclusion_Flag"]:
            r.append("Midstream Expansion > 0")
        reasons.append("; ".join(r))
    df["Exclusion Reason"] = reasons

    # No Data = numeric columns = 0 AND not excluded
    def is_no_data(r):
        numeric_zero = (
            (r["Fossil Fuel Share of Revenue"] == 0)
            and (r["Length of Pipelines under Development"] == 0)
            and (r["Liquefaction Capacity (Export)"] == 0)
            and (r["Regasification Capacity (Import)"] == 0)
            and (r["Total Capacity under Development"] == 0)
            and (not r["Excluded"])
        )
        return numeric_zero

    no_data_mask = df.apply(is_no_data, axis=1)

    excluded_companies = df[df["Excluded"]].copy()
    no_data_companies = df[no_data_mask].copy()
    retained_companies = df[~df["Excluded"] & ~no_data_mask].copy()

    # Final columns
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

    # Fill missing in each final subset
    for c in final_cols:
        if c not in excluded_companies.columns:
            excluded_companies[c] = None
        if c not in no_data_companies.columns:
            no_data_companies[c] = None
        if c not in retained_companies.columns:
            retained_companies[c] = None

    excluded_companies = excluded_companies[final_cols]
    no_data_companies = no_data_companies[final_cols]
    retained_companies = retained_companies[final_cols]

    return excluded_companies, retained_companies, no_data_companies

################################
# 3) STREAMLIT APP
################################

def main():
    st.title("All Companies Exclusion Filter (EXACT 'Company', not 'Company Name in Bloomberg')")
    uploaded_file = st.file_uploader("Upload Excel file with 'All Companies' sheet", type=["xlsx"])

    # If your real data starts at row 5 in Excel, use header=4, etc.
    header_row = 4

    if uploaded_file:
        xls = pd.ExcelFile(uploaded_file)
        if "All Companies" not in xls.sheet_names:
            st.error("Sheet 'All Companies' not found in uploaded file.")
            return

        df_all = pd.read_excel(uploaded_file, sheet_name="All Companies", header=header_row)

        # Filter logic
        excluded, retained, no_data = filter_all_companies(df_all)

        total_count = len(excluded) + len(retained) + len(no_data)
        st.write("### Statistics")
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

        # Download
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
