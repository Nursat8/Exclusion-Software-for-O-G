import streamlit as st
import pandas as pd
import io
from io import BytesIO

########################################
# 1) HELPER FUNCTIONS
########################################

def find_column_exact(df, exact_name):
    """
    Returns the actual column in df.columns whose case-insensitive
    name exactly matches `exact_name`. Otherwise None.
    """
    target_lower = exact_name.strip().lower()
    for col in df.columns:
        if col.strip().lower() == target_lower:
            return col
    return None

def find_column_partial(df, patterns):
    """
    Returns the first column that partially matches any of the strings in `patterns`.
    Matching is case-insensitive and checks if `pattern` is a substring of col.lower().
    """
    for col in df.columns:
        col_lower = col.strip().lower()
        for p in patterns:
            if p.strip().lower() in col_lower:
                return col
    return None

def rename_columns(df):
    """
    1) EXACT match to find "Company".
    2) PARTIAL match to find the Upstream/Midstream columns only.
       (No "BB" or "Bloomberg" columns used.)
    """
    # Strip out trailing spaces from all columns
    df.columns = df.columns.str.strip()

    ########## EXACT "Company" ##########
    found_company_col = find_column_exact(df, "Company")
    # If we found it and it's not literally named "Company", rename it:
    if found_company_col and found_company_col != "Company":
        df.rename(columns={found_company_col: "Company"}, inplace=True)
    # If we didn't find it at all, we'll create a blank "Company" column
    if "Company" not in df.columns:
        df["Company"] = None

    ########## PARTIAL matching for Upstream / Midstream fields ##########
    partial_map = {
        "Fossil Fuel Share of Revenue": ["fossil fuel share", "fossil fuel revenue"],
        "Length of Pipelines under Development": ["length of pipelines", "pipeline under dev"],
        "Liquefaction Capacity (Export)": ["liquefaction capacity (export)", "lng export capacity"],
        "Regasification Capacity (Import)": ["regasification capacity (import)", "lng import capacity"],
        "Total Capacity under Development": ["total capacity under development", "total dev capacity"]
    }
    for new_col_name, patterns in partial_map.items():
        old_col = find_column_partial(df, patterns)
        if old_col and old_col != new_col_name:
            df.rename(columns={old_col: new_col_name}, inplace=True)

########################################
# 2) CORE EXCLUSION LOGIC
########################################

def filter_all_companies(df):
    """
    Reads from "All Companies" sheet, uses EXACT "Company" column,
    partial for Upstream/Midstream columns, then splits into
    Excluded / Retained / No Data. 
    """

    # 1) Rename columns (exact for "Company", partial for Upstream/Midstream)
    rename_columns(df)

    # 2) Ensure numeric columns exist or fill them with 0
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

    # 3) Convert those columns to float
    for c in numeric_cols:
        df[c] = (
            df[c].astype(str)
            .str.replace("%", "", regex=True)
            .str.replace(",", "", regex=True)
        )
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    # 4) Upstream + Midstream flags
    df["Upstream_Exclusion_Flag"] = df["Fossil Fuel Share of Revenue"] > 0
    df["Midstream_Exclusion_Flag"] = (
        (df["Length of Pipelines under Development"] > 0)
        | (df["Liquefaction Capacity (Export)"] > 0)
        | (df["Regasification Capacity (Import)"] > 0)
        | (df["Total Capacity under Development"] > 0)
    )
    df["Excluded"] = df["Upstream_Exclusion_Flag"] | df["Midstream_Exclusion_Flag"]

    # 5) Build Exclusion Reason
    reasons = []
    for _, row in df.iterrows():
        r = []
        if row["Upstream_Exclusion_Flag"]:
            r.append("Upstream - Fossil Fuel Share > 0%")
        if row["Midstream_Exclusion_Flag"]:
            r.append("Midstream Expansion > 0")
        reasons.append("; ".join(r))
    df["Exclusion Reason"] = reasons

    # 6) No Data: all numeric columns = 0 AND not excluded
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

    excluded_data = df[df["Excluded"]].copy()
    no_data_data = df[no_data_mask].copy()
    retained_data = df[~df["Excluded"] & ~no_data_mask].copy()

    # 7) Final columns
    final_cols = [
        "Company",
        "Fossil Fuel Share of Revenue",
        "Length of Pipelines under Development",
        "Liquefaction Capacity (Export)",
        "Regasification Capacity (Import)",
        "Total Capacity under Development",
        "Exclusion Reason"
    ]

    # Fill missing columns in each subset
    for c in final_cols:
        if c not in excluded_data.columns:
            excluded_data[c] = None
        if c not in no_data_data.columns:
            no_data_data[c] = None
        if c not in retained_data.columns:
            retained_data[c] = None

    excluded_data = excluded_data[final_cols]
    no_data_data = no_data_data[final_cols]
    retained_data = retained_data[final_cols]

    return excluded_data, retained_data, no_data_data

########################################
# 3) STREAMLIT APP
########################################

def main():
    st.title("All Companies Exclusion Filter – Only 'Company' Column (No BB/Bloomberg)")

    uploaded_file = st.file_uploader("Upload Excel file (has 'All Companies' sheet)", type=["xlsx"])

    # Adjust if your real column headers are on a different row
    header_row = 4

    if uploaded_file:
        xls = pd.ExcelFile(uploaded_file)
        if "All Companies" not in xls.sheet_names:
            st.error("No sheet named 'All Companies'.")
            return

        # Load the sheet
        df_all = pd.read_excel(uploaded_file, sheet_name="All Companies", header=header_row)

        # Run logic
        excluded, retained, no_data = filter_all_companies(df_all)
        total_count = len(excluded) + len(retained) + len(no_data)

        st.subheader("Statistics")
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
    else:
        st.info("Please upload an Excel file first.")

if __name__ == "__main__":
    main()
