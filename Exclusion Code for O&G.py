import streamlit as st
import pandas as pd
import io
from io import BytesIO

#########################
# 1) HELPER FUNCTIONS
#########################

def flatten_multilevel_columns(df):
    """Flatten multi-level column headers into single strings."""
    df.columns = [
        " ".join(str(level) for level in col).strip()
        for col in df.columns
    ]
    return df

def find_column(df, possible_matches, how="partial", required=True):
    """Finds the first column matching any item in possible_matches."""
    for col in df.columns:
        col_lower = col.strip().lower()
        for pattern in possible_matches:
            pat_lower = pattern.strip().lower()
            if how == "exact" and col_lower == pat_lower:
                return col
            elif how == "partial" and pat_lower in col_lower:
                return col
    if required:
        raise ValueError(
            f"Could not find a required column among {possible_matches}\n"
            f"Available columns: {list(df.columns)}"
        )
    return None

def rename_columns(df):
    """
    Flatten multi-level headers, skip first 3 rows of data (so row 7 in Excel = row 0 in DF),
    then rename columns dynamically.
    """
    df = flatten_multilevel_columns(df)
    
    # Skip 3 rows so row 7 in Excel aligns to row 0
    df = df.iloc[3:].copy()
    df.reset_index(drop=True, inplace=True)

    rename_map = {
        "Company": ["company"],  # Ensuring we get the right "Company" column dynamically
        "Fossil Fuel Share of Revenue": ["fossil fuel share of revenue", "fossil fuel share"],
        "Length of Pipelines under Development": ["length of pipelines", "pipeline under dev"],
        "Liquefaction Capacity (Export)": ["liquefaction capacity (export)", "lng export capacity"],
        "Regasification Capacity (Import)": ["regasification capacity (import)", "lng import capacity"],
        "Total Capacity under Development": ["total capacity under development", "total dev capacity"]
    }

    for new_col, patterns in rename_map.items():
        old_col = find_column(df, patterns, how="partial", required=False)
        if old_col and old_col != new_col:
            df.rename(columns={old_col: new_col}, inplace=True)

    return df

#########################
# 2) CORE EXCLUSION LOGIC
#########################

def filter_all_companies(df):
    """Parses 'All Companies' sheet, applies exclusion logic, and splits into categories."""
    
    # 1) Flatten headers, rename columns
    df = rename_columns(df)

    # 2) Ensure numeric columns exist
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

    # 3) Convert numeric columns
    for c in numeric_cols:
        df[c] = (
            df[c].astype(str)
            .str.replace("%", "", regex=True)
            .str.replace(",", "", regex=True)
        )
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    # 4) Ensure "Company" is picked up correctly
    if "Company" not in df.columns:
        df["Company"] = None

    # 5) Apply Exclusion Logic
    df["Upstream_Exclusion_Flag"] = df["Fossil Fuel Share of Revenue"] > 0
    df["Midstream_Exclusion_Flag"] = (
        (df["Length of Pipelines under Development"] > 0)
        | (df["Liquefaction Capacity (Export)"] > 0)
        | (df["Regasification Capacity (Import)"] > 0)
        | (df["Total Capacity under Development"] > 0)
    )
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

    # 7) Identify No Data Companies (All Numeric = 0, Not Excluded)
    def is_no_data(r):
        zeroes = (
            (r["Fossil Fuel Share of Revenue"] == 0)
            and (r["Length of Pipelines under Development"] == 0)
            and (r["Liquefaction Capacity (Export)"] == 0)
            and (r["Regasification Capacity (Import)"] == 0)
            and (r["Total Capacity under Development"] == 0)
        )
        return zeroes and not r["Excluded"]

    no_data_mask = df.apply(is_no_data, axis=1)

    # 8) Split into categories
    excluded_df = df[df["Excluded"]].copy()
    no_data_df = df[no_data_mask].copy()
    retained_df = df[~df["Excluded"] & ~no_data_mask].copy()

    # 9) Keep only required columns
    final_cols = [
        "Company",
        "Fossil Fuel Share of Revenue",
        "Length of Pipelines under Development",
        "Liquefaction Capacity (Export)",
        "Regasification Capacity (Import)",
        "Total Capacity under Development",
        "Exclusion Reason"
    ]
    for c in final_cols:
        for d in [excluded_df, retained_df, no_data_df]:
            if c not in d.columns:
                d[c] = None

    return excluded_df[final_cols], retained_df[final_cols], no_data_df[final_cols]

#########################
# 3) STREAMLIT APP
#########################

def main():
    st.title("All Companies Exclusion Analysis (Correcting Row 7 Issue)")

    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

    if uploaded_file:
        xls = pd.ExcelFile(uploaded_file)

        if "All Companies" not in xls.sheet_names:
            st.error("No sheet named 'All Companies'.")
            return

        # Read with multi-level headers from rows 3 & 4 (0-based)
        df_all = pd.read_excel(
            uploaded_file,
            sheet_name="All Companies",
            header=[3,4]
        )

        excluded, retained, no_data = filter_all_companies(df_all)

        # STATS
        total_companies = len(excluded) + len(retained) + len(no_data)
        st.subheader("Summary Statistics")
        st.write(f"**Total Companies Processed:** {total_companies}")
        st.write(f"**Excluded Companies:** {len(excluded)}")
        st.write(f"**Retained Companies:** {len(retained)}")
        st.write(f"**No Data Companies:** {len(no_data)}")

        # Display DataFrames
        st.subheader("Excluded Companies")
        st.dataframe(excluded)

        st.subheader("Retained Companies")
        st.dataframe(retained)

        st.subheader("No Data Companies")
        st.dataframe(no_data)

        # Save to Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            excluded.to_excel(writer, sheet_name="Excluded", index=False)
            retained.to_excel(writer, sheet_name="Retained", index=False)
            no_data.to_excel(writer, sheet_name="No Data", index=False)
        output.seek(0)

        st.download_button(
            "Download Processed File",
            output,
            "all_companies_exclusion.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
