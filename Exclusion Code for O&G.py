import streamlit as st
import pandas as pd
import io
from io import BytesIO

def flatten_multilevel_columns(df):
    """
    Convert multi-level headers into single strings, e.g.:
      ("Company", "Unnamed: 11_level_1") -> "Company Unnamed: 11_level_1"
    """
    df.columns = [
        " ".join(str(level) for level in col).strip()
        for col in df.columns
    ]
    return df

def find_column(df, possible_matches, how="partial", required=True):
    """
    Finds the first column whose name (case-insensitive) partially or exactly
    matches any item in possible_matches.
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

def rename_columns(df):
    """
    1) Flatten multi-level columns into single-level strings.
    2) Skip 2 top data-rows so that row 7 in Excel becomes row 0 in the DF.
    3) Then rename partial matches for 'Company', 'Fossil Fuel Share', etc.
    """

    # A) Flatten multi-level headers
    df = flatten_multilevel_columns(df)

    # B) The user said "names of companies start from row 7",
    #    so we skip the first 2 lines after the header. Adjust if you need to skip more/fewer.
    df = df.iloc[2:].copy()  # skip the next 2 lines
    df.reset_index(drop=True, inplace=True)

    # C) rename partial matches
    rename_map = {
        "Company": ["company"],  # find e.g. "Company Unnamed" etc. then rename to "Company"
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

def filter_all_companies(df):
    """
    - Multi-level columns read from header=[3,4]
    - Flatten columns, skip 2 lines so row 7 is row 0
    - Find & rename partial matches to get "Company" + Upstream/Midstream columns
    - Exclusion logic
    """

    # 1) Flatten + rename
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

    # 3) Convert numeric
    for c in numeric_cols:
        df[c] = (
            df[c].astype(str)
            .str.replace("%", "", regex=True)
            .str.replace(",", "", regex=True)
        )
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    # 4) If "Company" not found, create blank
    if "Company" not in df.columns:
        df["Company"] = None

    # 5) Upstream + Midstream Exclusion
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

    # 7) No Data if numeric=0 & not excluded
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

    excluded_df = df[df["Excluded"]].copy()
    no_data_df = df[no_data_mask].copy()
    retained_df = df[~df["Excluded"] & ~no_data_mask].copy()

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
        if c not in excluded_df.columns:
            excluded_df[c] = None
        if c not in no_data_df.columns:
            no_data_df[c] = None
        if c not in retained_df.columns:
            retained_df[c] = None

    return excluded_df[final_cols], retained_df[final_cols], no_data_df[final_cols]

#########################
# STREAMLIT APP
#########################

def main():
    st.title("All Companies Multi-Level Header Parsing")

    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

    if uploaded_file:
        xls = pd.ExcelFile(uploaded_file)

        if "All Companies" not in xls.sheet_names:
            st.error("No sheet named 'All Companies'.")
            return

        # Parse with multi-level headers from rows 3 & 4 (0-based)
        # So if your Excel has 'Company' in row 4, 'Unnamed...' in row 5, etc.
        df_all = pd.read_excel(
            uploaded_file,
            sheet_name="All Companies",
            header=[3,4]  # 0-based
        )

        excluded, retained, no_data = filter_all_companies(df_all)

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
            "Download Results",
            output,
            "all_companies_exclusion.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
