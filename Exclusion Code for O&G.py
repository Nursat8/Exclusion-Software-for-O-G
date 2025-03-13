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

def find_column(df, possible_matches, required=True):
    """Finds the first column matching any item in possible_matches."""
    for col in df.columns:
        col_lower = col.strip().lower()
        for pattern in possible_matches:
            pat_lower = pattern.strip().lower()
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
    Flatten multi-level headers and ensure correct column detection.
    """
    df = flatten_multilevel_columns(df)
    
    # Ensure row 7 in Excel is row 0 in pandas (Shift up by 1 row)
    df = df.iloc[1:].reset_index(drop=True)

    rename_map = {
        "Company": ["company"],  # Ensure we get the correct "Company" column dynamically
        "GOGEL Tab": ["GOGEL Tab"]  # The key column for exclusion
    }

    for new_col, patterns in rename_map.items():
        old_col = find_column(df, patterns, required=False)
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

    # 2) Ensure "Company" and "GOGEL Tab" exist
    if "Company" not in df.columns:
        df["Company"] = None
    if "GOGEL Tab" not in df.columns:
        df["GOGEL Tab"] = ""

    # 3) Exclude companies where "GOGEL Tab" contains "upstream"
    df["Excluded"] = df["GOGEL Tab"].str.contains("upstream", case=False, na=False)

    # 4) Build Exclusion Reason
    df["Exclusion Reason"] = df["Excluded"].apply(lambda x: "Upstream in GOGEL Tab" if x else "")

    # 5) Identify No Data Companies (Empty "GOGEL Tab", Not Excluded)
    def is_no_data(r):
        return r["GOGEL Tab"].strip() == "" and not r["Excluded"]

    no_data_mask = df.apply(is_no_data, axis=1)

    # 6) Split into categories
    excluded_df = df[df["Excluded"]].copy()
    no_data_df = df[no_data_mask].copy()
    retained_df = df[~df["Excluded"] & ~no_data_mask].copy()

    # 7) Keep only required columns
    final_cols = ["Company", "GOGEL Tab", "Exclusion Reason"]
    for c in final_cols:
        for d in [excluded_df, retained_df, no_data_df]:
            if c not in d.columns:
                d[c] = None

    return excluded_df[final_cols], retained_df[final_cols], no_data_df[final_cols]

#########################
# 3) STREAMLIT APP
#########################

def main():
    st.title("All Companies Exclusion Analysis (Excluding Upstream)")

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
        st.write(f"**Excluded Companies (Upstream in GOGEL Tab):** {len(excluded)}")
        st.write(f"**Retained Companies:** {len(retained)}")
        st.write(f"**No Data Companies (Blank GOGEL Tab):** {len(no_data)}")

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
