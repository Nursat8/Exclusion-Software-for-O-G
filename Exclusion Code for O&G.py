import streamlit as st
import pandas as pd
import io
from io import BytesIO

########################################
# 1) CORE EXCLUSION LOGIC
########################################

def filter_all_companies(df):
    """
    Reads from "All Companies" sheet (with 'Company' column) and applies:
      - Upstream Exclusion if Fossil Fuel Share of Revenue > 0
      - Midstream Exclusion if any capacity > 0
    Splits data into Excluded, Retained, No Data.
    Keeps all rows (no dropping blanks).
    """

    # -- 1) Normalize columns: strip trailing spaces, etc. --
    df.columns = df.columns.str.strip()

    # -- 2) Ensure the "Company" column actually exists. --
    #    If your file truly has "Company" in the header row, it should be recognized now.
    if "Company" not in df.columns:
        # We'll create a blank column if not found, so script won't crash
        # But that likely means you used the wrong header row or your file uses a different column name
        df["Company"] = None

    # -- 3) Ensure numeric columns exist or fill with 0 if missing --
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

    # -- 4) Make sure we have Ticker, ISIN, LEI columns, fill if missing --
    for col in ["BB Ticker", "ISIN Equity", "LEI"]:
        if col not in df.columns:
            df[col] = None

    # -- 5) Convert numeric columns: remove '%' and commas, parse as float --
    for col in needed_for_numeric:
        df[col] = (
            df[col].astype(str)
            .str.replace("%", "", regex=True)
            .str.replace(",", "", regex=True)
        )
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # -- 6) Define exclusion flags --
    df["Upstream_Exclusion_Flag"] = df["Fossil Fuel Share of Revenue"] > 0
    df["Midstream_Exclusion_Flag"] = (
        (df["Length of Pipelines under Development"] > 0)
        | (df["Liquefaction Capacity (Export)"] > 0)
        | (df["Regasification Capacity (Import)"] > 0)
        | (df["Total Capacity under Development"] > 0)
    )
    df["Excluded"] = df["Upstream_Exclusion_Flag"] | df["Midstream_Exclusion_Flag"]

    # -- 7) Build Exclusion Reason --
    reason_list = []
    for _, row in df.iterrows():
        reasons = []
        if row["Upstream_Exclusion_Flag"]:
            reasons.append("Upstream - Fossil Fuel Share > 0%")
        if row["Midstream_Exclusion_Flag"]:
            reasons.append("Midstream Expansion > 0")
        reason_list.append("; ".join(reasons))
    df["Exclusion Reason"] = reason_list

    # -- 8) "No Data" => all numeric = 0 and not Excluded --
    def is_no_data(r):
        numeric_zero = (
            (r["Fossil Fuel Share of Revenue"] == 0)
            and (r["Length of Pipelines under Development"] == 0)
            and (r["Liquefaction Capacity (Export)"] == 0)
            and (r["Regasification Capacity (Import)"] == 0)
            and (r["Total Capacity under Development"] == 0)
        )
        return (numeric_zero and not r["Excluded"])

    no_data_mask = df.apply(is_no_data, axis=1)

    # Split
    excluded_companies = df[df["Excluded"]].copy()
    no_data_companies = df[no_data_mask].copy()
    retained_companies = df[~df["Excluded"] & ~no_data_mask].copy()

    # -- 9) Final columns --
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

    # Fill missing columns in each subset
    for c in final_cols:
        if c not in excluded_companies.columns:
            excluded_companies[c] = None
        if c not in no_data_companies.columns:
            no_data_companies[c] = None
        if c not in retained_companies.columns:
            retained_companies[c] = None

    excluded_companies = excluded_companies[final_cols]
    retained_companies = retained_companies[final_cols]
    no_data_companies = no_data_companies[final_cols]

    return excluded_companies, retained_companies, no_data_companies

########################################
# 2) STREAMLIT APP
########################################

def main():
    st.title("All Companies Exclusion Filter – Final Code")

    uploaded_file = st.file_uploader("Upload Excel file (sheet named 'All Companies')", type=["xlsx"])

    # Adjust this row index if your actual column headers are on a different row in Excel
    header_row = 4

    if uploaded_file:
        xls = pd.ExcelFile(uploaded_file)
        if "All Companies" not in xls.sheet_names:
            st.error("No sheet called 'All Companies' found in this file.")
            return

        # Load the single sheet
        df_all = pd.read_excel(uploaded_file, sheet_name="All Companies", header=header_row)

        # Apply the logic
        excluded_data, retained_data, no_data_data = filter_all_companies(df_all)

        # Stats
        total_count = len(excluded_data) + len(retained_data) + len(no_data_data)
        st.subheader("Statistics")
        st.write(f"**Total Companies Processed:** {total_count}")
        st.write(f"**Excluded:** {len(excluded_data)}")
        st.write(f"**Retained:** {len(retained_data)}")
        st.write(f"**No Data:** {len(no_data_data)}")

        # Show in the UI
        st.subheader("Excluded Companies")
        st.dataframe(excluded_data)

        st.subheader("Retained Companies")
        st.dataframe(retained_data)

        st.subheader("No Data Companies")
        st.dataframe(no_data_data)

        # Prepare Excel download
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            excluded_data.to_excel(writer, sheet_name="Excluded", index=False)
            retained_data.to_excel(writer, sheet_name="Retained", index=False)
            no_data_data.to_excel(writer, sheet_name="No Data", index=False)
        output.seek(0)

        st.download_button(
            label="Download Final Exclusion Results",
            data=output,
            file_name="All_Companies_Exclusion_Results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("Please upload an Excel file.")

if __name__ == "__main__":
    main()
