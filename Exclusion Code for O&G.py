import streamlit as st
import pandas as pd
import io
from io import BytesIO

def rename_columns(df):
    """
    Tries partial matching for 'Company' among synonyms, but 
    EXCLUDES any mention of 'Bloomberg' so we do NOT pick up 
    'Company Name in Bloomberg'.
    """
    # Strip column names of trailing spaces
    df.columns = df.columns.str.strip()

    # Define synonyms for the real Company column
    company_synonyms = ["company", "company name"]  # omit "bloomberg"

    # Try to find one that matches
    # We'll do a partial search that excludes "bloomberg"
    found_company_col = None
    for col in df.columns:
        col_lower = col.strip().lower()
        # If "bloomberg" is in col_lower, skip it
        if "bloomberg" in col_lower:
            continue
        for pattern in company_synonyms:
            if pattern in col_lower:
                found_company_col = col
                break
        if found_company_col:
            break

    if found_company_col and found_company_col != "Company":
        df.rename(columns={found_company_col: "Company"}, inplace=True)

    # Now do partial matching for other fields
    partial_map = {
        "Fossil Fuel Share of Revenue": [
            "fossil fuel share", "fossil fuel revenue"
        ],
        "Length of Pipelines under Development": [
            "length of pipelines", "pipeline under dev"
        ],
        "Liquefaction Capacity (Export)": ["liquefaction capacity", "lng export"],
        "Regasification Capacity (Import)": ["regasification capacity", "lng import"],
        "Total Capacity under Development": ["total capacity under development", "total dev capacity"],
        "BB Ticker": ["bb ticker", "bloomberg ticker"],
        "ISIN Equity": ["isin equity", "isin code"],
        "LEI": ["lei"]
    }

    for new_col, patterns in partial_map.items():
        for col in df.columns:
            col_lower = col.strip().lower()
            for pat in patterns:
                if pat in col_lower:
                    df.rename(columns={col: new_col}, inplace=True)
                    break

def filter_all_companies(df):
    # 1) Rename columns (handles 'Company' and partial matches for other fields)
    rename_columns(df)

    # 2) Ensure numeric fields exist or fill with 0
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

    # 3) Ensure 'Company', 'BB Ticker', 'ISIN Equity', 'LEI'
    for c in ["Company", "BB Ticker", "ISIN Equity", "LEI"]:
        if c not in df.columns:
            df[c] = None

    # 4) Convert numeric
    for c in numeric_cols:
        df[c] = (
            df[c].astype(str)
            .str.replace("%", "", regex=True)
            .str.replace(",", "", regex=True)
        )
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    # 5) Exclusion flags
    df["Upstream_Exclusion_Flag"] = df["Fossil Fuel Share of Revenue"] > 0
    df["Midstream_Exclusion_Flag"] = (
        (df["Length of Pipelines under Development"] > 0)
        | (df["Liquefaction Capacity (Export)"] > 0)
        | (df["Regasification Capacity (Import)"] > 0)
        | (df["Total Capacity under Development"] > 0)
    )
    df["Excluded"] = df["Upstream_Exclusion_Flag"] | df["Midstream_Exclusion_Flag"]

    # 6) Build reason
    reasons = []
    for _, row in df.iterrows():
        r = []
        if row["Upstream_Exclusion_Flag"]:
            r.append("Upstream - Fossil Fuel Share > 0%")
        if row["Midstream_Exclusion_Flag"]:
            r.append("Midstream Expansion > 0")
        reasons.append("; ".join(r))
    df["Exclusion Reason"] = reasons

    # 7) No Data
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

    excluded = df[df["Excluded"]].copy()
    no_data = df[no_data_mask].copy()
    retained = df[~df["Excluded"] & ~no_data_mask].copy()

    # 8) Final columns
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
        if c not in excluded.columns:
            excluded[c] = None
        if c not in no_data.columns:
            no_data[c] = None
        if c not in retained.columns:
            retained[c] = None

    return excluded[final_cols], retained[final_cols], no_data[final_cols]

import streamlit as st
import pandas as pd
import io
from io import BytesIO

def main():
    st.title("All Companies Exclusion - Force 'Company' & Exclude 'Bloomberg' Column Name")
    uploaded_file = st.file_uploader("Upload Excel (All Companies sheet)", type=["xlsx"])

    header_row = 4  # Adjust as needed

    if uploaded_file:
        xls = pd.ExcelFile(uploaded_file)
        if "All Companies" not in xls.sheet_names:
            st.error("No sheet called 'All Companies' found.")
            return

        df_all = pd.read_excel(uploaded_file, sheet_name="All Companies", header=header_row)
        excluded, retained, no_data = filter_all_companies(df_all)

        total_count = len(excluded) + len(retained) + len(no_data)

        st.write(f"**Total:** {total_count}")
        st.write(f"**Excluded:** {len(excluded)}")
        st.write(f"**Retained:** {len(retained)}")
        st.write(f"**No Data:** {len(no_data)}")

        st.subheader("Excluded")
        st.dataframe(excluded)

        st.subheader("Retained")
        st.dataframe(retained)

        st.subheader("No Data")
        st.dataframe(no_data)

        # Download
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            excluded.to_excel(writer, sheet_name="Excluded", index=False)
            retained.to_excel(writer, sheet_name="Retained", index=False)
            no_data.to_excel(writer, sheet_name="No Data", index=False)
        output.seek(0)

        st.download_button(
            "Download Exclusion Results",
            output,
            "All_Companies_Exclusion_Results.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
