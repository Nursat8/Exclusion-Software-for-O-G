import streamlit as st
import pandas as pd
import io

def load_data(file, sheet_name, header_row):
    """Load Excel data starting from the given header row (0-based)."""
    return pd.read_excel(file, sheet_name=sheet_name, header=header_row)

def filter_exclusions_and_retained(upstream_df, midstream_df):
    """
    1) Rename/select columns of interest.
    2) Group by 'Company' and identify if it meets exclusion criteria
       in Upstream or Midstream.
    3) Combine the data into one DataFrame with columns indicating
       whether the company is excluded by Upstream or Midstream criteria.
    4) Create Exclusion Reason and split into 'Excluded' vs 'Retained' vs 'No Data'.
    """

    # --- 1. Rename columns and select relevant ones ---

    # Upstream: columns by index -> [5, 27, 41, 42, 46]
    upstream_subset = upstream_df.iloc[:, [5, 27, 41, 42, 46]].copy()
    upstream_subset.columns = [
        "Company",                        # col index 5
        "Fossil Fuel Share of Revenue",   # col index 27
        "BB Ticker",                      # col index 41
        "ISIN Equity",                    # col index 42
        "LEI",                            # col index 46
    ]

    # Convert fossil-fuel share to numeric (remove '%' if present)
    upstream_subset["Fossil Fuel Share of Revenue"] = (
        upstream_subset["Fossil Fuel Share of Revenue"]
        .astype(str)
        .str.replace('%', '', regex=True)
    )
    upstream_subset["Fossil Fuel Share of Revenue"] = pd.to_numeric(
        upstream_subset["Fossil Fuel Share of Revenue"],
        errors='coerce'
    ).fillna(0)

    # Midstream: columns by index -> [5, 8, 9, 10, 11]
    midstream_subset = midstream_df.iloc[:, [5, 8, 9, 10, 11]].copy()
    midstream_subset.columns = [
        "Company",                              # col index 5
        "Length of Pipelines under Development",# col index 8
        "Liquefaction Capacity (Export)",       # col index 9
        "Regasification Capacity (Import)",     # col index 10
        "Total Capacity under Development",     # col index 11
    ]

    # Convert all midstream capacity columns to numeric
    numeric_cols = [
        "Length of Pipelines under Development",
        "Liquefaction Capacity (Export)",
        "Regasification Capacity (Import)",
        "Total Capacity under Development"
    ]
    for col in numeric_cols:
        midstream_subset[col] = pd.to_numeric(
            midstream_subset[col], errors='coerce'
        ).fillna(0)

    # --- 2. Determine Upstream/Midstream flags at the "Company" level ---

    def combine_identifiers(series):
        """Collect unique non-null items into a comma-separated string."""
        unique_vals = series.dropna().unique().tolist()
        return ", ".join(map(str, unique_vals)) if unique_vals else ""

    # Upstream criterion: any row for that company has Fossil Fuel Revenue > 0
    upstream_grouped = (
        upstream_subset
        .groupby("Company", dropna=False)  # keep all company names
        .agg({
            "Fossil Fuel Share of Revenue": lambda x: (x > 0).any(),  # bool
            "BB Ticker": combine_identifiers,
            "ISIN Equity": combine_identifiers,
            "LEI": combine_identifiers
        })
        .reset_index()
    )
    upstream_grouped.rename(
        columns={"Fossil Fuel Share of Revenue": "Upstream_Exclusion_Flag"},
        inplace=True
    )

    # Midstream criterion: any pipeline or capacity column is > 0
    def has_midstream_expansion(row):
        return (
            (row["Length of Pipelines under Development"] > 0)
            or (row["Liquefaction Capacity (Export)"] > 0)
            or (row["Regasification Capacity (Import)"] > 0)
            or (row["Total Capacity under Development"] > 0)
        )

    midstream_grouped = (
        midstream_subset
        .groupby("Company", dropna=False)
        .agg({
            "Length of Pipelines under Development": "max",
            "Liquefaction Capacity (Export)": "max",
            "Regasification Capacity (Import)": "max",
            "Total Capacity under Development": "max"
        })
        .reset_index()
    )
    midstream_grouped["Midstream_Exclusion_Flag"] = midstream_grouped.apply(
        has_midstream_expansion, axis=1
    )

    # --- 3. Combine (merge) upstream+midstream groupings by company ---

    combined = pd.merge(
        upstream_grouped,
        midstream_grouped,
        on="Company",
        how="outer"  # full outer join so we don't lose any companies
    )

    # If a company was not in Upstream or Midstream, fill missing booleans with False
    combined["Upstream_Exclusion_Flag"] = combined["Upstream_Exclusion_Flag"].fillna(False)
    combined["Midstream_Exclusion_Flag"] = combined["Midstream_Exclusion_Flag"].fillna(False)

    # --- 4. Determine final exclusion and reason ---

    # Exclude if Upstream_Exclusion_Flag == True OR Midstream_Exclusion_Flag == True
    combined["Excluded"] = (
        combined["Upstream_Exclusion_Flag"] | combined["Midstream_Exclusion_Flag"]
    )

    # Build up an exclusion reason text
    reasons = []
    for _, row in combined.iterrows():
        r = []
        if row["Upstream_Exclusion_Flag"]:
            r.append("Upstream - Fossil Fuel Revenue > 0%")
        if row["Midstream_Exclusion_Flag"]:
            r.append("Midstream Expansion - Capacity > 0")
        reasons.append("; ".join(r))
    combined["Exclusion Reason"] = reasons

    # --- 5. Split into Excluded / Retained / No Data ---
    # Define "No Data" = not excluded, plus no tickers, empty LEI, and all capacities are 0.
    def is_empty_string_or_nan(val):
        return pd.isna(val) or str(val).strip() == ""

    no_data_cond = (
        (~combined["Excluded"])  # must not be excluded
        & combined["BB Ticker"].apply(is_empty_string_or_nan)
        & combined["ISIN Equity"].apply(is_empty_string_or_nan)
        & combined["LEI"].apply(is_empty_string_or_nan)
        & (combined["Length of Pipelines under Development"] == 0)
        & (combined["Liquefaction Capacity (Export)"] == 0)
        & (combined["Regasification Capacity (Import)"] == 0)
        & (combined["Total Capacity under Development"] == 0)
    )

    # Subset data
    no_data_companies = combined[no_data_cond].copy()
    excluded_companies = combined[combined["Excluded"]].copy()
    retained_companies = combined[~combined["Excluded"] & ~no_data_cond].copy()

    return excluded_companies, retained_companies, no_data_companies

def main():
    st.title("Level 2 Exclusion Filter")
    uploaded_file = st.file_uploader("Upload the Excel file", type=["xlsx"])

    if uploaded_file is not None:
        # Load Upstream / Midstream with correct header row (row 4 => 0-based indexing)
        upstream_df = load_data(uploaded_file, sheet_name="Upstream", header_row=4)
        midstream_df = load_data(uploaded_file, sheet_name="Midstream Expansion", header_row=4)

        # Get excluded vs retained vs no data
        excluded_data, retained_data, no_data_data = filter_exclusions_and_retained(
            upstream_df, midstream_df
        )

        # --- Calculate basic statistics ---
        excluded_count = len(excluded_data)
        retained_count = len(retained_data)
        no_data_count = len(no_data_data)
        total_count = excluded_count + retained_count + no_data_count

        st.markdown("### Statistics")
        st.write(f"**Total companies:** {total_count}")
        st.write(f"**Excluded:** {excluded_count}")
        st.write(f"**Retained:** {retained_count}")
        st.write(f"**No Data:** {no_data_count}")

        st.subheader("Excluded Companies")
        st.dataframe(excluded_data)

        st.subheader("Retained Companies")
        st.dataframe(retained_data)

        st.subheader("No Data Companies")
        st.dataframe(no_data_data)

        # Save the output as an Excel file with 3 sheets
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            excluded_data.to_excel(writer, index=False, sheet_name='Exclusions')
            retained_data.to_excel(writer, index=False, sheet_name='Retained')
            no_data_data.to_excel(writer, index=False, sheet_name='No Data')
        output.seek(0)

        # Provide download option
        st.download_button(
            "Download Exclusion & Retention & NoData List",
            output,
            "O&G companies Level 2 Exclusion.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
