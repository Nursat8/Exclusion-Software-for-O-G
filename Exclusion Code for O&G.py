import streamlit as st
import pandas as pd

def load_data(file, sheet_name, header_row):
    return pd.read_excel(file, sheet_name=sheet_name, header=header_row)

def filter_exclusions(upstream_df, midstream_df):
    # Define relevant columns
    upstream_col = "Fossil Fuel Share of Revenue"
    midstream_cols = [
        "Length of Pipelines under Development",
        "Liquefaction Capacity (Export)",
        "Regasification Capacity (Import)",
        "Total Capacity under Development"
    ]
    
    # Identify exclusion criteria
    upstream_exclusion = upstream_df[upstream_col].astype(str).str.replace('%', '').astype(float) > 0
    midstream_exclusion = midstream_df[midstream_cols].notna().any(axis=1)
    
    # Create exclusion reason
    upstream_df["Exclusion Reason"] = ""
    upstream_df.loc[upstream_exclusion, "Exclusion Reason"] = "Upstream - Fossil Fuel Revenue > 0%"
    midstream_df["Exclusion Reason"] = ""
    midstream_df.loc[midstream_exclusion, "Exclusion Reason"] = "Midstream Expansion - Capacity in Development"
    
    # Select relevant columns
    relevant_cols = ["Company", "BB Ticker", "ISIN equity", "LEI", "Exclusion Reason"] + midstream_cols + [upstream_col]
    
    # Filter and combine data
    upstream_filtered = upstream_df.loc[upstream_exclusion, relevant_cols]
    midstream_filtered = midstream_df.loc[midstream_exclusion, relevant_cols]
    
    excluded_companies = pd.concat([upstream_filtered, midstream_filtered], ignore_index=True)
    
    return excluded_companies

def main():
    st.title("Level 2 Exclusion Filter")
    uploaded_file = st.file_uploader("Upload the Excel file", type=["xlsx"])
    
    if uploaded_file is not None:
        upstream_df = load_data(uploaded_file, sheet_name="Upstream", header_row=6)
        midstream_df = load_data(uploaded_file, sheet_name="Midstream Expansion", header_row=6)
        
        excluded_data = filter_exclusions(upstream_df, midstream_df)
        
        st.subheader("Excluded Companies")
        st.dataframe(excluded_data)
        
        # Provide download option
        csv = excluded_data.to_csv(index=False).encode('utf-8')
        st.download_button("Download Exclusion List", csv, "excluded_companies.csv", "text/csv")

if __name__ == "__main__":
    main()
