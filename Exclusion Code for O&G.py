import streamlit as st
import pandas as pd

def load_data(file):
    return pd.read_excel(file)

def filter_exclusions(df):
    # Define relevant columns
    upstream_col = "Fossil Fuel Share of Revenue"
    midstream_cols = [
        "Length of Pipelines under Development",
        "Liquefaction Capacity (Export)",
        "Regasification Capacity (Import)",
        "Total Capacity under Development"
    ]
    
    # Apply filters
    upstream_exclusion = df[upstream_col] > 0
    midstream_exclusion = df[midstream_cols].notna().any(axis=1)
    
    # Create exclusion reason
    df["Exclusion Reason"] = ""
    df.loc[upstream_exclusion, "Exclusion Reason"] = "Upstream - Fossil Fuel Revenue > 0%"
    df.loc[midstream_exclusion, "Exclusion Reason"] += "Midstream Expansion - Capacity in Development"
    
    # Filter excluded companies
    exclusion_criteria = upstream_exclusion | midstream_exclusion
    excluded_companies = df.loc[exclusion_criteria, [
        "Name of Company", "BB Ticker", "ISIN equity", "LEI", "Exclusion Reason"
    ] + midstream_cols + [upstream_col]]
    
    return excluded_companies

def main():
    st.title("Level 2 Exclusion Filter")
    
    uploaded_file = st.file_uploader("Upload the Excel file", type=["xlsx"])
    
    if uploaded_file is not None:
        df = load_data(uploaded_file)
        excluded_data = filter_exclusions(df)
        
        st.subheader("Excluded Companies")
        st.dataframe(excluded_data)
        
        # Provide download option
        csv = excluded_data.to_csv(index=False).encode('utf-8')
        st.download_button("Download Exclusion List", csv, "excluded_companies.csv", "text/csv")

if __name__ == "__main__":
    main()
