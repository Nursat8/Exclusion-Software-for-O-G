import streamlit as st
import pandas as pd
import io

def load_data(file, sheet_name, header_row):
    return pd.read_excel(file, sheet_name=sheet_name, header=4)  # Adjusted to match row 5 where data starts

def filter_exclusions(upstream_df, midstream_df):
    # Select correct columns using index positions
    upstream_df = upstream_df.iloc[:, [5, 27, 41, 42, 46]]  # Company, AB, AP, AQ, AU
    upstream_df.columns = ["Company", "Fossil Fuel Share of Revenue", "BB Ticker", "ISIN Equity", "LEI"]
    
    midstream_df = midstream_df.iloc[:, [5, 8, 9, 10, 11]]  # Company, I, J, K, L
    midstream_df.columns = [
        "Company",
        "Length of Pipelines under Development",
        "Liquefaction Capacity (Export)",
        "Regasification Capacity (Import)",
        "Total Capacity under Development"
    ]
    
    # Convert Fossil Fuel Share of Revenue to numeric, handling errors
    upstream_df["Fossil Fuel Share of Revenue"] = pd.to_numeric(
        upstream_df["Fossil Fuel Share of Revenue"].astype(str).str.replace('%', ''), errors='coerce'
    ).fillna(0)  # Replace NaN with 0
    
    # Identify exclusion criteria
    upstream_exclusion = upstream_df["Fossil Fuel Share of Revenue"] > 0
    midstream_exclusion = midstream_df.iloc[:, 1:].notna().any(axis=1)  # Check if any midstream column has a value
    
    # Create exclusion reason
    upstream_df["Exclusion Reason"] = ""
    upstream_df.loc[upstream_exclusion, "Exclusion Reason"] = "Upstream - Fossil Fuel Revenue > 0%"
    midstream_df["Exclusion Reason"] = ""
    midstream_df.loc[midstream_exclusion, "Exclusion Reason"] = "Midstream Expansion - Capacity in Development"
    
    # Combine data
    excluded_companies = pd.concat([
        upstream_df.loc[upstream_exclusion, ["Company", "BB Ticker", "ISIN Equity", "LEI", "Fossil Fuel Share of Revenue", "Exclusion Reason"]],
        midstream_df.loc[midstream_exclusion, ["Company", "Exclusion Reason", "Length of Pipelines under Development", "Liquefaction Capacity (Export)", "Regasification Capacity (Import)", "Total Capacity under Development"]]
    ], ignore_index=True)
    
    return excluded_companies

def main():
    st.title("Level 2 Exclusion Filter")
    uploaded_file = st.file_uploader("Upload the Excel file", type=["xlsx"])
    
    if uploaded_file is not None:
        upstream_df = load_data(uploaded_file, sheet_name="Upstream", header_row=4)
        midstream_df = load_data(uploaded_file, sheet_name="Midstream Expansion", header_row=4)
        
        excluded_data = filter_exclusions(upstream_df, midstream_df)
        
        st.subheader("Excluded Companies")
        st.dataframe(excluded_data)
        
        # Save the output as an Excel file
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            excluded_data.to_excel(writer, index=False, sheet_name='Exclusions')
        output.seek(0)
        
        # Provide download option
        st.download_button("Download Exclusion List", output, "excluded_companies.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
