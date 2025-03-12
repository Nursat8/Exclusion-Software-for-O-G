import streamlit as st
import pandas as pd
import io

def load_data(file, sheet_name, header_row):
    return pd.read_excel(file, sheet_name=sheet_name, header=3)  # Adjusted to row 4 based on analysis

def level1_exclusion(df):
    # Select relevant columns using index positions
    df = df.iloc[:, [6, 27, 42, 46]]  # Company, Fossil Fuel Share of Revenue, ISIN, LEI
    df.columns = ["Company", "Fossil Fuel Share of Revenue", "ISIN Equity", "LEI"]
    
    # Convert Fossil Fuel Share of Revenue to numeric
    df["Fossil Fuel Share of Revenue"] = pd.to_numeric(
        df["Fossil Fuel Share of Revenue"].astype(str).str.replace('%', ''), errors='coerce'
    ).fillna(0)  # Replace NaN with 0
    
    # Define Level 1 exclusion criteria (example: exclude if fossil fuel share > 50%)
    level1_criteria = df["Fossil Fuel Share of Revenue"] > 50
    
    # Mark exclusion reason
    df["Exclusion Reason"] = ""
    df.loc[level1_criteria, "Exclusion Reason"] = "Level 1 - Fossil Fuel Share > 50%"
    
    # Separate excluded and retained companies
    excluded_level1 = df[level1_criteria]
    retained_level1 = df[~level1_criteria]  # Passes Level 1, moves to Level 2
    
    return excluded_level1, retained_level1

def level2_exclusion(upstream_df, midstream_df):
    # Select correct columns using index positions
    upstream_df = upstream_df.iloc[:, [6, 27, 42, 46]]  # Company, Fossil Fuel Share of Revenue, ISIN, LEI
    upstream_df.columns = ["Company", "Fossil Fuel Share of Revenue", "ISIN Equity", "LEI"]
    
    midstream_df = midstream_df.iloc[:, [6, 8, 9, 10, 11]]  # Company, I, J, K, L
    midstream_df.columns = [
        "Company",
        "Length of Pipelines under Development",
        "Liquefaction Capacity (Export)",
        "Regasification Capacity (Import)",
        "Total Capacity under Development"
    ]
    
    # Convert Fossil Fuel Share of Revenue to numeric
    upstream_df["Fossil Fuel Share of Revenue"] = pd.to_numeric(
        upstream_df["Fossil Fuel Share of Revenue"].astype(str).str.replace('%', ''), errors='coerce'
    ).fillna(0)  # Replace NaN with 0
    
    # Identify Level 2 exclusion criteria
    upstream_exclusion = upstream_df["Fossil Fuel Share of Revenue"] > 0
    midstream_exclusion = midstream_df.iloc[:, 1:].notna().any(axis=1)  # Check if any midstream column has a value
    
    # Create exclusion reason
    upstream_df["Exclusion Reason"] = ""
    upstream_df.loc[upstream_exclusion, "Exclusion Reason"] = "Level 2 - Upstream Fossil Fuel Revenue > 0%"
    midstream_df["Exclusion Reason"] = ""
    midstream_df.loc[midstream_exclusion, "Exclusion Reason"] = "Level 2 - Midstream Expansion in Development"
    
    # Separate excluded and retained companies
    excluded_level2 = pd.concat([
        upstream_df.loc[upstream_exclusion, :],
        midstream_df.loc[midstream_exclusion, :]
    ], ignore_index=True)
    
    retained_level2 = pd.concat([
        upstream_df.loc[~upstream_exclusion, :],
        midstream_df.loc[~midstream_exclusion, :]
    ], ignore_index=True)
    
    return excluded_level2, retained_level2

def main():
    st.title("Exclusion Filter - Level 1 & Level 2")
    uploaded_file = st.file_uploader("Upload the Excel file", type=["xlsx"])
    
    if uploaded_file is not None:
        upstream_df = load_data(uploaded_file, sheet_name="Upstream", header_row=3)
        midstream_df = load_data(uploaded_file, sheet_name="Midstream Expansion", header_row=3)
        
        # Apply Level 1 exclusion
        excluded_level1, retained_level1 = level1_exclusion(upstream_df)
        
        # Apply Level 2 exclusion to retained Level 1 companies
        excluded_level2, retained_level2 = level2_exclusion(retained_level1, midstream_df)
        
        # Save the output as an Excel file with multiple sheets
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            excluded_level1.to_excel(writer, index=False, sheet_name='Excluded Level 1')
            retained_level1.to_excel(writer, index=False, sheet_name='Retained 1')
            excluded_level2.to_excel(writer, index=False, sheet_name='Excluded Level 2')
            retained_level2.to_excel(writer, index=False, sheet_name='Retained 2')
        output.seek(0)
        
        # Provide download option
        st.download_button("Download Exclusion Report", output, "exclusion_report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()
