import streamlit as st
import pandas as pd

def main():
    st.title("Debug the 'Company' Column in 'All Companies' Sheet")
    
    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
    
    header_row = st.number_input(
        "Header row index (0-based)", 
        min_value=0, 
        max_value=50, 
        value=4,
        help="Try different values until you see the correct column headers in the preview."
    )
    
    if uploaded_file:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        if "All Companies" not in sheet_names:
            st.error("No 'All Companies' sheet found in this file.")
            return
        
        # Load with the chosen header row
        df_all = pd.read_excel(
            uploaded_file, 
            sheet_name="All Companies", 
            header=header_row
        )
        
        st.subheader("Columns Found")
        st.write(list(df_all.columns))
        
        st.subheader("Data Preview (top 20 rows)")
        st.dataframe(df_all.head(20))
        
        # If there's a column that visually looks like "Company" or "Company ",
        # you'll see its exact name in 'Columns Found'.
        
        # Next steps:
        # 1) If the column name is exactly "Company", you're good.
        # 2) If you see "Company " or something else, adjust your rename/matching code.
        # 3) If columns are all messed up, try changing header_row above.
        
    else:
        st.info("Please upload a file to debug.")

if __name__ == "__main__":
    main()
