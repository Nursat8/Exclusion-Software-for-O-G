import streamlit as st
import pandas as pd

def main():
    st.title("Debug 'Company' Column")
    uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"])

    # Try adjusting this if columns look wrong
    header_row = st.number_input(
        "Header row index (0-based)",
        min_value=0,
        max_value=50,
        value=4,
        help="Adjust until the columns match your real header row in Excel."
    )

    if uploaded_file:
        xls = pd.ExcelFile(uploaded_file)
        if "All Companies" not in xls.sheet_names:
            st.error("No 'All Companies' sheet found.")
            return

        df = pd.read_excel(uploaded_file, sheet_name="All Companies", header=header_row)

        st.write("## Columns Found:")
        st.write(list(df.columns))

        st.write("## Preview of top 10 rows:")
        st.dataframe(df.head(10))

        st.write("""
        - If you see "Company" in the columns, great.
        - If you see it spelled differently or not at all, you may need to change the header_row or rename columns.
        """)
    else:
        st.info("Upload a file to debug.")

if __name__ == "__main__":
    main()
