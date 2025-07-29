import re
import pandas as pd
import numpy as np
from io import BytesIO
import streamlit as st

# 🔹 Helper Functions 🔹
# 🔹 This function removes duplicate column names in a table — and keeps only the first copy of each name. 🔹
# 🔹 This function is helpful when your Excel file has multiple sheets and some of them have columns with the same name repeated. It cleans that up by keeping just one version of each column name. 🔹
def ensure_unique_columns(df):
    return df.loc[:, ~df.columns.duplicated()].copy()
    
# 🔹 If your Excel sheet uses two rows for column names (Multindex columns), this function joins them into one clean name so your table is easier to use 🔹    
def flatten_multilevel_columns(df):
    df.columns = [
        " ".join(str(l).strip() for l in col).strip()
        for col in df.columns
    ]
    return df

# 🔹Searches column headers using any of three modes: exact (must match exactly), partial (look for the pattern inside column names), regex (use advanced matching (like wildcards)). Normalises spaces, case, and line-breaks before matching. Raises ValueError when nothing found 🔹 
def find_column(df, patterns, how="partial", required=True):
    # 🔹 This creates a cleaned-up version of all the column names in the table by removing spaces and breaks, making everything in lowercase, and replacing multiple spaces with just one 🔹
    norm_map = {
        col: re.sub(r"\s+", " ", col.strip().lower().replace("\n", " "))
        for col in df.columns
    }
     # 🔹 This does the same cleanup for the words you're looking for. 🔹
    pats = [
        re.sub(r"\s+", " ", p.strip().lower().replace("\n", " "))
        for p in patterns
    ]
    # 🔹 If any cleaned column name matches the cleaned pattern, return that column name. 🔹
    # exact
    for pat in pats:
        for col, norm in norm_map.items():
            if norm == pat:
                return col
    # partial
    if how == "partial":
        for col, norm in norm_map.items():
            for pat in pats:
                if pat in norm:
                    return col
    # regex
    if how == "regex":
        for pattern in patterns:
            for col in df.columns:
                if re.search(pattern, col, flags=re.IGNORECASE):
                    return col
    if required:
        raise ValueError(f"Could not find a required column among {patterns}\nAvailable: {list(df.columns)}")
    return None

# 🔹 It renames column headers in your table so that they all follow a clean, standard name — even if the original names in the Excel file are messy or inconsistent. Takes names from "rename_map" table (presented later in the code) 🔹 
def rename_columns(df, rename_map):
    for new, pats in rename_map.items():
        old = find_column(df, pats, how="partial", required=False)
        if old and old != new:
            df.rename(columns={old: new}, inplace=True)
    return df

# 🔹 Removes hard-space (\u00A0) characters. Strips any case-insensitive " Equity" suffix. Returns a copy so original df is untouched.🔹 
def remove_equity_from_bb_ticker(df):
    df = df.copy()
    if "BB Ticker" in df.columns:
        df["BB Ticker"] = (
            df["BB Ticker"]
              .astype(str)
              .str.replace(r"\u00A0", " ", regex=True)
              .str.replace(r"(?i)\s*Equity\s*", "", regex=True)
              .str.strip()
        )
    return df

# 🔹🔹🔹 Level 1 Exclusion 🔹🔹🔹
# 🔹 It opens an Excel file, reads a sheet called “All Companies”, cleans up the column names, and removes any company listed as a “Parent Company". 🔹
# 🔹 Reads data from rows 4 and 5 (0-indexed) from a two-level column index. It is needed as a column name located not in the first row. Data clearingand ignores "parent company" column🔹
def filter_companies_by_revenue(uploaded_file, sector_exclusions, total_thresholds):
    xls = pd.ExcelFile(uploaded_file)
    df = xls.parse("All Companies", header=[3,4])
    df.columns = [" ".join(map(str,c)).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.str.lower().str.startswith("parent company")]
    df = remove_equity_from_bb_ticker(df)

    # 🔹 Standardized names for integrity. Renames inconsistent column names using the function "rename_columns"🔹
    rename_map = {
        "Company": ["company name","company"],
        "BB Ticker": ["bb ticker"],
        "ISIN equity": ["isin equity"],
        "LEI": ["lei"],
        "Hydrocarbons Production (%)": ["hydrocarbons production"],
        "Fracking Revenue": ["fracking"],
        "Tar Sand Revenue": ["tar sands"],
        "Coalbed Methane Revenue": ["coalbed methane"],
        "Extra Heavy Oil Revenue": ["extra heavy oil"],
        "Ultra Deepwater Revenue": ["ultra deepwater"],
        "Arctic Revenue": ["arctic"],
        "Unconventional Production Revenue": ["unconventional production"]
    }
    df = rename_columns(df, rename_map)
   
    # 🔹 It checks if any important column is missing from the Excel table, and if it is, it adds that column anyway — but fills it with empty values (NaN)🔹
    needed = list(rename_map.keys())
    for c in needed:
        if c not in df.columns:
            df[c] = np.nan

    # 🔹 It checks which companies are completely missing revenue data, and separates them from the rest. Checks for revenue data to ignore columns with company names and tickers🔹
    revenue_cols = needed[4:]
    no_data = df[df[revenue_cols].isnull().all(axis=1)].copy()
    df = df.dropna(subset=revenue_cols, how="all")
 
    # 🔹 European comma → US dot, percent sign removed, cast to float data, and corrupt values cleaning 🔹
    for c in revenue_cols:
        df[c] = (
            df[c]
              .astype(str)
              .str.replace("%","",regex=True)
              .str.replace(",","",regex=True)
        )
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
   
    # 🔹 This part adds up revenue columns for each company, based on what the user selected in Streamlit (Custom Totals). 🔹
    # 🔹 Line 2 (which is "secs = ...") builds a list of valid sector columns from the user's selection 🔹
    # 🔹 Line 3 creates a new column in the table named after key, like "Custom Total 1".🔹 🔹
    for key,info in total_thresholds.items():
        secs = [s for s in info["sectors"] if s in df.columns]
        df[key] = df[secs].sum(axis=1) if secs else 0.0
        
    # 🔹 This checks each company, one by one, to see if it should be excluded based on user-defined sector thresholds🔹
    # 🔹 A bit of explanation of the code: in the for loop, "r" represents one row (one company), "_" is just a throwaway variable for the row index (it means the code would ignore the index and use only the row's content)🔹
    # 🔹 We create an empty list called "parts" to store all the reasons that apply to this one company. "sector" might be something like "Fracking Revenue", "flag" is True or False (whether the user checked the box to exclude), "thr" is the threshold string the user typed (like "10")
    reasons = []
    for _,r in df.iterrows():
        parts = []
        for sector,(flag,thr) in sector_exclusions.items():
            if flag and thr.strip():
                try:
                    if r[sector] > float(thr)/100:
                        parts.append(f"{sector} > {thr}%")
                except:
                    pass
    
    # 🔹 It checks whether the company exceeds any custom total threshold (like “Custom Total 1 > 15%”), and if so, adds a reason explaining that.🔹
    # 🔹 A bit of details on how the "info" dictionary works: info = { "sectors": ["Fracking Revenue", "Arctic Revenue"], "threshold": "10"}. The "section" part is a combination of sectors selected by the user in Custom Total, and the "threshold" is a value set by the user.
        for key,info in total_thresholds.items():
            t = info.get("threshold","").strip()
            if t:
                try:
                    if r[key] > float(t)/100:
                        parts.append(f"{key} > {t}%")
                except:
                    pass
        reasons.append("; ".join(parts))
    df["Exclusion Reason"] = reasons
   
    # 🔹 It splits the companies into two groups: Retained and Excluded 🔹
    excluded = df[df["Exclusion Reason"]!=""].copy()
    retained = df[df["Exclusion Reason"]==""].copy()

    # 🔹 If there's only one custom total (called "Custom Total 1"), it renames that column to a friendlier name: "Custom Total Revenue"🔹
    if "Custom Total 1" in df.columns:
        for d in (excluded, retained, no_data):
            d.rename(columns={"Custom Total 1":"Custom Total Revenue"}, inplace=True)
  
    # 🔹 This section repairs company names in the final output. It fixes cases where the name is missing or is just a "." Any blank values in the "Company" column are replaced with empty text "". Ensures all entries are strings (text), even if they were originally numbers or empty.🔹
    # 🔹 Makes sure the column headers are turned into simple strings (in case it's a multi-level header like before).🔹
    # 🔹 Removes the column if its name starts with "parent company" — just being safe. 🔹
    # 🔹 Wherever a company name is missing (even if it was "."), we fill it in using the clean names from the original Excel sheet (raw["Company"]). 🔹
    raw = xls.parse("All Companies", header=[3,4]).iloc[:,[6]]
    raw = flatten_multilevel_columns(raw)
    raw = raw.loc[:, ~raw.columns.str.lower().str.startswith("parent company")]
    raw.columns = ["Company"]
    raw["Company"] = raw["Company"].fillna("").astype(str)
    for d in (excluded, retained, no_data):
        d.reset_index(drop=True, inplace=True)
        d["Company"] = d["Company"].replace(".", np.nan)
        d["Company"].fillna(raw["Company"], inplace=True)

    return excluded, retained, no_data

# 🔹🔹🔹 Level-2 — Upstream filter 🔹🔹🔹
# 🔹 It prepares the "Upstream" sheet of the Excel file by flattening the column names and Removing any columns that start with "Parent Company"🔹
def filter_upstream_companies(df):
    df = flatten_multilevel_columns(df)
    df = df.loc[:, ~df.columns.str.lower().str.startswith("parent company")]

    # 🔹 This finds and stores the correct column names from the Excel sheet, even if they don’t match exactly. Type of searching can be adjusted in "how =" 🔹
    comp_col      = find_column(df, ["company"], how="partial", required=True)
    res_col       = find_column(df, ["resources under development and field evaluation"],
                                how="partial", required=True)
    capex_avg_col = find_column(df, ["exploration capex 3-year average"],
                                how="partial", required=True)
    short_col     = find_column(df, ["short-term expansion ≥20 mmboe"],
                                how="partial", required=True)
    capex10_col   = find_column(df, ["exploration capex ≥10 musd"],
                                how="partial", required=True)

    # 🔹 It renames the columns in the DataFrame to a standard set of names, no matter what the original Excel file called them. 🔹
    df = df.rename(columns={
        comp_col     : "Company",
        res_col      : "Resources under Development and Field Evaluation",
        capex_avg_col: "Exploration CAPEX 3-year average",
        short_col    : "Short-Term Expansion ≥20 mmboe",
        capex10_col  : "Exploration CAPEX ≥10 MUSD",
    })

    # 🔹 It takes two columns (which should have numbers), cleans them up, and converts them to proper numbers, so we can safely do comparisons and math. 🔹
    num_cols = [
        "Resources under Development and Field Evaluation",
        "Exploration CAPEX 3-year average",
    ]
    for c in num_cols:
        df[c] = pd.to_numeric(
            df[c].astype(str)               # 🔹 Ensures even numeric or null cells are treated as strings.
                 .str.replace(",", "", regex=True)   # 🔹 Removes commas
                 .str.replace(r"[^\d.\-]", "", regex=True),   # 🔹 Removes everything except digits, decimal points, and minus signs.
            errors="coerce"
        ).fillna(0)


    # 🔹 Checks whether the company has any resources under development, invested any CAPEX over the past 3 years, short-term expansion exceeds 20 MMBOE, larger exploration projects with CAPEX ≥ $10 million, Exclude if any condition is true 🔹
    df["F2_Res"] = df["Resources under Development and Field Evaluation"] > 0
    df["F2_Avg"] = df["Exploration CAPEX 3-year average"] > 0
    df["F2_ST"]  = df["Short-Term Expansion ≥20 mmboe"].astype(str).str.lower().eq("yes")
    df["F2_10M"] = df["Exploration CAPEX ≥10 MUSD"].astype(str).str.lower().eq("yes")
    df["Excluded"] = df[["F2_Res","F2_Avg","F2_ST","F2_10M"]].any(axis=1)

    # 🔹 For each company (row), it builds a text summary of the reasons why that company was excluded — based on which conditions were true.🔹
    df["Exclusion Reason"] = df.apply(
        lambda r: "; ".join(p for p in (
            "Resources under development and field evaluation > 0" if r["F2_Res"] else None,
            "3-yr CAPEX avg > 0" if r["F2_Avg"] else None,
            "Short-Term Expansion = Yes"   if r["F2_ST"]  else None,
            "CAPEX ≥10 MUSD = Yes"         if r["F2_10M"] else None,
        ) if p),
        axis=1
    )

    # 🔹 This part splits the companies into two groups:: excluded and retained companies🔹
    exc = df[df["Excluded"]].copy()
    ret = df[~df["Excluded"]].copy()
    return exc[[
        "Company",
        "Resources under Development and Field Evaluation",
        "Exploration CAPEX 3-year average",
        "Short-Term Expansion ≥20 mmboe",
        "Exploration CAPEX ≥10 MUSD",
        "Exclusion Reason"
    ]], ret[[
        "Company",
        "Resources under Development and Field Evaluation",
        "Exploration CAPEX 3-year average",
        "Short-Term Expansion ≥20 mmboe",
        "Exploration CAPEX ≥10 MUSD",
        "Exclusion Reason"
    ]]

# 🔹 Excel Helpers 🔹
# 🔹 This function prepares and exports the Level 1 results into an Excel file with 3 separate sheets: Excluded Level 1, Retained Level 1, L1 No Data 🔹
    cols = [
        "Company","BB Ticker","ISIN equity","LEI",
        "Hydrocarbons Production (%)","Fracking Revenue","Tar Sand Revenue",
        "Coalbed Methane Revenue","Extra Heavy Oil Revenue","Ultra Deepwater Revenue",
        "Arctic Revenue","Unconventional Production Revenue","Exclusion Reason","Custom Total Revenue"
    ]
    exc     = remove_equity_from_bb_ticker(exc)
    ret     = remove_equity_from_bb_ticker(ret)
    no_data = remove_equity_from_bb_ticker(no_data)
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        exc.reindex(columns=cols).to_excel(w, "Excluded Level 1", index=False)
        ret.reindex(columns=cols).to_excel(w, "Retained Level 1", index=False)
        no_data.reindex(columns=cols).to_excel(w, "L1 No Data", index=False)
    buf.seek(0)
    return buf

# 🔹 This function seperates in 7 different sheets: All excluded, Excluded level 1, Excluded level 2, Retained level 1, Retained level 2 midstream filter, Excluded by upstream filter, Retained by upstream filter  🔹
def to_excel_l2(all_exc, exc1, exc2, ret1, ret2, exc_up, ret_up):
    cols = [
        # identity / Level-1 data
        "Company","BB Ticker","ISIN equity","LEI",
        "Hydrocarbons Production (%)","Fracking Revenue","Tar Sand Revenue",
        "Coalbed Methane Revenue","Extra Heavy Oil Revenue","Ultra Deepwater Revenue",
        "Arctic Revenue","Unconventional Production Revenue","Custom Total Revenue",
        # midstream
        "GOGEL Tab","Length of Pipelines under Development","Liquefaction Capacity (Export)",
        "Regasification Capacity (Import)","Total Capacity under Development",
        # upstream
        "Resources under Development and Field Evaluation","Exploration CAPEX 3-year average",
        "Short-Term Expansion ≥20 mmboe","Exploration CAPEX ≥10 MUSD",
        # …and finally the reason string
        "Exclusion Reason"
    ]
    cols = list(dict.fromkeys(cols))   # keep order, drop accidental dups

    # 🔹 Remove duplicates while preserving order 🔹
    cols = list(dict.fromkeys(cols))
    for df in (all_exc, exc1, exc2, ret1, ret2, exc_up, ret_up):
        df.update(remove_equity_from_bb_ticker(df))
    
    # 🔹 This creates a temporary "file" in memory. It acts like a blank Excel file, but it's stored in RAM (not saved on your computer yet). 🔹
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        all_exc.reindex(columns=cols).to_excel(w, "All Excluded Companies", index=False)
        exc1   .reindex(columns=cols).to_excel(w, "Excluded Level 1",      index=False)
        exc2   .reindex(columns=cols).to_excel(w, "Midstream Excluded",     index=False)
        ret1   .reindex(columns=cols).to_excel(w, "Retained Level 1",       index=False)
        ret2   .reindex(columns=cols).to_excel(w, "Midstream Retained",     index=False)
        exc_up .reindex(columns=cols).to_excel(w, "Upstream Excluded",      index=False)
        ret_up .reindex(columns=cols).to_excel(w, "Upstream Retained",      index=False)
    buf.seek(0)
    return buf


# 🔹🔹🔹 Streamlit App (UI)🔹🔹🔹

def main():
    st.title("Level 1 & Level 2 Exclusion Filter for O&G")
    uploaded = st.file_uploader("Upload Excel file", type=["xlsx"])

    # 🔹 Level 1 sidebar 🔹
    st.sidebar.header("Level 1 Settings")
    sectors = [
        "Hydrocarbons Production (%)","Fracking Revenue","Tar Sand Revenue",
        "Coalbed Methane Revenue","Extra Heavy Oil Revenue","Ultra Deepwater Revenue",
        "Arctic Revenue","Unconventional Production Revenue",
    ]
    sector_excs = {}
    for s in sectors:
        chk = st.sidebar.checkbox(f"Exclude {s}")
        thr = ""
        if chk:
            thr = st.sidebar.text_input(f"{s} Threshold (%)")
        sector_excs[s] = (chk, thr)

    st.sidebar.header("Custom Total Thresholds")
    total_thresholds = {}
    n = st.sidebar.number_input("How many totals?", 1, 5, 1)
    for i in range(n):
        sels = st.sidebar.multiselect(f"Sectors for Total {i+1}", sectors, key=f"sel{i}")
        thr  = st.sidebar.text_input(f"Threshold {i+1} (%)", key=f"thr{i}")
        if sels and thr:
            total_thresholds[f"Custom Total {i+1}"] = {"sectors":sels,"threshold":thr}

    if st.sidebar.button("Run Level 1 Exclusion"):
        if not uploaded:
            st.warning("Please upload a file first.")
        else:
            exc1, ret1, no1 = filter_companies_by_revenue(uploaded, sector_excs, total_thresholds)
            st.success("Level 1 complete")
            st.download_button(
                "Download Level 1 Results",
                data=to_excel_l1(exc1, ret1, no1),
                file_name="O&G_Level1_Exclusion.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    st.markdown("---")
    st.header("Level 2 Exclusion")
    st.write("Applies All-Companies + Upstream filters, merges duplicates, and fills in all data.")

    if st.button("Run Level 2 Exclusion"):
        df_all = pd.read_excel(uploaded, "All Companies", header=[3, 4])
        df_all = ensure_unique_columns(df_all)      #  <-- after reading
        exc_all, ret_all = filter_all_companies(df_all)
       # 🔹 Upstream L2
        df_up = pd.read_excel(uploaded, "Upstream", header=[3, 4])
        df_up = ensure_unique_columns(df_up)        #  <-- after reading
        exc_up, ret_up = filter_upstream_companies(df_up)
        if not uploaded:
            st.warning("Please upload a file first.")
            return

        # 🔹 Rerun L1 to get full df_l1_all 🔹 
        exc1, ret1, no1 = filter_companies_by_revenue(uploaded, sector_excs, total_thresholds)
        df_l1_all = pd.concat([exc1, ret1, no1], ignore_index=True)
        df_l1_all = ensure_unique_columns(df_l1_all)

        # 🔹 after you read the two Level-2 source sheets 🔹
        df_all = ensure_unique_columns(df_all)
        df_up  = ensure_unique_columns(df_up)

        # 🔹 after you build exc_all / ret_all, exc_up / ret_up 🔹
        exc_all = ensure_unique_columns(exc_all)
        exc_up  = ensure_unique_columns(exc_up)

        # 🔹 All-Companies L2 🔹
        df_all = pd.read_excel(uploaded, "All Companies", header=[3,4])
        exc_all, ret_all = filter_all_companies(df_all)

        # 🔹 Upstream L2 🔹
        df_up = pd.read_excel(uploaded, "Upstream", header=[3,4])
        exc_up, ret_up = filter_upstream_companies(df_up)

        # 🔹 Build All Excluded Companies union 🔹
        union = pd.concat([
            exc1[["Company"]],
            exc_all[["Company"]],
            exc_up[["Company"]]
        ]).drop_duplicates()

        # 🔹 Merge in Level 1 Reason 🔹 
        df_l1_meta = df_l1_all[["Company","Exclusion Reason"]].rename(columns={"Exclusion Reason":"L1_Reason"})
        union = union.merge(df_l1_meta, on="Company", how="left")

        # 🔹 Merge in Level 2 All-Companies Reason 🔹
        union = union.merge(
            exc_all[["Company","Exclusion Reason"]].rename(columns={"Exclusion Reason":"L2_Reason_AC"}),
            on="Company", how="left"
        )
        # 🔹 Merge in Level 2 Upstream Reason 🔹
        union = union.merge(
            exc_up[["Company","Exclusion Reason"]].rename(columns={"Exclusion Reason":"L2_Reason_UP"}),
            on="Company", how="left"
        )

        # 🔹 Combine all three reasons into final Exclusion Reason 🔹
        union["Exclusion Reason"] = (
            union[["L1_Reason","L2_Reason_AC","L2_Reason_UP"]]
              .fillna("")
              .agg("; ".join, axis=1)
              .str.replace(r"(; )+", "; ")
              .str.strip("; ")
        )
        union.drop(columns=["L1_Reason","L2_Reason_AC","L2_Reason_UP"], inplace=True)
        meta_cols = df_l1_all.columns.difference(
            ["Company", "Exclusion Reason"]
        ).tolist()


        union = (
            union
                .merge(df_l1_all[["Company", *meta_cols]], on="Company", how="left")
                .merge(exc_all  .drop(columns=["Exclusion Reason"]), on="Company", how="left")
                .merge(exc_up   .drop(columns=["Exclusion Reason"]), on="Company", how="left")
        )
        union = union.merge(
            df_l1_all[["Company", "BB Ticker", "ISIN equity", "LEI"]],
            on="Company", how="left", suffixes=("", "_y")
        )
        union = union.drop(columns=[c for c in union.columns if c.endswith("_y")])

        # 🔹 Retained Level 2 🔹
        all_names = set(df_l1_all["Company"])
        exc2_names = set(union["Company"])
        ret2 = pd.DataFrame({"Company":[c for c in all_names if c not in exc2_names]})
        ret2 = ret2.merge(df_l1_all, on="Company", how="left")

        # 🔹 Upstream full merge 🔹 
        exc_up_full = (
            exc_up
            .merge(df_l1_all.drop(columns=["Exclusion Reason"]),
            on="Company", how="left")
        )
        ret_up_full = (
            ret_up
                .merge(df_l1_all.drop(columns=["Exclusion Reason"]),
                    on="Company", how="left")
        )

        
        buf = to_excel_l2(
            all_exc=union,
            exc1=exc1,
            exc2=(
                # Build Excluded Level 2 sheet properly with its own L2 reasons
                df_l1_all
                  .merge(exc_all[["Company","Exclusion Reason"]]
                           .rename(columns={"Exclusion Reason":"L2_Reason"}), 
                         on="Company", how="inner")
                  .assign(**{"Exclusion Reason": lambda d: d["L2_Reason"]})
                  .drop(columns=["L2_Reason"])
            ),
            ret1=ret1,
            ret2=ret2,
            exc_up=exc_up_full,
            ret_up=ret_up_full
        )
        st.success("Level 2 complete")
        st.download_button(
            "Download Combined Level 1 & 2 Results",
            data=buf,
            file_name="O&G_Level1_Level2_Exclusion.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


# 🔹 Level-2 — Midstream / “All-Companies” filter 🔹
def filter_all_companies(df: pd.DataFrame):
    """
    🔹 Implements Level-2 ‘mid-stream’ screen on the **All Companies** sheet.
    Returns (excluded_df, retained_df) with the canonical columns
    and an ‘Exclusion Reason’ column. 🔹
    """
    # 1. tidy columns ---------------------------------------------------------
    df = flatten_multilevel_columns(df)
    df = df.loc[:, ~df.columns.str.lower().str.startswith("parent company")]         
    df = df.iloc[1:].reset_index(drop=True)          # 🔹 Drops the first data row (df.iloc[1:]), which may contain merged header content or notes from Excel. Resets the index afterward so the rows are renumbered properly. 🔹
    df = ensure_unique_columns(df)

    # 🔹 2. rename the few columns we care about 🔹
    rename_map = {
        "Company": ["company"],
        "GOGEL Tab": ["gogel tab"],
        "BB Ticker": ["bb ticker"],
        "ISIN equity": ["isin equity"],
        "LEI": ["lei"],
        "Length of Pipelines under Development": ["length of pipelines"],
        "Liquefaction Capacity (Export)":        ["liquefaction capacity"],
        "Regasification Capacity (Import)":      ["regasification capacity"],
        "Total Capacity under Development":      ["total capacity under development"],
    }
    df = rename_columns(df, rename_map)

    # 🔹 3. make sure every canonical column exists 🔹
    needed = list(rename_map.keys())
    for c in needed:
        if c not in df.columns:
            df[c] = np.nan

    # 🔹 4. numeric conversion for the four capacity columns 🔹
    for c in needed[5:]:
        df[c] = pd.to_numeric(
            df[c].astype(str).str.replace(",", "", regex=True),
            errors="coerce"
        ).fillna(0)

    # 🔹 5. flag & reason. For midstream exclusion 🔹
    df["Midstream_Flag"] = (
        (df["Length of Pipelines under Development"] > 0) |
        (df["Liquefaction Capacity (Export)"]        > 0) |
        (df["Regasification Capacity (Import)"]      > 0) |
        (df["Total Capacity under Development"]      > 0)
    )
    df["Excluded"] = df["Midstream_Flag"]
    df["Exclusion Reason"] = np.where(
        df["Midstream_Flag"],
        "Midstream Expansion > 0",
        ""
    )

    excluded = df[df["Excluded"]].copy()
    retained = df[~df["Excluded"]].copy()
    return excluded, retained



if __name__ == "__main__":
    main()
