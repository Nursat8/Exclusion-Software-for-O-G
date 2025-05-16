import re
import pandas as pd
import numpy as np
from io import BytesIO
import streamlit as st

# ---------------- Helper Functions ----------------

def flatten_multilevel_columns(df):
    df.columns = [
        " ".join(str(level).strip() for level in col).strip()
        for col in df.columns
    ]
    return df

def find_column(df, patterns, how="partial", required=True):
    """Normalize and find a column by partial, exact or regex match."""
    norm_map = {}
    for col in df.columns:
        norm = col.strip().lower().replace("\n", " ")
        norm = re.sub(r"\s+", " ", norm)
        norm_map[col] = norm

    pats = []
    for p in patterns:
        norm = p.strip().lower().replace("\n", " ")
        norm = re.sub(r"\s+", " ", norm)
        pats.append(norm)

    # exact match
    for pat in pats:
        for col, norm in norm_map.items():
            if norm == pat:
                return col
    # partial match
    if how in ("partial",):
        for pat in pats:
            for col, norm in norm_map.items():
                if pat in norm:
                    return col
    # regex match
    if how == "regex":
        for pattern in patterns:
            for col in df.columns:
                if re.search(pattern, col, flags=re.IGNORECASE):
                    return col

    if required:
        raise ValueError(f"Could not find column among {patterns}\nAvailable: {list(df.columns)}")
    return None

def rename_columns(df, rename_map):
    for new_name, pats in rename_map.items():
        old = find_column(df, pats, how="partial", required=False)
        if old and old != new_name:
            df.rename(columns={old: new_name}, inplace=True)
    return df

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

# ---------------- Level 1 Exclusion ----------------

def filter_companies_by_revenue(uploaded_file, sector_exclusions, total_thresholds):
    xls = pd.ExcelFile(uploaded_file)
    df = xls.parse("All Companies", header=[3,4])
    df.columns = [" ".join(map(str,c)).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.str.lower().str.startswith("parent company")]
    df = remove_equity_from_bb_ticker(df)

    # rename revenue columns
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

    # ensure all columns exist
    needed = list(rename_map.keys())
    for col in needed:
        if col not in df.columns:
            df[col] = np.nan

    # slice no-data
    revenue_cols = needed[4:]
    no_data = df[df[revenue_cols].isnull().all(axis=1)].copy()
    df = df.dropna(subset=revenue_cols, how="all")

    # clean & numeric
    for c in revenue_cols:
        df[c] = df[c].astype(str).str.replace("%","",regex=True).str.replace(",","",regex=True)
        df[c] = pd.to_numeric(df[c],errors="coerce").fillna(0)

    # custom totals
    for key,info in total_thresholds.items():
        secs = [s for s in info["sectors"] if s in df.columns]
        df[key] = df[secs].sum(axis=1) if secs else 0.0

    # build reasons
    reasons = []
    for _,r in df.iterrows():
        parts=[]
        for sector,(flag,thr) in sector_exclusions.items():
            if flag and thr.strip():
                try:
                    if r[sector] > float(thr)/100: parts.append(f"{sector} > {thr}%")
                except: pass
        for key,info in total_thresholds.items():
            thr=info.get("threshold","").strip()
            if thr:
                try:
                    if r[key] > float(thr)/100: parts.append(f"{key} > {thr}%")
                except: pass
        reasons.append("; ".join(parts))
    df["Reason"] = reasons

    excluded = df[df["Reason"]!=""].copy()
    retained = df[df["Reason"]==""].copy()

    # rename first custom total
    if "Custom Total 1" in df.columns:
        excluded.rename(columns={"Custom Total 1":"Custom Total Revenue"}, inplace=True)
        retained.rename(columns={"Custom Total 1":"Custom Total Revenue"}, inplace=True)
        no_data.rename(columns={"Custom Total 1":"Custom Total Revenue"}, inplace=True)

    # drop all others custom totals if any
    for col in [c for c in df.columns if c.startswith("Custom Total") and c!="Custom Total Revenue"]:
        excluded.drop(columns=[col],errors="ignore",inplace=True)
        retained.drop(columns=[col],errors="ignore",inplace=True)
        no_data.drop(columns=[col],errors="ignore",inplace=True)

    return excluded, retained, no_data

# ---------------- Level 2 Exclusion ----------------

def filter_all_companies(df):
    df = flatten_multilevel_columns(df)
    df = df.loc[:, ~df.columns.str.lower().str.startswith("parent company")]
    df = df.iloc[1:].reset_index(drop=True)
    rename_map = {
        "Company":["company"],
        "GOGEL Tab":["gogel tab"],
        "BB Ticker":["bb ticker"],
        "ISIN equity":["isin equity"],
        "LEI":["lei"],
        "Length of Pipelines under Development":["length of pipelines"],
        "Liquefaction Capacity (Export)":["liquefaction capacity"],
        "Regasification Capacity (Import)":["regasification capacity"],
        "Total Capacity under Development":["total capacity under development"]
    }
    df = rename_columns(df, rename_map)
    # ensure columns
    req=list(rename_map.keys())
    for c in req:
        if c not in df.columns:
            df[c]=0 if c.startswith(("Length","Liquefaction","Regasification","Total")) else None
    # numeric coercion
    for c in req[5:]:
        df[c]=pd.to_numeric(df[c].astype(str).str.replace(",","",regex=True),errors="coerce").fillna(0)
    df["Upstream_Flag"]=df["GOGEL Tab"].str.contains("upstream",case=False,na=False)
    df["Midstream_Flag"]=(
        (df["Length of Pipelines under Development"]>0)
        |(df["Liquefaction Capacity (Export)"]>0)
        |(df["Regasification Capacity (Import)"]>0)
        |(df["Total Capacity under Development"]>0)
    )
    df["Excluded_F1"]=df["Upstream_Flag"]|df["Midstream_Flag"]
    df["Exclusion Reason_F1"]=df.apply(
        lambda r: "; ".join(p for p in (
            "Upstream in GOGEL Tab" if r["Upstream_Flag"] else None,
            "Midstream Expansion > 0" if r["Midstream_Flag"] else None
        ) if p), axis=1
    )
    exc=df[df["Excluded_F1"]].copy()
    ret=df[~df["Excluded_F1"]].copy()
    cols=req+["Exclusion Reason_F1"]
    exc=exc[cols]; ret=ret[cols]
    exc.rename(columns={"Exclusion Reason_F1":"Exclusion Reason"},inplace=True)
    ret.rename(columns={"Exclusion Reason_F1":"Exclusion Reason"},inplace=True)
    return exc,ret

def filter_upstream_companies(df):
    df=flatten_multilevel_columns(df)
    df=df.loc[:,~df.columns.str.lower().str.startswith("parent company")]

    # find and rename Company
    comp = find_column(df,["company"],how="partial",required=True)
    df.rename(columns={comp:"Company"},inplace=True)

    resources = find_column(df,["resources under development and field evaluation"],how="partial")
    capex_avg = find_column(df,["exploration capex 3-year average"],how="partial")
    shortterm = find_column(df,["short-term expansion ≥20 mmboe"],how="partial")
    capex10   = find_column(df,["exploration capex ≥10 MUSD"],how="partial")

    # numeric coercion
    df[resources]=pd.to_numeric(df[resources].astype(str).str.replace(",","",regex=True),errors="coerce")
    df[capex_avg]=pd.to_numeric(df[capex_avg].astype(str).str.replace(",","",regex=True),errors="coerce")

    # flags
    df["F2_Res"]=df[resources]>0
    df["F2_Avg"]=df[capex_avg]>0
    df["F2_ST"]=df[shortterm].astype(str).str.lower().eq("yes")
    df["F2_10M"]=df[capex10].astype(str).str.lower().eq("yes")
    df["Excluded"]=df[["F2_Res","F2_Avg","F2_ST","F2_10M"]].any(axis=1)
    df["Exclusion Reason"]=df.apply(
        lambda r: "; ".join(p for p in (
            "Resources > 0"       if r["F2_Res"] else None,
            "3-yr CAPEX avg > 0"  if r["F2_Avg"] else None,
            "Short-Term Expansion = Yes" if r["F2_ST"] else None,
            "CAPEX ≥10 MUSD = Yes"     if r["F2_10M"] else None
        ) if p), axis=1
    )

    exc = df[df["Excluded"]].copy()
    ret = df[~df["Excluded"]].copy()
    out_cols = [
        "Company","GOGEL Tab","Length of Pipelines under Development",
        "Liquefaction Capacity (Export)","Regasification Capacity (Import)",
        "Total Capacity under Development",resources,capex_avg,shortterm,capex10,"Exclusion Reason"
    ]
    return exc[out_cols], ret[out_cols]

# ---------------- Streamlit App ----------------

def main():
    st.title("Level 1 & Level 2 Exclusion Filter for O&G")
    uploaded = st.file_uploader("Upload Excel file",type=["xlsx"])

    # Sidebar Level 1 settings
    st.sidebar.header("Level 1 Exclusion Settings")
    sector_list = [
        "Hydrocarbons Production (%)","Fracking Revenue","Tar Sand Revenue",
        "Coalbed Methane Revenue","Extra Heavy Oil Revenue","Ultra Deepwater Revenue",
        "Arctic Revenue","Unconventional Production Revenue",
    ]
    sector_excs = {}
    for s in sector_list:
        chk = st.sidebar.checkbox(f"Exclude {s}",value=False)
        thr = ""
        if chk:
            thr = st.sidebar.text_input(f"{s} Threshold (%)")
        sector_excs[s]=(chk,thr)

    st.sidebar.header("Custom Total Thresholds")
    total_thrs = {}
    n = st.sidebar.number_input("How many totals?",1,5,1)
    for i in range(n):
        sels = st.sidebar.multiselect(f"Sectors for Total {i+1}",sector_list,key=f"sel{i}")
        thr = st.sidebar.text_input(f"Total {i+1} (%)",key=f"thr{i}")
        if sels and thr:
            total_thrs[f"Custom Total {i+1}"]={"sectors":sels,"threshold":thr}

    # Run L1
    if st.sidebar.button("Run Level 1 Exclusion"):
        if not uploaded:
            st.warning("Please upload first.")
        else:
            exc1,ret1,no1 = filter_companies_by_revenue(uploaded,sector_excs,total_thrs)
            st.success("Level 1 done")
            st.download_button("Download L1 Results",data=lambda: to_excel_l1(exc1,ret1,no1),
                                file_name="O&G_L1_Exclusion.xlsx")

    st.markdown("---")
    st.header("Level 2 Exclusion")
    st.write("Applies All-Companies + Upstream filters and merges duplicates.")

    if st.button("Run Level 2 Exclusion"):
        if not uploaded:
            st.warning("Please upload first.")
            return

        xls = pd.ExcelFile(uploaded)
        # All Companies filter
        df_all = pd.read_excel(uploaded,"All Companies",header=[3,4])
        exc_all,ret_all = filter_all_companies(df_all)
        # Upstream filter
        df_up = pd.read_excel(uploaded,"Upstream",header=[3,4])
        exc_up,ret_up = filter_upstream_companies(df_up)

        # merge L2 excluded
        merged = pd.concat([
            exc_all[["Company","Exclusion Reason"]],
            exc_up[["Company","Exclusion Reason"]]
        ])
        merged = (merged.groupby("Company")["Exclusion Reason"]
                      .apply(lambda rs: "; ".join(sorted(set("; ".join(rs).split("; ")))))
                      .reset_index())

        # union all companies then subtract excluded for retained L2
        all_comps = set(ret_all["Company"])|set(ret_up["Company"])|set(exc_all["Company"])|set(exc_up["Company"])
        ret2 = pd.DataFrame({"Company":[c for c in all_comps if c not in set(merged["Company"])]})

        st.success("Level 2 done")
        st.download_button("Download Combined Results",data=lambda: to_excel_l2(exc1,exc_all,merged,ret1,ret_all,ret2,exc_up,ret_up),
                           file_name="O&G_L1_L2_Exclusion.xlsx")

# Helper to write L1 Excel
def to_excel_l1(exc,ret,no_data):
    desired_cols = [
        "Company","BB Ticker","ISIN equity","LEI",
        "Hydrocarbons Production (%)","Fracking Revenue","Tar Sand Revenue",
        "Coalbed Methane Revenue","Extra Heavy Oil Revenue","Ultra Deepwater Revenue",
        "Arctic Revenue","Unconventional Production Revenue","Reason",
        "Custom Total Revenue"
    ]
    buf = BytesIO()
    with pd.ExcelWriter(buf,engine="xlsxwriter") as w:
        exc.reindex(columns=desired_cols).to_excel(w,"Excluded Level 1",index=False)
        ret.reindex(columns=desired_cols).to_excel(w,"Retained Level 1",index=False)
        no_data.reindex(columns=desired_cols).to_excel(w,"L1 No Data",index=False)
    buf.seek(0)
    return buf

# Helper to write combined L2 Excel
def to_excel_l2(exc1,exc_all,exc_l2,ret1,ret_all,ret2,exc_up,ret_up):
    desired = [
        "Company","BB Ticker","ISIN equity","LEI",
        "Hydrocarbons Production (%)","Fracking Revenue","Tar Sand Revenue",
        "Coalbed Methane Revenue","Extra Heavy Oil Revenue","Ultra Deepwater Revenue",
        "Arctic Revenue","Unconventional Production Revenue","Reason",
        "Custom Total Revenue","GOGEL Tab","Length of Pipelines under Development",
        "Liquefaction Capacity (Export)","Regasification Capacity (Import)",
        "Total Capacity under Development",
        "Resources under Development and Field Evaluation",
        "Exploration CAPEX 3-year average","Short-Term Expansion ≥20 mmboe",
        "Exploration CAPEX ≥10 MUSD","Exclusion Reason"
    ]
    buf=BytesIO()
    with pd.ExcelWriter(buf,engine="xlsxwriter") as w:
        exc_all.reindex(columns=desired).to_excel(w,"All Excluded",index=False)
        exc1   .reindex(columns=desired).to_excel(w,"Excluded Level 1",index=False)
        exc_l2 .reindex(columns=desired).to_excel(w,"Excluded Level 2",index=False)
        ret1   .reindex(columns=desired).to_excel(w,"Retained Level 1",index=False)
        ret2   .reindex(columns=desired).to_excel(w,"Retained Level 2",index=False)
        exc_up .reindex(columns=desired).to_excel(w,"Upstream Excluded",index=False)
        ret_up .reindex(columns=desired).to_excel(w,"Upstream Retained",index=False)
    buf.seek(0)
    return buf

if __name__ == "__main__":
    main()

