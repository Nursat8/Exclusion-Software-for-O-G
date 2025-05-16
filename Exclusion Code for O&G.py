import re
import pandas as pd
import numpy as np
from io import BytesIO
import streamlit as st

# -------------------------- Helper Functions --------------------------

def flatten_multilevel_columns(df):
    df.columns = [
        " ".join(str(level).strip() for level in col).strip()
        for col in df.columns
    ]
    return df

def find_column(df, possible_matches, how="exact", required=True):
    """Locate a column by exact, partial, or regex match (normalized)."""
    norm_map = {}
    for col in df.columns:
        norm = col.strip().lower().replace("\n", " ")
        norm = re.sub(r"\s+", " ", norm)
        norm_map[col] = norm

    pats = []
    for pattern in possible_matches:
        p = pattern.strip().lower().replace("\n", " ")
        p = re.sub(r"\s+", " ", p)
        pats.append(p)

    # exact
    for pat in pats:
        for col, col_norm in norm_map.items():
            if col_norm == pat:
                return col
    # partial
    if how == "partial":
        for pat in pats:
            for col, col_norm in norm_map.items():
                if pat in col_norm:
                    return col
    # regex
    if how == "regex":
        for pattern in possible_matches:
            for col in df.columns:
                if re.search(pattern, col, flags=re.IGNORECASE):
                    return col

    if required:
        raise ValueError(
            f"Could not find a required column among {possible_matches}\n"
            f"Available: {list(df.columns)}"
        )
    return None

def rename_columns(df, rename_map, how="exact"):
    for new_name, patterns in rename_map.items():
        old = find_column(df, patterns, how=how, required=False)
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

# -------------------------- Level 1 Exclusion --------------------------

def filter_companies_by_revenue(uploaded_file, sector_exclusions, total_thresholds):
    if uploaded_file is None:
        return None, None

    xls = pd.ExcelFile(uploaded_file)
    df = xls.parse("All Companies", header=[3,4])
    df.columns = [" ".join(map(str, col)).strip() for col in df.columns]
    df = df.loc[:, ~df.columns.str.lower().str.startswith("parent company")]

    df = remove_equity_from_bb_ticker(df)

    rename_map = {
        "Company": ["company name", "company"],
        "BB Ticker": ["bloomberg bb ticker", "bb ticker"],
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
    df = rename_columns(df, rename_map, how="partial")

    needed = list(rename_map.keys())
    for col in needed:
        if col not in df.columns:
            df[col] = np.nan

    revenue_cols = needed[4:]
    no_data = df[df[revenue_cols].isnull().all(axis=1)].copy()
    df = df.dropna(subset=revenue_cols, how="all")

    for c in revenue_cols:
        df[c] = (
            df[c].astype(str)
                 .str.replace("%","",regex=True)
                 .str.replace(",","",regex=True)
        )
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    for key, info in total_thresholds.items():
        sectors = [s for s in info["sectors"] if s in df.columns]
        df[key] = df[sectors].sum(axis=1) if sectors else 0.0

    reasons = []
    for _, row in df.iterrows():
        parts = []
        for sector, (flag, thresh) in sector_exclusions.items():
            if flag and thresh.strip():
                try:
                    dec = float(thresh)/100
                    if row[sector] > dec:
                        parts.append(f"{sector} > {float(thresh):.1f}%")
                except:
                    pass
        for key, info in total_thresholds.items():
            t = info.get("threshold","").strip()
            if t:
                try:
                    dec = float(t)/100
                    if row[key] > dec:
                        parts.append(f"{key} > {float(t):.1f}%")
                except:
                    pass
        reasons.append("; ".join(parts))
    df["Exclusion Reason"] = reasons

    excluded = df[df["Exclusion Reason"] != ""].copy()
    retained = df[df["Exclusion Reason"] == ""].copy()

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        retained.to_excel(writer, sheet_name="L1 Retained", index=False)
        excluded.to_excel(writer, sheet_name="L1 Excluded", index=False)

        no_data = no_data.reindex(columns=df.columns)
        if "Exclusion Reason" not in no_data.columns:
            no_data["Exclusion Reason"] = ""
        no_data.to_excel(writer, sheet_name="L1 No Data", index=False)
    output.seek(0)

    stats = {
        "Total Companies": len(df) + len(no_data),
        "Retained Companies": len(retained),
        "Excluded Companies": len(excluded),
        "Companies with No Data": len(no_data)
    }
    return output, stats

# -------------------------- Level 2 Exclusion --------------------------

def filter_all_companies(df):
    df = flatten_multilevel_columns(df)
    df = df.loc[:, ~df.columns.str.lower().str.startswith("parent company")]
    df = df.iloc[1:].reset_index(drop=True)

    rename_map = {
        "Company": ["company"],
        "GOGEL Tab": ["gogel tab"],
        "BB Ticker": ["bb ticker", "bloomberg ticker"],
        "ISIN Equity": ["isin equity"],
        "LEI": ["lei"],
        "Length of Pipelines under Development": ["length of pipelines", "pipeline under dev"],
        "Liquefaction Capacity (Export)": ["liquefaction capacity", "lng export capacity"],
        "Regasification Capacity (Import)": ["regasification capacity", "lng import capacity"],
        "Total Capacity under Development": ["total capacity under development"]
    }
    df = rename_columns(df, rename_map, how="partial")

    req = list(rename_map.keys())
    for c in req:
        if c not in df.columns:
            df[c] = 0 if c.startswith(("Length","Liquefaction","Regasification","Total")) else None

    num_cols = req[5:]
    for c in num_cols:
        df[c] = (
            df[c].astype(str)
                 .str.replace("%","",regex=True)
                 .str.replace(",","",regex=True)
        )
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    df["Upstream_Excl"] = df["GOGEL Tab"].str.contains("upstream", case=False, na=False)
    df["Midstream_Excl"] = (
        (df["Length of Pipelines under Development"] > 0)
        | (df["Liquefaction Capacity (Export)"] > 0)
        | (df["Regasification Capacity (Import)"] > 0)
        | (df["Total Capacity under Development"] > 0)
    )
    df["Excluded"] = df["Upstream_Excl"] | df["Midstream_Excl"]
    df["Exclusion Reason"] = df.apply(
        lambda r: "; ".join(
            p for p in (
                "Upstream in GOGEL Tab" if r["Upstream_Excl"] else None,
                "Midstream Expansion > 0" if r["Midstream_Excl"] else None
            ) if p
        ), axis=1
    )

    excluded = df[df["Excluded"]].copy()
    retained = df[~df["Excluded"]].copy()
    cols = req + ["Exclusion Reason"]
    return excluded[cols], retained[cols]

def filter_upstream_companies(df):
    df = flatten_multilevel_columns(df)
    df = df.loc[:, ~df.columns.str.lower().str.startswith("parent company")]

    # 1) Identify columns
    resources_col   = find_column(
        df,
        ["Resources under Development and Field Evaluation"],
        how="partial",
        required=True
    )
    capex_avg_col   = find_column(
        df,
        ["Exploration CAPEX 3-year average", "Exploration CAPEX 3 year average"],
        how="partial",
        required=True
    )
    short_term_col  = find_column(
        df,
        ["Short-Term Expansion ≥20 mmboe", "Short Term Expansion"],
        how="partial",
        required=True
    )
    exp_capex10_col = find_column(
        df,
        ["Exploration CAPEX ≥10 MUSD", "Exploration CAPEX 10 MUSD"],
        how="partial",
        required=True
    )

    # 2) Rename the real company column to "Company"
    company_col = find_column(df, ["Company"], how="partial", required=True)
    df.rename(columns={company_col: "Company"}, inplace=True)

    # 3) Numeric coercion
    for c in (resources_col, capex_avg_col):
        df[c] = pd.to_numeric(df[c].astype(str).str.replace(",","",regex=True), errors="coerce")

    # 4) Build exclusion flags
    df["Resources_Excl"]      = df[resources_col] > 0
    df["CAPEX_Avg_Excl"]      = df[capex_avg_col]   > 0
    df["ShortTerm_Excl"]      = df[short_term_col].astype(str).str.lower().eq("yes")
    df["Exploration10M_Excl"] = df[exp_capex10_col].astype(str).str.lower().eq("yes")
    df["Excluded"] = df[[
        "Resources_Excl",
        "CAPEX_Avg_Excl",
        "ShortTerm_Excl",
        "Exploration10M_Excl"
    ]].any(axis=1)

    # 5) Build exclusion reason
    df["Exclusion Reason"] = df.apply(
        lambda r: "; ".join(
            p for p in (
                "Resources > 0" if r["Resources_Excl"] else None,
                "3-yr CAPEX avg > 0" if r["CAPEX_Avg_Excl"] else None,
                "Short-Term Expansion = Yes" if r["ShortTerm_Excl"] else None,
                "CAPEX ≥10 MUSD = Yes" if r["Exploration10M_Excl"] else None,
            ) if p
        ), axis=1
    )

    # 6) Split
    excluded = df[df["Excluded"]].copy()
    retained = df[~df["Excluded"]].copy()

    # 7) Return with standardized Company column
    out_cols = [
        "Company",
        resources_col,
        capex_avg_col,
        short_term_col,
        exp_capex10_col,
        "Exclusion Reason"
    ]
    return excluded[out_cols], retained[out_cols]

# -------------------------- Streamlit App --------------------------

def main():
    st.title("Level 1 & Level 2 Exclusion Filter for O&G")
    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
    st.sidebar.header("Level 1 Exclusion Settings")

    # Level 1 inputs
    def sector_input(name):
        excl = st.sidebar.checkbox(f"Exclude {name}", value=False)
        thresh = ""
        if excl:
            thresh = st.sidebar.text_input(f"{name} Threshold (%)")
        return name, (excl, thresh)

    sector_list = [
        "Hydrocarbons Production",
        "Fracking Revenue",
        "Tar Sand Revenue",
        "Coalbed Methane Revenue",
        "Extra Heavy Oil Revenue",
        "Ultra Deepwater Revenue",
        "Arctic Revenue",
        "Unconventional Production Revenue",
    ]
    sector_exclusions = dict(sector_input(s) for s in sector_list)

    st.sidebar.header("Custom Total Thresholds")
    total_thresholds = {}
    n = st.sidebar.number_input("How many totals?", 1, 5, 1)
    for i in range(n):
        sels = st.sidebar.multiselect(f"Sectors for Total {i+1}", sector_list, key=f"sel{i}")
        thr  = st.sidebar.text_input(f"Threshold {i+1} (%)", key=f"thr{i}")
        if sels and thr:
            total_thresholds[f"Custom Total {i+1}"] = {"sectors": sels, "threshold": thr}

    # Run Level 1
    if st.sidebar.button("Run Level 1 Exclusion"):
        if not uploaded_file:
            st.warning("Upload a file first.")
        else:
            out1, stats1 = filter_companies_by_revenue(
                uploaded_file, sector_exclusions, total_thresholds
            )
            if out1:
                st.success("Level 1 done")
                st.subheader("Stats")
                for k,v in stats1.items():
                    st.write(f"**{k}:** {v}")
                st.download_button(
                    "Download Level 1 Excel",
                    data=out1,
                    file_name="O&G_Level1.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    st.markdown("---")
    st.header("Level 2 Exclusion (All Companies + Upstream)")
    st.write("Run Level 2 to apply both filters, merge duplicates, and download one file.")

    if st.button("Run Level 2 Exclusion"):
        if not uploaded_file:
            st.warning("Upload a file first.")
        else:
            xls = pd.ExcelFile(uploaded_file)
            # All Companies
            if "All Companies" in xls.sheet_names:
                df_all = pd.read_excel(uploaded_file, "All Companies", header=[3,4])
                excl_all, ret_all = filter_all_companies(df_all)
            else:
                st.error("No 'All Companies' sheet."); return
            # Upstream
            if "Upstream" in xls.sheet_names:
                df_up = pd.read_excel(uploaded_file, "Upstream", header=[3,4])
                excl_up, ret_up = filter_upstream_companies(df_up)
            else:
                st.error("No 'Upstream' sheet."); return

            # Merge exclusions
            merged_exc = pd.concat([
                excl_all[["Company","Exclusion Reason"]],
                excl_up[["Company","Exclusion Reason"]]
            ])
            merged_exc = (
                merged_exc
                .groupby("Company")["Exclusion Reason"]
                .apply(lambda rs: "; ".join(sorted(set("; ".join(rs).split("; ")))))
                .reset_index()
            )

            all_comps = set(excl_all["Company"]) | set(ret_all["Company"]) \
                      | set(excl_up["Company"])  | set(ret_up["Company"])
            retained_l2 = pd.DataFrame({
                "Company": [c for c in all_comps if c not in set(merged_exc["Company"])]
            })

            out2 = BytesIO()
            with pd.ExcelWriter(out2, engine="xlsxwriter") as w:
                # Level 1 sheets
                out1, _ = filter_companies_by_revenue(uploaded_file, sector_exclusions, total_thresholds)
                pd.read_excel(out1, sheet_name="L1 Retained").to_excel(w, "L1 Retained", index=False)
                pd.read_excel(out1, sheet_name="L1 Excluded").to_excel(w, "L1 Excluded", index=False)
                pd.read_excel(out1, sheet_name="L1 No Data").to_excel(w, "L1 No Data", index=False)
                # Level 2 sheets
                merged_exc.to_excel(w, "L2 Excluded", index=False)
                retained_l2.to_excel(w, "L2 Retained", index=False)
            out2.seek(0)

            st.success("Level 2 done")
            st.download_button(
                "Download Combined Results",
                data=out2,
                file_name="O&G_Level1_Level2.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
