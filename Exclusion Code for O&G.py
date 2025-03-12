import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

def filter_companies_by_revenue(uploaded_file, sector_exclusions, total_thresholds):
    if uploaded_file is None:
        return None, None
    
    # Load the Excel file
    xls = pd.ExcelFile(uploaded_file)
    df = xls.parse("All Companies", header=[3, 4])
    upstream_df = xls.parse("Upstream", header=[3, 4])
    midstream_df = xls.parse("Midstream Expansion", header=[3, 4])
    
    # Flatten multi-level columns
    df.columns = [' '.join(map(str, col)).strip() for col in df.columns]
    upstream_df.columns = [' '.join(map(str, col)).strip() for col in upstream_df.columns]
    midstream_df.columns = [' '.join(map(str, col)).strip() for col in midstream_df.columns]
    
    # Column Mapping
    column_mapping = {
        "Company Unnamed: 11_level_1": "Company",
        "Bloomberg BB Ticker": "BB Ticker",
        "ISIN Codes ISIN equity": "ISIN equity",
        "LEI LEI": "LEI",
        "Unconventionals Fracking": "Fracking Revenue",
        "Unconventionals Tar Sands": "Tar Sand Revenue",
        "Unconventionals Coalbed Methane": "Coalbed Methane Revenue",
        "Unconventionals Extra Heavy Oil": "Extra Heavy Oil Revenue",
        "Unconventionals Ultra Deepwater": "Ultra Deepwater Revenue",
        "Unconventionals Arctic": "Arctic Revenue",
        "Unconventional Production Unnamed: 25_level_1": "Unconventional Production Revenue"
    }
    
    df.rename(columns=column_mapping, inplace=True, errors='ignore')
    
    # Keep only required columns
    required_columns = list(column_mapping.values()) + ["Exclusion Reason"]
    df = df[list(column_mapping.values())]
    
    # Separate companies with no data
    companies_with_no_data = df[df[list(column_mapping.values())[4:]].isnull().all(axis=1)]
    df = df.dropna(subset=list(column_mapping.values())[4:], how='all')
    
    revenue_columns = list(column_mapping.values())[4:]
    for col in revenue_columns:
        df[col] = df[col].astype(str).str.replace('%', '', regex=True).str.replace(',', '', regex=True)
        df[col] = pd.to_numeric(df[col], errors='coerce')
    
    df[revenue_columns] = df[revenue_columns].fillna(0)
    if df[revenue_columns].max().max() <= 1:
        df[revenue_columns] = df[revenue_columns] * 100
    
    # Calculate total exclusion revenues for selected sectors
    for key, threshold_data in total_thresholds.items():
        selected_sectors = threshold_data["sectors"]
        threshold_value = threshold_data["threshold"]
        valid_sectors = [sector for sector in selected_sectors if sector in df.columns]
        if valid_sectors:
            df[key] = df[valid_sectors].sum(axis=1)
    
    # **Apply Level 1 exclusion logic**
    excluded_reasons = []
    for index, row in df.iterrows():
        reasons = []
        for sector, (exclude, threshold) in sector_exclusions.items():
            if exclude and (threshold == "" or row[sector] > float(threshold)):
                reasons.append(f"{sector} Revenue Exceeded")
        for key, threshold_data in total_thresholds.items():
            threshold_value = threshold_data["threshold"]
            if key in df.columns and row[key] > float(threshold_value):
                reasons.append(f"{key} Revenue Exceeded")
        excluded_reasons.append(", ".join(reasons) if reasons else "")
    
    df["Exclusion Reason"] = excluded_reasons
    retained_companies = df[df["Exclusion Reason"] == ""]
    level1_excluded = df[df["Exclusion Reason"] != ""]

    # **Level 2 Exclusion Logic**
    level2_excluded = set()

    # **Exclude companies from Upstream with fossil fuel share of revenue > 0%**
    if "Fossil Fuel Share of Revenue" in upstream_df.columns:
        upstream_df["Fossil Fuel Share of Revenue"] = upstream_df["Fossil Fuel Share of Revenue"].astype(str).str.replace('%', '', regex=True)
        upstream_df["Fossil Fuel Share of Revenue"] = pd.to_numeric(upstream_df["Fossil Fuel Share of Revenue"], errors='coerce')
        upstream_excluded = upstream_df[upstream_df["Fossil Fuel Share of Revenue"] > 0]["Company"]
        level2_excluded.update(upstream_excluded.tolist())
    
    # **Exclude companies from Midstream Expansion with any expansion activity**
    midstream_columns = [
        "Pipelines Length of Pipelines under Development",
        "Midstream Expansion Liquefaction Capacity (Export)",
        "Midstream Expansion Regasification Capacity (Import)",
        "Midstream Expansion Total Capacity under Development"
    ]
    
    for col in midstream_columns:
        if col in midstream_df.columns:
            midstream_df[col] = midstream_df[col].astype(str).str.replace(',', '', regex=True)
            midstream_df[col] = pd.to_numeric(midstream_df[col], errors='coerce')
            midstream_excluded = midstream_df[midstream_df[col] > 0]["Company"]
            level2_excluded.update(midstream_excluded.tolist())

    # **Store Level 2 Excluded Companies**
    level2_excluded_df = retained_companies[retained_companies["Company"].isin(level2_excluded)]
    level2_retained_df = retained_companies[~retained_companies["Company"].isin(level2_excluded)]

    # **Save to Excel in memory**
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        retained_companies.to_excel(writer, sheet_name="Retained Companies", index=False)
        level1_excluded.to_excel(writer, sheet_name="Excluded Companies (Level 1)", index=False)
        level2_excluded_df.to_excel(writer, sheet_name="Excluded Companies (Level 2)", index=False)
        level2_retained_df.to_excel(writer, sheet_name="Retained Companies (After Level 2)", index=False)
        companies_with_no_data.to_excel(writer, sheet_name="No Data Companies", index=False)
    output.seek(0)
    
    return output, {
        "Total Companies": len(df) + len(companies_with_no_data),
        "Retained Companies": len(retained_companies),
        "Excluded Companies (Level 1)": len(level1_excluded),
        "Excluded Companies (Level 2)": len(level2_excluded_df),
        "Retained Companies (After Level 2)": len(level2_retained_df),
        "Companies with No Data": len(companies_with_no_data)
    }
