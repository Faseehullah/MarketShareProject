# aggregator.py
import logging
import pandas as pd
import numpy as np
from datetime import datetime

logger = logging.getLogger(__name__)

def standardize_brand(brand):
    if pd.isnull(brand):
        return None
    brand_str = str(brand).strip().upper()
    if brand_str in ["NILL", "", "0"]:
        return None
    return brand_str

def allocate_row_brands(row, brand_cols, workload_cols, days_per_year):
    daily_sum = 0.0
    brand_workloads = []
    for bcol, wcol in zip(brand_cols, workload_cols):
        raw_brand = row.get(bcol)
        brand = standardize_brand(raw_brand)
        w = row.get(wcol, 0) or 0
        if brand and w > 0:
            brand_workloads.append((brand, w))
            daily_sum += w
    if daily_sum <= 0:
        return []
    total_yearly = daily_sum * days_per_year
    allocations = []
    for (brand, w) in brand_workloads:
        proportion = w / daily_sum
        allocated = total_yearly * proportion
        allocations.append((brand, allocated))
    return allocations

def aggregate_analyzer(df, brand_cols, workload_cols, days_per_year):
    brand_totals = {}
    for _, row in df.iterrows():
        pairs = allocate_row_brands(row, brand_cols, workload_cols, days_per_year)
        for brand, allocated in pairs:
            brand_totals[brand] = brand_totals.get(brand, 0) + allocated
    return brand_totals

def calculate_market_share(brand_totals):
    total = sum(brand_totals.values())
    if total <= 0:
        return {}
    items = sorted(brand_totals.items(), key=lambda x: x[1], reverse=True)
    return {b: (val / total)*100 for b, val in items}

def city_pivot_advanced(df, brand_cols, workload_cols, days_per_year):
    rows = []
    for _, row_data in df.iterrows():
        pairs = allocate_row_brands(row_data, brand_cols, workload_cols, days_per_year)
        city = str(row_data.get("CITY", "UNKNOWN")).strip()
        for (brand, allocated) in pairs:
            rows.append({"CITY": city, "BRAND": brand, "ALLOCATED_YEARLY": allocated})
    if not rows:
        return None
    city_df = pd.DataFrame(rows)
    pivoted = city_df.groupby(["CITY", "BRAND"])["ALLOCATED_YEARLY"].sum().reset_index()
    final_pivot = pivoted.pivot(index="CITY", columns="BRAND", values="ALLOCATED_YEARLY").fillna(0)
    final_pivot.reset_index(inplace=True)
    return final_pivot

def class_pivot_advanced(df, brand_cols, workload_cols, days_per_year):
    rows = []
    for _, row_data in df.iterrows():
        pairs = allocate_row_brands(row_data, brand_cols, workload_cols, days_per_year)
        clss = str(row_data.get("Class", "UNKNOWN")).strip()
        for (brand, allocated) in pairs:
            rows.append({"CLASS": clss, "BRAND": brand, "ALLOCATED_YEARLY": allocated})
    if not rows:
        return None
    class_df = pd.DataFrame(rows)
    pivoted = class_df.groupby(["CLASS", "BRAND"])["ALLOCATED_YEARLY"].sum().reset_index()
    final_pivot = pivoted.pivot(index="CLASS", columns="BRAND", values="ALLOCATED_YEARLY").fillna(0)
    final_pivot.reset_index(inplace=True)
    return final_pivot
