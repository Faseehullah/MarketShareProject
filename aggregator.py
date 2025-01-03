# aggregator.py
import logging
import pandas as pd
import numpy as np
from typing import Dict, List, Tuple, Optional, Any, Union
from dataclasses import dataclass

logger = logging.getLogger(__name__)

@dataclass
class AnalysisResult:
    """Container for analysis results.

    Attributes:
        brand_totals (Dict[str, float]): Total yearly workload for each brand
        market_share (Dict[str, float]): Market share percentage for each brand
        brand_values (Optional[Dict[str, float]]): Value-based analysis for each brand
        city_pivot (Optional[pd.DataFrame]): City-wise analysis pivot table
        class_pivot (Optional[pd.DataFrame]): Class-wise analysis pivot table
        region_pivot (Optional[pd.DataFrame]): Region-wise analysis pivot table
        type_pivot (Optional[pd.DataFrame]): Type-wise analysis pivot table
        summary_stats (Dict[str, Any]): Summary statistics of the analysis
    """
    brand_totals: Dict[str, float]
    market_share: Dict[str, float]
    brand_values: Optional[Dict[str, float]] = None
    city_pivot: Optional[pd.DataFrame] = None
    class_pivot: Optional[pd.DataFrame] = None
    region_pivot: Optional[pd.DataFrame] = None
    type_pivot: Optional[pd.DataFrame] = None
    summary_stats: Dict[str, Any] = None

class WorkingAggregator:
    """Handles market share analysis and data aggregation."""

    VALID_CATEGORIES = ["CITY", "Class", "Region", "Type"]
    NULL_VALUES = ["NILL", "", "0", "NIL", "NILL", "null", "NULL", "na", "NA"]

    @staticmethod
    def validate_columns(df: pd.DataFrame, required_cols: List[str]) -> bool:
        """Validate that required columns exist in the dataframe."""
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            raise ValueError(f"Missing required columns: {missing_cols}")
        return True

    @staticmethod
    def validate_data(df: pd.DataFrame) -> bool:
        """Validate data quality and completeness."""
        if df.empty:
            raise ValueError("Empty dataset provided")

        # Check for minimum required columns
        required_base_cols = ["Customer Name", "CITY", "Class", "Region", "Type"]
        WorkingAggregator.validate_columns(df, required_base_cols)

        # Check for data completeness
        null_counts = df[required_base_cols].isnull().sum()
        if null_counts.any():
            logger.warning(f"Null values found in columns: {null_counts[null_counts > 0]}")

        return True

    @staticmethod
    def standardize_brand(brand: Any) -> Optional[str]:
        """Standardize brand names and handle null values."""
        if pd.isnull(brand):
            return None
        brand_str = str(brand).strip().upper()
        if brand_str in WorkingAggregator.NULL_VALUES:
            return None
        return brand_str

    @staticmethod
    def clean_numeric(value: Any) -> float:
        """Clean and convert numeric values."""
        if pd.isnull(value):
            return 0.0
        try:
            return float(value)
        except (ValueError, TypeError):
            return 0.0

    @classmethod
    def allocate_row_brands(
        cls,
        row: pd.Series,
        brand_cols: List[str],
        workload_cols: List[str],
        days_per_year: int
    ) -> List[Tuple[str, float]]:
        """Calculate workload allocation for each brand in a row."""
        daily_sum = 0.0
        brand_workloads = []

        for bcol, wcol in zip(brand_cols, workload_cols):
            brand = cls.standardize_brand(row.get(bcol))
            workload = cls.clean_numeric(row.get(wcol, 0))

            if brand and workload > 0:
                brand_workloads.append((brand, workload))
                daily_sum += workload

        if daily_sum <= 0:
            return []

        total_yearly = daily_sum * days_per_year
        return [(brand, (workload/daily_sum) * total_yearly)
                for brand, workload in brand_workloads]

    @staticmethod
    def calculate_market_share(brand_totals: Dict[str, float]) -> Dict[str, float]:
        """Calculate market share percentages from brand totals."""
        total = sum(brand_totals.values())
        if total <= 0:
            return {}

        sorted_brands = sorted(brand_totals.items(), key=lambda x: x[1], reverse=True)
        return {brand: (value / total) * 100 for brand, value in sorted_brands}

    @classmethod
    def create_pivot_table(
        cls,
        df: pd.DataFrame,
        brand_cols: List[str],
        workload_cols: List[str],
        days_per_year: int,
        groupby_col: str
    ) -> Optional[pd.DataFrame]:
        """Create pivot table for analysis by category."""
        if groupby_col not in df.columns:
            logger.warning(f"Column {groupby_col} not found in dataset")
            return None

        rows = []
        for _, row_data in df.iterrows():
            allocations = cls.allocate_row_brands(
                row_data, brand_cols, workload_cols, days_per_year
            )
            group_value = str(row_data.get(groupby_col, "UNKNOWN")).strip()

            for brand, allocated in allocations:
                rows.append({
                    groupby_col: group_value,
                    "BRAND": brand,
                    "ALLOCATED_YEARLY": allocated
                })

        if not rows:
            return None

        pivot_df = pd.DataFrame(rows)
        grouped = pivot_df.groupby([groupby_col, "BRAND"])["ALLOCATED_YEARLY"].sum()
        final_pivot = grouped.reset_index().pivot(
            index=groupby_col,
            columns="BRAND",
            values="ALLOCATED_YEARLY"
        ).fillna(0)

        return final_pivot.reset_index()

    @classmethod
    def calculate_summary_stats(
        cls,
        df: pd.DataFrame,
        brand_totals: Dict[str, float],
        market_share: Dict[str, float]
    ) -> Dict[str, Any]:
        """Calculate summary statistics for the analysis."""
        return {
            "total_sites": df["Customer Name"].nunique(),
            "total_volume": sum(brand_totals.values()),
            "top_brand": max(market_share.items(), key=lambda x: x[1])[0],
            "unique_cities": df["CITY"].nunique(),
            "unique_classes": df["Class"].nunique(),
            "class_distribution": df["Class"].value_counts().to_dict(),
            "region_distribution": df["Region"].value_counts().to_dict()
        }

    @classmethod
    def analyze_market_data(
        cls,
        df: pd.DataFrame,
        config: Dict,
        analyzer_type: str
    ) -> AnalysisResult:
        """Perform comprehensive market analysis for an analyzer type."""
        try:
            # Validate input data
            cls.validate_data(df)

            if analyzer_type not in config["analyzers"]:
                raise ValueError(f"Invalid analyzer type: {analyzer_type}")

            analyzer_config = config["analyzers"][analyzer_type]
            brand_cols = analyzer_config.get("brand_columns", [])
            workload_cols = analyzer_config.get("workload_columns", [])

            if not brand_cols or not workload_cols:
                raise ValueError(f"Missing column configuration for {analyzer_type}")

            days_per_year = config["analysis_settings"]["days_per_year"]
            test_price = analyzer_config.get("test_price", 0)

            # Volume-based analysis
            brand_totals = cls.calculate_brand_totals(
                df, brand_cols, workload_cols, days_per_year
            )
            market_share = cls.calculate_market_share(brand_totals)

            # Value-based analysis
            brand_values = {
                brand: total * test_price
                for brand, total in brand_totals.items()
            }

            # Create pivot tables
            pivots = {
                category: cls.create_pivot_table(
                    df, brand_cols, workload_cols, days_per_year, category
                ) for category in cls.VALID_CATEGORIES
            }

            # Calculate summary statistics
            summary_stats = cls.calculate_summary_stats(df, brand_totals, market_share)

            return AnalysisResult(
                brand_totals=brand_totals,
                market_share=market_share,
                brand_values=brand_values,
                city_pivot=pivots.get("CITY"),
                class_pivot=pivots.get("Class"),
                region_pivot=pivots.get("Region"),
                type_pivot=pivots.get("Type"),
                summary_stats=summary_stats
            )

        except Exception as e:
            logger.error(f"Error analyzing {analyzer_type} data: {str(e)}")
            raise

    @classmethod
    def calculate_brand_totals(
        cls,
        df: pd.DataFrame,
        brand_cols: List[str],
        workload_cols: List[str],
        days_per_year: int
    ) -> Dict[str, float]:
        """Calculate total yearly workload for each brand."""
        brand_totals = {}

        for _, row in df.iterrows():
            allocations = cls.allocate_row_brands(
                row, brand_cols, workload_cols, days_per_year
            )
            for brand, allocated in allocations:
                brand_totals[brand] = brand_totals.get(brand, 0) + allocated

        return brand_totals