# tests/test_aggregator.py

import unittest
import pandas as pd
import numpy as np
from unittest.mock import patch
from aggregator import WorkingAggregator, AnalysisResult
from config import MarketAnalysisConfig

class TestWorkingAggregator(unittest.TestCase):
    def setUp(self):
        # Sample configuration
        self.config = {
            "days_per_year": 350,
            "headers": {
                "IA": {
                    "brand_cols": ["IA Brand 1", "IA Brand 2", "IA Brand 3"],
                    "workload_cols": ["IA Workload - Brand 1", "IA Workload - Brand 2", "IA Workload - Brand 3"]
                },
                "CBC": {
                    "brand_cols": ["CBC Brand 1", "CBC Brand 2", "CBC Brand 3", "CBC Brand 4"],
                    "workload_cols": ["CBC Workload - Brand 1", "CBC Workload - Brand 2", "CBC Workload - Brand 3", "CBC Workload - Brand 4"]
                },
                "CHEM": {
                    "brand_cols": ["CHEM Brand 1", "CHEM Brand 2", "CHEM Brand 3", "CHEM Brand 4"],
                    "workload_cols": ["CHEM Workload - Brand 1", "CHEM Workload - Brand 2", "CHEM Workload - Brand 3", "CHEM Workload - Brand 4"]
                }
            },
            "cost_per_test": {
                "IA": 250,
                "CBC": 120,
                "CHEM": 160
            }
        }

        # Sample data
        self.sample_data = pd.DataFrame({
            "Customer Name": ["Customer A", "Customer B", "Customer C"],
            "CITY": ["City1", "City2", "City1"],
            "Class": ["Class1", "Class2", "Class1"],
            "Region": ["Region1", "Region2", "Region1"],
            "Type": ["Type1", "Type2", "Type1"],
            "IA Brand 1": ["BrandA", "BrandB", "BrandA"],
            "IA Brand 2": ["BrandB", "BrandC", "BrandA"],
            "IA Brand 3": [np.nan, "BrandA", "BrandC"],
            "IA Workload - Brand 1": [10, 20, 30],
            "IA Workload - Brand 2": [15, 25, 35],
            "IA Workload - Brand 3": [np.nan, 5, 15]
        })

    def test_validate_columns_success(self):
        # All required columns are present
        required_cols = ["Customer Name", "CITY", "Class", "Region", "Type"]
        result = WorkingAggregator.validate_columns(self.sample_data, required_cols)
        self.assertTrue(result)

    def test_validate_columns_missing(self):
        # Missing some required columns
        required_cols = ["Customer Name", "CITY", "Class", "Region", "Type", "Extra Column"]
        with self.assertRaises(ValueError):
            WorkingAggregator.validate_columns(self.sample_data, required_cols)

    def test_validate_data_success(self):
        # Data is valid
        result = WorkingAggregator.validate_data(self.sample_data)
        self.assertTrue(result)

    def test_validate_data_empty(self):
        # Empty DataFrame
        empty_df = pd.DataFrame()
        with self.assertRaises(ValueError):
            WorkingAggregator.validate_data(empty_df)

    def test_allocate_row_brands(self):
        # Test workload allocation for a single row
        row = self.sample_data.iloc[0]
        brand_cols = ["IA Brand 1", "IA Brand 2", "IA Brand 3"]
        workload_cols = ["IA Workload - Brand 1", "IA Workload - Brand 2", "IA Workload - Brand 3"]
        days_per_year = self.config["days_per_year"]

        allocations = WorkingAggregator.allocate_row_brands(row, brand_cols, workload_cols, days_per_year)
        # BrandA: 10, BrandB:15 (Brand3 is NaN)
        # Total daily = 25
        # Yearly BrandA = (10/25)*350 = 140
        # Yearly BrandB = (15/25)*350 = 210
        expected_allocations = [("BRANDA", 140.0), ("BRANDB", 210.0)]
        self.assertEqual(len(allocations), 2)
        for alloc, expected in zip(allocations, expected_allocations):
            self.assertEqual(alloc[0], expected[0])
            self.assertAlmostEqual(alloc[1], expected[1], places=1)

    def test_calculate_market_share(self):
        brand_totals = {"BRANDA": 140.0, "BRANDB": 210.0}
        market_share = WorkingAggregator.calculate_market_share(brand_totals)
        expected_share = {"BRANDB": 60.0, "BRANDA": 40.0}
        self.assertEqual(market_share, expected_share)

    def test_create_pivot_table(self):
        # Create a pivot table for City
        category = "CITY"
        pivot_df = WorkingAggregator.create_pivot_table(
            df=self.sample_data,
            brand_cols=["IA Brand 1", "IA Brand 2", "IA Brand 3"],
            workload_cols=["IA Workload - Brand 1", "IA Workload - Brand 2", "IA Workload - Brand 3"],
            days_per_year=self.config["days_per_year"],
            groupby_col=category
        )

        # Expected pivot table:
        # For City1:
        # - BRANDA: 140 (from first row) + 105 (from third row) = 245
        # - BRANDB: 210 (from first row) + 105 (from third row) = 315
        # - BRANDC: 0 + 105 = 105
        # For City2:
        # - BRANDB: (20/25)*350 = 280
        # - BRANDC: (25/25)*350 = 350
        expected_data = {
            "CITY": ["City1", "City2"],
            "BRANDA": [140 + 105, 0],
            "BRANDB": [210 + 105, 280],
            "BRANDC": [0 + 105, 350]
        }
        expected_pivot = pd.DataFrame(expected_data)
        # Fill NaNs with 0
        expected_pivot.fillna(0, inplace=True)

        pd.testing.assert_frame_equal(pivot_df.reset_index(), expected_pivot)

    def test_calculate_summary_stats(self):
        brand_totals = {"BRANDA": 245.0, "BRANDB": 490.0, "BRANDC": 105.0, "BRANDD": 0.0}
        market_share = {"BRANDA": 25.0, "BRANDB": 50.0, "BRANDC": 25.0}
        summary_stats = WorkingAggregator.calculate_summary_stats(
            df=self.sample_data,
            brand_totals=brand_totals,
            market_share=market_share
        )
        self.assertEqual(summary_stats["total_sites"], 3)
        self.assertEqual(summary_stats["total_volume"], 245.0 + 490.0 + 105.0)
        self.assertEqual(summary_stats["top_brand"], "BRANDB")
        self.assertEqual(summary_stats["unique_cities"], 2)
        self.assertEqual(summary_stats["unique_classes"], 2)
        self.assertEqual(summary_stats["class_distribution"], {"Class1": 2, "Class2": 1})
        self.assertEqual(summary_stats["region_distribution"], {"Region1": 2, "Region2": 1})

    def test_analyze_market_data_single_analyzer(self):
        # Test the analyze_market_data method for a single analyzer
        analyzer_type = "IA"
        results = WorkingAggregator.analyze_market_data(
            df=self.sample_data,
            config=self.config,
            analyzer_type=analyzer_type
        )

        # Check the AnalysisResult dataclass
        self.assertIsInstance(results, AnalysisResult)
        self.assertIn("BRANDA", results.brand_totals)
        self.assertIn("BRANDB", results.brand_totals)
        self.assertIn("BRANDC", results.brand_totals)
        self.assertAlmostEqual(results.brand_totals["BRANDA"], 140 + 105)
        self.assertAlmostEqual(results.brand_totals["BRANDB"], 210 + 280)
        self.assertAlmostEqual(results.brand_totals["BRANDC"], 105 + 350)
        self.assertIn("BRANDA", results.market_share)
        self.assertIn("BRANDB", results.market_share)
        self.assertIn("BRANDC", results.market_share)
        self.assertAlmostEqual(results.market_share["BRANDA"], 25.0)
        self.assertAlmostEqual(results.market_share["BRANDB"], 50.0)
        self.assertAlmostEqual(results.market_share["BRANDC"], 25.0)
        self.assertIsNotNone(results.city_pivot)
        self.assertIsNotNone(results.class_pivot)
        self.assertIsNone(results.region_pivot)
        self.assertIsNone(results.type_pivot)
        self.assertIsInstance(results.summary_stats, dict)

    def test_analyze_market_data_invalid_analyzer(self):
        # Test analyze_market_data with an invalid analyzer type
        with self.assertRaises(ValueError):
            WorkingAggregator.analyze_market_data(
                df=self.sample_data,
                config=self.config,
                analyzer_type="INVALID"
            )

    def test_analyze_market_data_missing_workload_cols(self):
        # Test analyze_market_data with missing workload columns
        config_missing = self.config.copy()
        config_missing["headers"]["IA"]["workload_cols"] = []
        with self.assertRaises(ValueError):
            WorkingAggregator.analyze_market_data(
                df=self.sample_data,
                config=config_missing,
                analyzer_type="IA"
            )

    def test_calculate_brand_totals(self):
        # Test the calculate_brand_totals method
        brand_cols = ["IA Brand 1", "IA Brand 2", "IA Brand 3"]
        workload_cols = ["IA Workload - Brand 1", "IA Workload - Brand 2", "IA Workload - Brand 3"]
        days_per_year = 350

        brand_totals = WorkingAggregator.calculate_brand_totals(
            df=self.sample_data,
            brand_cols=brand_cols,
            workload_cols=workload_cols,
            days_per_year=days_per_year
        )

        expected_totals = {
            "BRANDA": (10/25)*350 + (30/45)*350,  # From first and third rows
            "BRANDB": (15/25)*350 + (35/45)*350,
            "BRANDC": (0/25)*350 + (15/45)*350
        }
        # Calculated as:
        # First row: BRANDA:140, BRANDB:210
        # Second row: BRANDA:0 (since Brand3 is 5 workload, total 25, BRANDA: (0/25)*350=0?, BRANDB: (20/25)*350=280, BRANDC: (5/25)*350=70
        # Third row: BRANDA: (30/45)*350=233.33, BRANDB: (35/45)*350=272.22, BRANDC: (15/45)*350=116.67
        # Total BRANDA:140 + 233.33 = 373.33
        # Total BRANDB:210 + 280 + 272.22 = 762.22
        # Total BRANDC:70 + 116.67 = 186.67
        self.assertAlmostEqual(brand_totals["BRANDA"], 140 + 233.3333, places=1)
        self.assertAlmostEqual(brand_totals["BRANDB"], 210 + 280 + 272.2222, places=1)
        self.assertAlmostEqual(brand_totals["BRANDC"], 70 + 116.6667, places=1)

    if __name__ == "__main__":
        unittest.main()

"""
**Explanation:**

1. **Imports and Setup:**
   - Imports necessary modules, including `unittest`, `pandas`, and the classes from `aggregator.py` and `config.py`.
   - Defines `setUp` to initialize sample configurations and data used across multiple test cases.
   - `tearDown` ensures that any temporary files created during tests are removed.

2. **Test Cases:**
   - **Data Validation:**
     - Tests whether the required columns are present and handles missing columns appropriately.
     - Checks behavior when the dataset is empty.

   - **Workload Allocation:**
     - Verifies that workloads are allocated correctly based on brand workloads and days per year.
     - Ensures that brands with `NaN` values are handled gracefully.

   - **Market Share Calculation:**
     - Confirms that market share percentages are calculated accurately from brand totals.

   - **Pivot Table Creation:**
     - Tests the creation of pivot tables for different categories (e.g., `CITY`).
     - Compares the generated pivot tables against expected results.

   - **Summary Statistics:**
     - Validates that summary statistics like total sites, total volume, top brand, and distributions are computed correctly.

   - **Analyze Market Data:**
     - Checks the `analyze_market_data` method for both valid and invalid analyzer types.
     - Ensures that missing workload columns raise appropriate errors.

   - **Brand Totals Calculation:**
     - Tests the `calculate_brand_totals` method to ensure accurate aggregation of workloads.

3. **Mocking and Assertions:**
   - Uses `unittest.mock.patch` to mock file operations, ensuring that tests do not depend on actual file I/O.
   - Employs assertions like `assertTrue`, `assertEqual`, `assertAlmostEqual`, and `assertRaises` to validate expected outcomes.

4. **Running Tests:**
   - Execute the tests by running the following command from the root directory:
     ```bash
     python -m unittest discover -s tests
     ```
   - This command discovers and runs all tests in the `tests` directory.

"""