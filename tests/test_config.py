# tests/test_config.py

import unittest
import os
import json
from unittest.mock import patch, mock_open
from config import MarketAnalysisConfig, AnalyzerConfig

class TestMarketAnalysisConfig(unittest.TestCase):
    def setUp(self):
        # Define a sample default configuration
        self.default_config = {
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
        self.mock_config_path = "test_config.json"

    def tearDown(self):
        # Remove the test config file if it exists
        if os.path.exists(self.mock_config_path):
            os.remove(self.mock_config_path)

    @patch("builtins.open", new_callable=mock_open, read_data=json.dumps({}))
    def test_load_default_config_when_no_file_exists(self, mock_file):
        # Simulate no existing config file
        with patch("os.path.exists", return_value=False):
            config_manager = MarketAnalysisConfig(config_path=self.mock_config_path)
            mock_file.assert_called_with(self.mock_config_path, 'w')
            self.assertEqual(config_manager.config_data, self.default_config)

    @patch("builtins.open", new_callable=mock_open, read_data=json.dumps({
        "days_per_year": 300,
        "headers": {
            "IA": {
                "brand_cols": ["IA Brand A", "IA Brand B"],
                "workload_cols": ["IA Workload A", "IA Workload B"]
            }
        },
        "cost_per_test": {
            "IA": 200
        }
    }))
    def test_load_existing_config(self, mock_file):
        # Simulate existing config file
        with patch("os.path.exists", return_value=True):
            config_manager = MarketAnalysisConfig(config_path=self.mock_config_path)
            mock_file.assert_called_with(self.mock_config_path, 'r')
            expected_config = self.default_config.copy()
            expected_config.update({
                "days_per_year": 300,
                "headers": {
                    "IA": AnalyzerConfig(
                        brand_cols=["IA Brand A", "IA Brand B"],
                        workload_cols=["IA Workload A", "IA Workload B"]
                    )
                },
                "cost_per_test": {
                    "IA": 200
                }
            })
            # Since headers are merged, verify specific updates
            self.assertEqual(config_manager.config_data["days_per_year"], 300)
            self.assertEqual(config_manager.config_data["headers"]["IA"]["brand_cols"], ["IA Brand A", "IA Brand B"])
            self.assertEqual(config_manager.config_data["headers"]["IA"]["workload_cols"], ["IA Workload A", "IA Workload B"])
            self.assertEqual(config_manager.config_data["cost_per_test"]["IA"], 200)

    def test_get_headers(self):
        config_manager = MarketAnalysisConfig(config_path=self.mock_config_path)
        config_manager.config_data = self.default_config  # Manually set config_data
        headers = config_manager.get_headers()
        self.assertIsInstance(headers, dict)
        self.assertIn("IA", headers)
        self.assertIsInstance(headers["IA"], AnalyzerConfig)
        self.assertEqual(headers["IA"].brand_cols, ["IA Brand 1", "IA Brand 2", "IA Brand 3"])
        self.assertEqual(headers["IA"].workload_cols, ["IA Workload - Brand 1", "IA Workload - Brand 2", "IA Workload - Brand 3"])

    def test_set_headers(self):
        config_manager = MarketAnalysisConfig(config_path=self.mock_config_path)
        new_headers = {
            "IA": AnalyzerConfig(
                brand_cols=["IA Brand X", "IA Brand Y"],
                workload_cols=["IA Workload X", "IA Workload Y"]
            ),
            "NEW_ANALYZER": AnalyzerConfig(
                brand_cols=["New Brand 1"],
                workload_cols=["New Workload 1"]
            )
        }
        config_manager.set_headers(new_headers)
        self.assertIn("IA", config_manager.config_data["headers"])
        self.assertIn("NEW_ANALYZER", config_manager.config_data["headers"])
        self.assertEqual(config_manager.config_data["headers"]["IA"]["brand_cols"], ["IA Brand X", "IA Brand Y"])
        self.assertEqual(config_manager.config_data["headers"]["NEW_ANALYZER"]["workload_cols"], ["New Workload 1"])

    def test_get_cost_per_test(self):
        config_manager = MarketAnalysisConfig(config_path=self.mock_config_path)
        config_manager.config_data = self.default_config  # Manually set config_data
        costs = config_manager.get_cost_per_test()
        self.assertIsInstance(costs, dict)
        self.assertIn("IA", costs)
        self.assertEqual(costs["IA"], 250)

    def test_set_cost_per_test(self):
        config_manager = MarketAnalysisConfig(config_path=self.mock_config_path)
        new_costs = {
            "IA": 300,
            "NEW_ANALYZER": 150
        }
        config_manager.set_cost_per_test(new_costs)
        self.assertEqual(config_manager.config_data["cost_per_test"]["IA"], 300)
        self.assertEqual(config_manager.config_data["cost_per_test"]["NEW_ANALYZER"], 150)

    def test_get_days_per_year(self):
        config_manager = MarketAnalysisConfig(config_path=self.mock_config_path)
        config_manager.config_data = self.default_config
        days = config_manager.get_days_per_year()
        self.assertEqual(days, 350)

    def test_set_days_per_year(self):
        config_manager = MarketAnalysisConfig(config_path=self.mock_config_path)
        config_manager.set_days_per_year(320)
        self.assertEqual(config_manager.config_data["days_per_year"], 320)

    @patch("builtins.open", new_callable=mock_open)
    def test_save_config(self, mock_file):
        config_manager = MarketAnalysisConfig(config_path=self.mock_config_path)
        config_manager.config_data = self.default_config
        config_manager.save_config()
        mock_file.assert_called_with(self.mock_config_path, 'w')
        handle = mock_file()
        handle.write.assert_called_once_with(json.dumps(self.default_config, indent=2))

    if __name__ == "__main__":
        unittest.main()

"""
**Explanation:**

1. **Imports and Setup:**
   - Imports necessary modules, including `unittest`, `os`, `json`, and `unittest.mock` for mocking file operations.
   - Imports the `MarketAnalysisConfig` and `AnalyzerConfig` classes from `config.py`.

2. **Test Cases:**
   - **`test_load_default_config_when_no_file_exists`**: Verifies that when no configuration file exists, the default configuration is loaded and saved correctly.
   - **`test_load_existing_config`**: Checks that an existing configuration file is loaded properly, updating only specified keys.
   - **`test_get_headers` & `test_set_headers`**: Ensures that headers are retrieved and set correctly.
   - **`test_get_cost_per_test` & `test_set_cost_per_test`**: Validates the retrieval and setting of cost per test.
   - **`test_get_days_per_year` & `test_set_days_per_year`**: Tests getting and setting the number of working days per year.
   - **`test_save_config`**: Confirms that the configuration is saved to a file with the correct formatting.

3. **Mocking:**
   - Uses `unittest.mock.patch` to mock file operations, preventing actual file I/O during testing.

4. **Running Tests:**
   - Execute the tests by running the following command from the root directory:
     ```bash
     python -m unittest discover -s tests
     ```
   - This command discovers and runs all tests in the `tests` directory.

### **b. Testing `aggregator.py` with `test_aggregator.py`**

The `aggregator.py` module performs the core data processing and market analysis. We'll write tests to verify its functionalities, including data validation, brand workload allocation, market share calculation, and summary statistics.

**File:** `tests/test_aggregator.py`
"""

