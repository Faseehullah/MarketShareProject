# config.py
import json
import logging
import os
from typing import Dict, List, Any, Optional
from dataclasses import dataclass

logger = logging.getLogger(__name__)

@dataclass
class AnalyzerConfig:
    """Configuration structure for each analyzer type."""
    name: str
    brand_columns: List[str]
    workload_columns: List[str]
    test_price: float = 0.0

class MarketAnalysisConfig:
    def __init__(self, config_path: str = "config.json"):
        self.config_path = config_path
        self.config_data = self._load_default_config()
        self.load_config()

    def _load_default_config(self) -> Dict:
        """Load default configuration settings."""
        return {
            "analyzers": {
                "IA": {
                    "name": "Immunoassay",
                    "brand_columns": [f"IA Brand {i}" for i in range(1, 4)],
                    "workload_columns": [f"IA Workload - Brand {i}" for i in range(1, 4)],
                    "test_price": 250.0
                },
                "CBC": {
                    "name": "Hematology",
                    "brand_columns": [f"CBC Brand {i}" for i in range(1, 5)],
                    "workload_columns": [f"CBC Workload - Brand {i}" for i in range(1, 5)],
                    "test_price": 150.0
                },
                "CHEM": {
                    "name": "Clinical Chemistry",
                    "brand_columns": [f"CHEM Brand {i}" for i in range(1, 5)],
                    "workload_columns": [f"CHEM Workload - Brand {i}" for i in range(1, 5)],
                    "test_price": 100.0
                }
            },
            "metadata": {
                "regions": ["SOUTH", "NORTH", "CENTRAL"],
                "classes": ["CLASS A", "CLASS B", "CLASS C", "CLASS D"],
                "types": ["PRIVATE", "GOVT", "NGO", "ARMFORCES"]
            },
            "analysis_settings": {
                "days_per_year": 330,
                "include_model_analysis": True,
                "value_analysis": True
            }
        }

    def load_config(self):
        """Load configuration from file."""
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r') as f:
                    loaded_config = json.load(f)
                    self.config_data.update(loaded_config)
        except Exception as e:
            logger.error(f"Error loading config: {e}")

    def save_config(self):
        """Save current configuration to file."""
        try:
            with open(self.config_path, 'w') as f:
                json.dump(self.config_data, f, indent=2)
        except Exception as e:
            logger.error(f"Error saving config: {e}")

    def get_analyzer_config(self, analyzer_type: str) -> Dict:
        """Get configuration for specific analyzer type."""
        return self.config_data["analyzers"].get(analyzer_type, {})

    def get_analysis_settings(self) -> Dict:
        """Get general analysis settings."""
        return self.config_data.get("analysis_settings", {})