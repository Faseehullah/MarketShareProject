# config.py
import json
import logging
import os
from typing import Dict, List
from dataclasses import dataclass

logger = logging.getLogger(__name__)

@dataclass
class AnalyzerConfig:
    """Configuration structure for each analyzer type."""
    brand_cols: List[str]
    workload_cols: List[str]

class MarketAnalysisConfig:
    def __init__(self, config_path: str = "config.json"):
        self.config_path = config_path
        self.config_data = self._load_default_config()
        self.load_config()

    def _load_default_config(self) -> Dict:
        """Load default configuration settings."""
        return {
            "days_per_year": 350,
            "headers": {
                "IA": {
                    "brand_cols": [f"IA Brand {i}" for i in range(1, 4)],
                    "workload_cols": [f"IA Workload - Brand {i}" for i in range(1, 4)]
                },
                "CBC": {
                    "brand_cols": [f"CBC Brand {i}" for i in range(1, 5)],
                    "workload_cols": [f"CBC Workload - Brand {i}" for i in range(1, 5)]
                },
                "CHEM": {
                    "brand_cols": [f"CHEM Brand {i}" for i in range(1, 5)],
                    "workload_cols": [f"CHEM Workload - Brand {i}" for i in range(1, 5)]
                }
            },
            "cost_per_test": {
                "IA": 250.0,
                "CBC": 120.0,
                "CHEM": 160.0
            }
        }

    def load_config(self):
        """Load configuration from file."""
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r') as f:
                    loaded_config = json.load(f)
                    self._update_config(loaded_config)
                    logger.info(f"Configuration loaded from {self.config_path}")
            else:
                self.save_config()
                logger.info(f"Default configuration saved to {self.config_path}")
        except Exception as e:
            logger.error(f"Error loading config: {e}")

    def _update_config(self, loaded_config: Dict):
        """Update the default config with loaded config."""
        for key, value in loaded_config.items():
            if isinstance(value, dict) and key in self.config_data:
                self.config_data[key].update(value)
            else:
                self.config_data[key] = value

    def save_config(self):
        """Save current configuration to file."""
        try:
            with open(self.config_path, 'w') as f:
                json.dump(self.config_data, f, indent=2)
            logger.info(f"Configuration saved to {self.config_path}")
        except Exception as e:
            logger.error(f"Error saving config: {e}")

    def get_headers(self) -> Dict[str, AnalyzerConfig]:
        """Get headers configuration for analyzers."""
        headers = self.config_data.get("headers", {})
        analyzer_configs = {}
        for analyzer, config in headers.items():
            analyzer_configs[analyzer] = AnalyzerConfig(
                brand_cols=config.get("brand_cols", []),
                workload_cols=config.get("workload_cols", [])
            )
        return analyzer_configs

    def set_headers(self, headers: Dict[str, AnalyzerConfig]):
        """Set headers configuration for analyzers."""
        self.config_data["headers"] = {}
        for analyzer, config in headers.items():
            self.config_data["headers"][analyzer] = {
                "brand_cols": config.brand_cols,
                "workload_cols": config.workload_cols
            }

    def get_cost_per_test(self) -> Dict[str, float]:
        """Get cost per test for each analyzer."""
        return self.config_data.get("cost_per_test", {})

    def set_cost_per_test(self, costs: Dict[str, float]):
        """Set cost per test for each analyzer."""
        self.config_data["cost_per_test"] = costs

    def get_days_per_year(self) -> int:
        """Get the number of working days per year."""
        return self.config_data.get("days_per_year", 350)

    def set_days_per_year(self, days: int):
        """Set the number of working days per year."""
        self.config_data["days_per_year"] = days
