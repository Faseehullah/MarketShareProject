# config.py
import json
import logging
import os
from typing import Dict, Any, Optional

logger = logging.getLogger(__name__)

class ConfigurationError(Exception):
    """Custom exception for configuration-related errors."""
    pass

DEFAULT_CONFIG = {
    "days_per_year": 330,
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
        "IA": 100,
        "CBC": 80,
        "CHEM": 120
    }
}

class ConfigManager:
    def __init__(self, config_path: Optional[str] = None):
        self.config_path = config_path or os.path.join(os.path.dirname(__file__), "config.json")
        self._config = self._load_config()

    def _load_config(self) -> Dict[str, Any]:
        if not os.path.exists(self.config_path):
            logger.info("Configuration file not found; using default config.")
            return DEFAULT_CONFIG

        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                config_data = json.load(f)
            return self._validate_config(config_data)
        except json.JSONDecodeError:
            logger.error("Invalid JSON format in configuration file")
            return DEFAULT_CONFIG
        except Exception as e:
            logger.exception("Error loading configuration")
            return DEFAULT_CONFIG

    def _validate_config(self, config_data: Dict[str, Any]) -> Dict[str, Any]:
        try:
            # days_per_year check
            if not isinstance(config_data.get("days_per_year"), int):
                raise ConfigurationError("days_per_year must be an integer")

            # headers
            headers = config_data.get("headers", {})
            required_departments = {"IA", "CBC", "CHEM"}
            for dept in required_departments:
                if dept not in headers:
                    raise ConfigurationError(f"Missing configuration for department: {dept}")
                dept_config = headers[dept]
                if not all(key in dept_config for key in ["brand_cols", "workload_cols"]):
                    raise ConfigurationError(f"Missing column definitions for {dept}")
                if len(dept_config["brand_cols"]) != len(dept_config["workload_cols"]):
                    raise ConfigurationError(f"Mismatched brand vs. workload columns for {dept}")

            # cost_per_test
            cost_dict = config_data.get("cost_per_test", {})
            for dept in required_departments:
                # default 100 if missing
                if dept not in cost_dict:
                    cost_dict[dept] = 100
            config_data["cost_per_test"] = cost_dict

            return config_data

        except Exception as e:
            logger.error(f"Configuration validation failed: {str(e)}")
            return DEFAULT_CONFIG

    def get_config(self) -> Dict[str, Any]:
        return self._config

    def get_brand_columns(self, department: str):
        return self._config["headers"].get(department, {}).get("brand_cols", [])

    def get_workload_columns(self, department: str):
        return self._config["headers"].get(department, {}).get("workload_cols", [])

    def get_days_per_year(self) -> int:
        return self._config.get("days_per_year", 330)

    def get_cost_per_test(self, department: str) -> float:
        return float(self._config["cost_per_test"].get(department, 0))

    def save_config(self, config_data: Dict[str, Any]) -> None:
        validated = self._validate_config(config_data)
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(validated, f, indent=2)
            logger.info("Configuration saved successfully")
            self._config = validated
        except Exception as e:
            logger.exception("Error saving configuration")
            raise ConfigurationError(f"Failed to save config: {str(e)}")


# Global instance
config_manager = ConfigManager()

def load_config():
    return config_manager.get_config()

def save_config(cd: Dict[str, Any]):
    config_manager.save_config(cd)
