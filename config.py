# config.py
import json
import logging
import os

logger = logging.getLogger(__name__)

CONFIG_FILE = "config.json"

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
    }
}

def load_config():
    if not os.path.exists(CONFIG_FILE):
        logger.info("config.json not found; using default config.")
        return DEFAULT_CONFIG
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data
    except Exception as e:
        logger.exception("Error loading config.json, using defaults.")
        return DEFAULT_CONFIG

def save_config(config_data):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config_data, f, indent=2)
        logger.info("Configuration saved to config.json")
    except Exception as e:
        logger.exception("Error saving config.json")
