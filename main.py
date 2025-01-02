# main.py

import sys
import os
import logging
import subprocess
import pandas as pd
import numpy as np
from datetime import datetime

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QComboBox, QMessageBox, QLineEdit,
    QFormLayout, QCheckBox, QSpinBox, QInputDialog
)
from PyQt5.QtCore import Qt

from config import load_config, save_config
from settings_dialog import SettingsDialog
from modern_dashboard import ModernDataAnalysisApp

##############################################################
# AGGREGATOR LOGIC
##############################################################

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

console_handler = logging.StreamHandler(sys.stdout)
console_formatter = logging.Formatter(
    "[%(asctime)s] %(levelname)s - %(name)s : %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)
console_handler.setFormatter(console_formatter)
logger.addHandler(console_handler)

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

##############################################################
# END aggregator logic
##############################################################

def load_sheet_names(file_path):
    try:
        xls = pd.ExcelFile(file_path)
        return xls.sheet_names
    except:
        logger.exception("Error reading sheet names.")
        return []

def load_data(file_path, sheet_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    df = df.replace(["NILL", "Nill", "nill"], np.nan)
    return df

def check_missing_columns(df, required_cols):
    missing = [col for col in required_cols if col not in df.columns]
    return missing

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Multi-Analyzer with Configurable Headers + Git Integration")
        self.setGeometry(200, 200, 900, 600)

        # Load config from config.json
        self.config_data = load_config()

        widget = QWidget()
        main_layout = QVBoxLayout(widget)

        form_layout = QFormLayout()

        # Input file
        self.input_edit = QLineEdit()
        btn_input = QPushButton("Browse Input Excel")
        btn_input.clicked.connect(self.browse_input)
        hl_in = QHBoxLayout()
        hl_in.addWidget(self.input_edit)
        hl_in.addWidget(btn_input)
        form_layout.addRow(QLabel("Input File:"), hl_in)

        # Output file
        self.output_edit = QLineEdit()
        btn_output = QPushButton("Browse Output Excel")
        btn_output.clicked.connect(self.browse_output)
        hl_out = QHBoxLayout()
        hl_out.addWidget(self.output_edit)
        hl_out.addWidget(btn_output)
        form_layout.addRow(QLabel("Output File:"), hl_out)

        # Analyzer combo
        self.combo_analyzer = QComboBox()
        self.combo_analyzer.addItems(["IA", "CBC", "CHEM", "Consolidated"])
        form_layout.addRow(QLabel("Analyzer Mode:"), self.combo_analyzer)

        # Sheet combo
        self.sheet_combo = QComboBox()
        form_layout.addRow(QLabel("Excel Sheet:"), self.sheet_combo)

        # Region filter
        self.region_filter_edit = QLineEdit()
        self.region_filter_edit.setPlaceholderText("Optional region filter...")
        form_layout.addRow(QLabel("Region Filter:"), self.region_filter_edit)

        # City/Class check
        self.checkbox_city = QCheckBox("City Pivot?")
        self.checkbox_class = QCheckBox("Class Pivot?")
        form_layout.addRow(self.checkbox_city)
        form_layout.addRow(self.checkbox_class)

        # Days spin
        self.days_label = QLabel("Days per Year:")
        self.days_spin = QSpinBox()
        self.days_spin.setRange(1, 1000)
        self.days_spin.setValue(self.config_data.get("days_per_year", 330))
        form_layout.addRow(self.days_label, self.days_spin)

        main_layout.addLayout(form_layout)

        # Action buttons
        hl_bottom = QHBoxLayout()

        btn_settings = QPushButton("Settings")
        btn_settings.clicked.connect(self.open_settings)

        btn_process = QPushButton("Process")
        btn_process.clicked.connect(self.process_data)

        btn_theme = QPushButton("Toggle Theme")
        btn_theme.clicked.connect(self.toggle_theme)

        btn_push_git = QPushButton("Push to GitHub")
        btn_push_git.setToolTip("Commits local changes and pushes to the 'origin' remote in Git repository.")
        btn_push_git.clicked.connect(self.push_to_github)

        self.dashboard_button = QPushButton("Open Data Visualization Dashboard")
        self.dashboard_button.clicked.connect(self.open_dashboard)
        main_layout.addWidget(self.dashboard_button)

        hl_bottom.addWidget(btn_settings)
        hl_bottom.addWidget(btn_process)
        hl_bottom.addWidget(btn_theme)
        hl_bottom.addWidget(btn_push_git)

        main_layout.addLayout(hl_bottom)

        self.setCentralWidget(widget)
        self.apply_modern_light_theme()

    def open_dashboard(self):
            """Launch the modern data analysis dashboard (modern_dashboard.py)."""
            self.dash_window = ModernDataAnalysisApp()
            self.dash_window.show()

    def open_settings(self):
        from settings_dialog import SettingsDialog
        dialog = SettingsDialog(self.config_data, self)
        if dialog.exec_():
            self.config_data = load_config()  # reload from file
            self.days_spin.setValue(self.config_data.get("days_per_year", 330))
            QMessageBox.information(self, "Settings Saved", "Settings updated successfully.")

    def browse_input(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Input Excel", "", "Excel Files (*.xlsx *.xls)")
        if path:
            self.input_file = path
            self.input_edit.setText(path)
            sheets = load_sheet_names(path)
            self.sheet_combo.clear()
            self.sheet_combo.addItems(sheets)

    def browse_output(self):
        path, _ = QFileDialog.getSaveFileName(self, "Select Output Excel", "", "Excel Files (*.xlsx *.xls)")
        if path:
            self.output_file = path
            self.output_edit.setText(path)

    def process_data(self):
        if not self.input_edit.text() or not self.output_edit.text():
            QMessageBox.warning(self, "Error", "Please select input and output files.")
            return

        chosen_sheet = self.sheet_combo.currentText()
        if not chosen_sheet:
            QMessageBox.warning(self, "Error", "Please pick a valid sheet.")
            return

        region_filter = self.region_filter_edit.text().strip()
        days_per_year = self.days_spin.value()
        self.config_data["days_per_year"] = days_per_year

        # Load data
        try:
            df = load_data(self.input_edit.text(), chosen_sheet)
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))
            return

        # Region filter
        if region_filter:
            if "Region" not in df.columns:
                QMessageBox.warning(self, "Warning", "No 'Region' column in data.")
            else:
                old_len = len(df)
                df = df[df["Region"] == region_filter]
                logger.info(f"Region filter '{region_filter}': {old_len} -> {len(df)} rows")

        analyzer_mode = self.combo_analyzer.currentText()
        if analyzer_mode == "Consolidated":
            self.process_consolidated(df, days_per_year)
        else:
            self.process_single_analyzer(df, analyzer_mode, days_per_year)

    def process_single_analyzer(self, df, analyzer_mode, days_per_year):
        brand_cols = self.config_data["headers"][analyzer_mode]["brand_cols"]
        workload_cols = self.config_data["headers"][analyzer_mode]["workload_cols"]

        missing = check_missing_columns(df, brand_cols + workload_cols)
        if missing:
            QMessageBox.critical(self, "Missing Cols", f"Missing columns for {analyzer_mode}:\n{missing}")
            return

        # 1) volume aggregator
        brand_totals = aggregate_analyzer(df, brand_cols, workload_cols, days_per_year)
        if not brand_totals:
            QMessageBox.information(self, "No Data", f"No valid data for {analyzer_mode}")
            return
        market_share = calculate_market_share(brand_totals)

        # 2) value aggregator (cost-based)
        cost_per_test = float(self.config_data["cost_per_test"].get(analyzer_mode, 0))
        brand_values = {b: (brand_totals[b] * cost_per_test) for b in brand_totals}
        market_share_value = calculate_market_share(brand_values)

        # city/class pivot (volume-based)
        df_city = city_pivot_advanced(df, brand_cols, workload_cols, days_per_year) if self.checkbox_city.isChecked() else None
        df_class = class_pivot_advanced(df, brand_cols, workload_cols, days_per_year) if self.checkbox_class.isChecked() else None

        # Make a combined single sheet with columns for volume & share, plus value & share
        # Brand, Volume, VolumeShare, Value, ValueShare
        df_combined = self.build_combined_df(brand_totals, market_share, brand_values, market_share_value)

        # Now save everything to just 1 sheet for the analyzer + optional pivot sheets
        self.save_results_one_sheet(analyzer_mode, df_combined, df_city, df_class)

        # Show summary
        top_brand = max(market_share, key=market_share.get) if market_share else None
        sites = df["Customer Name"].nunique() if "Customer Name" in df.columns else len(df)
        msg = f"{analyzer_mode} done.\n"
        if top_brand:
            msg += f"Top brand (Volume): {top_brand} => {market_share[top_brand]:.1f}%\n"
        if market_share_value:
            top_brand_val = max(market_share_value, key=market_share_value.get)
            msg += f"Top brand (Value): {top_brand_val} => {market_share_value[top_brand_val]:.1f}%\n"
        msg += f"Sites processed: {sites}"
        QMessageBox.information(self, "Done", msg)

    def process_consolidated(self, df, days_per_year):
        """
        For each analyzer, compute volume & value, but store each in a single sheet
        e.g. 'IA', 'CBC', 'CHEM'. Optionally city/class pivot in separate sheets if needed.
        """
        for mode in ["IA", "CBC", "CHEM"]:
            brand_cols = self.config_data["headers"][mode]["brand_cols"]
            workload_cols = self.config_data["headers"][mode]["workload_cols"]

            missing = check_missing_columns(df, brand_cols + workload_cols)
            if missing:
                logger.warning(f"Skipping {mode}, missing columns: {missing}")
                continue

            brand_totals = aggregate_analyzer(df, brand_cols, workload_cols, days_per_year)
            if not brand_totals:
                logger.info(f"No data for {mode}, skipping.")
                continue
            market_share = calculate_market_share(brand_totals)

            cost_per_test = float(self.config_data["cost_per_test"].get(mode, 0))
            brand_values = {b: (brand_totals[b] * cost_per_test) for b in brand_totals}
            market_share_value = calculate_market_share(brand_values)

            df_city = city_pivot_advanced(df, brand_cols, workload_cols, days_per_year) if self.checkbox_city.isChecked() else None
            df_class = class_pivot_advanced(df, brand_cols, workload_cols, days_per_year) if self.checkbox_class.isChecked() else None

            df_combined = self.build_combined_df(brand_totals, market_share, brand_values, market_share_value)

            self.save_results_one_sheet(mode, df_combined, df_city, df_class)

        sites = df["Customer Name"].nunique() if "Customer Name" in df.columns else len(df)
        QMessageBox.information(self, "Done", f"Consolidated run done. Sites processed: {sites}")

    def build_combined_df(self, brand_totals, market_share, brand_values, market_share_value):
        """
        Merge volume totals & shares with value totals & shares into a single DataFrame
        with columns: Brand, Volume, VolumeShare, Value, ValueShare
        """

        # Convert dict -> DataFrame
        df_vol = pd.DataFrame(list(brand_totals.items()), columns=["Brand", "Volume"])
        df_vol_share = pd.DataFrame(list(market_share.items()), columns=["Brand", "VolumeShare"])

        df_val = pd.DataFrame(list(brand_values.items()), columns=["Brand", "Value"])
        df_val_share = pd.DataFrame(list(market_share_value.items()), columns=["Brand", "ValueShare"])

        # Merge them all on Brand
        df_merge = pd.merge(df_vol, df_vol_share, on="Brand", how="outer")
        df_merge = pd.merge(df_merge, df_val, on="Brand", how="outer")
        df_merge = pd.merge(df_merge, df_val_share, on="Brand", how="outer")

        # Round columns if needed
        df_merge["Volume"] = df_merge["Volume"].round(1)
        df_merge["VolumeShare"] = df_merge["VolumeShare"].round(1)
        df_merge["Value"] = df_merge["Value"].round(1)
        df_merge["ValueShare"] = df_merge["ValueShare"].round(1)

        return df_merge

    def save_results_one_sheet(self, analyzer_mode, df_combined, df_city, df_class):
        """
        Writes a single sheet named e.g. 'IA' for volume+value results,
        plus optional 'IA_CityPivot', 'IA_ClassPivot' for pivot data.
        """
        from openpyxl.utils.exceptions import InvalidFileException
        from openpyxl import load_workbook

        sheet_base = analyzer_mode.upper()
        output_path = self.output_edit.text()

        if os.path.exists(output_path):
            try:
                with pd.ExcelWriter(output_path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
                    df_combined.to_excel(writer, sheet_name=sheet_base, index=False)

                    if df_city is not None:
                        df_city.to_excel(writer, sheet_name=f"{sheet_base}_CityPivot", index=False)
                    if df_class is not None:
                        df_class.to_excel(writer, sheet_name=f"{sheet_base}_ClassPivot", index=False)
            except InvalidFileException:
                with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                    df_combined.to_excel(writer, sheet_name=sheet_base, index=False)

                    if df_city is not None:
                        df_city.to_excel(writer, sheet_name=f"{sheet_base}_CityPivot", index=False)
                    if df_class is not None:
                        df_class.to_excel(writer, sheet_name=f"{sheet_base}_ClassPivot", index=False)
        else:
            with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                df_combined.to_excel(writer, sheet_name=sheet_base, index=False)
                if df_city is not None:
                    df_city.to_excel(writer, sheet_name=f"{sheet_base}_CityPivot", index=False)
                if df_class is not None:
                    df_class.to_excel(writer, sheet_name=f"{sheet_base}_ClassPivot", index=False)

    def toggle_theme(self):
        if not hasattr(self, 'dark_mode'):
            self.dark_mode = False
        if not self.dark_mode:
            self.apply_modern_dark_theme()
            self.dark_mode = True
        else:
            self.apply_modern_light_theme()
            self.dark_mode = False

    def apply_modern_light_theme(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: #fafafa;
            }
            QLabel {
                color: #333333;
                font: 13px 'Segoe UI';
            }
            QLineEdit, QCheckBox, QComboBox, QSpinBox {
                color: #333333;
                background-color: #ffffff;
                border: 1px solid #cccccc;
                font: 12px 'Segoe UI';
            }
            QPushButton {
                color: #333333;
                background-color: #e0f7fa;
                border: 1px solid #b2ebf2;
                padding: 5px;
                font: bold 12px 'Segoe UI';
            }
            QPushButton:hover {
                background-color: #b2ebf2;
                border: 1px solid #80deea;
            }
        """)

    def apply_modern_dark_theme(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: #303030;
            }
            QLabel {
                color: #ffffff;
                font: 13px 'Segoe UI';
            }
            QLineEdit, QCheckBox, QComboBox, QSpinBox {
                color: #ffffff;
                background-color: #424242;
                border: 1px solid #666666;
                font: 12px 'Segoe UI';
            }
            QPushButton {
                color: #ffffff;
                background-color: #006064;
                border: 1px solid #004d40;
                padding: 5px;
                font: bold 12px 'Segoe UI';
            }
            QPushButton:hover {
                background-color: #00838f;
                border: 1px solid #006064;
            }
        """)

    def push_to_github(self):
        commit_msg, ok = QInputDialog.getText(
            self,
            "Commit Message",
            "Enter commit message:"
        )
        if not ok or not commit_msg.strip():
            QMessageBox.information(self, "No Commit", "Commit canceled (no message).")
            return

        project_dir = os.path.dirname(__file__)

        try:
            subprocess.run(["git", "add", "."], cwd=project_dir, check=True)
            subprocess.run(["git", "commit", "-m", commit_msg.strip()], cwd=project_dir, check=True)
            subprocess.run(["git", "push", "origin", "main"], cwd=project_dir, check=True)
            QMessageBox.information(self, "Git", "Changes pushed to GitHub successfully!")
        except subprocess.CalledProcessError as e:
            logger.exception("Git command failed")
            QMessageBox.critical(self, "Git Error", f"Git command failed:\n{str(e)}")
        except Exception as e:
            logger.exception("Unexpected error with Git push")
            QMessageBox.critical(self, "Git Error", f"Unexpected error:\n{str(e)}")

def main():
    print("DEBUG: About to create QApplication")
    logging.info("Launching Multi-Analyzer with Configurable Headers + Git Integration.")
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
