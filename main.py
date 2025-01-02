# main.py
import sys
import os
import logging
import pandas as pd
import numpy as np
from datetime import datetime

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QComboBox, QMessageBox, QLineEdit,
    QFormLayout, QCheckBox, QSpinBox
)
from PyQt5.QtCore import Qt

from config import load_config, save_config
from settings_dialog import SettingsDialog
from aggregator import (
    aggregate_analyzer, calculate_market_share,
    city_pivot_advanced, class_pivot_advanced
)
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

console_handler = logging.StreamHandler(sys.stdout)
console_formatter = logging.Formatter(
    "[%(asctime)s] %(levelname)s - %(name)s : %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)
console_handler.setFormatter(console_formatter)
logger.addHandler(console_handler)

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
        self.setWindowTitle("Multi-Analyzer with Configurable Headers")
        self.setGeometry(200, 200, 900, 600)

        # Load config
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

        # city/class check
        self.checkbox_city = QCheckBox("City Pivot?")
        self.checkbox_class = QCheckBox("Class Pivot?")
        form_layout.addRow(self.checkbox_city)
        form_layout.addRow(self.checkbox_class)

        # Days spin: default from config
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

        hl_bottom.addWidget(btn_settings)
        hl_bottom.addWidget(btn_process)
        hl_bottom.addWidget(btn_theme)

        main_layout.addLayout(hl_bottom)

        self.setCentralWidget(widget)
        self.apply_modern_light_theme()

    def open_settings(self):
        from settings_dialog import SettingsDialog
        dialog = SettingsDialog(self.config_data, self)
        if dialog.exec_():
            # user clicked save
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
        self.config_data["days_per_year"] = days_per_year  # store in memory (and optionally save_config)

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
                logger.info(f"Region filter '{region_filter}': {old_len}->{len(df)} rows")

        analyzer_mode = self.combo_analyzer.currentText()
        if analyzer_mode == "Consolidated":
            self.process_consolidated(df, days_per_year)
        else:
            self.process_single_analyzer(df, analyzer_mode, days_per_year)

    def process_single_analyzer(self, df, analyzer_mode, days_per_year):
        # Load brand_cols/workload_cols from config_data
        brand_cols = self.config_data["headers"][analyzer_mode]["brand_cols"]
        workload_cols = self.config_data["headers"][analyzer_mode]["workload_cols"]

        # Check columns
        missing = check_missing_columns(df, brand_cols + workload_cols)
        if missing:
            QMessageBox.critical(self, "Missing Cols", f"Missing columns for {analyzer_mode}:\n{missing}")
            return

        from aggregator import aggregate_analyzer, calculate_market_share, city_pivot_advanced, class_pivot_advanced

        brand_totals = aggregate_analyzer(df, brand_cols, workload_cols, days_per_year)
        if not brand_totals:
            QMessageBox.information(self, "No Data", f"No valid data for {analyzer_mode}")
            return
        market_share = calculate_market_share(brand_totals)

        df_city = None
        if self.checkbox_city.isChecked():
            df_city = city_pivot_advanced(df, brand_cols, workload_cols, days_per_year)
        df_class = None
        if self.checkbox_class.isChecked():
            df_class = class_pivot_advanced(df, brand_cols, workload_cols, days_per_year)

        from datetime import datetime
        now = datetime.now()

        df_totals = pd.DataFrame(list(brand_totals.items()), columns=["Brand", "Total Yearly"])
        df_share = pd.DataFrame(list(market_share.items()), columns=["Brand", "Market Share (%)"])

        self.save_results(analyzer_mode, df_totals, df_share, df_city, df_class)

        top_brand = max(market_share, key=market_share.get) if market_share else None
        sites = df["Customer Name"].nunique() if "Customer Name" in df.columns else len(df)
        msg = f"{analyzer_mode} done.\n"
        if top_brand:
            msg += f"Top brand: {top_brand} => {market_share[top_brand]:.1f}%\n"
        msg += f"Sites processed: {sites}"
        QMessageBox.information(self, "Done", msg)

    def process_consolidated(self, df, days_per_year):
        from aggregator import aggregate_analyzer, calculate_market_share

        # We'll do IA, CBC, CHEM in one pass:
        for mode in ["IA", "CBC", "CHEM"]:
            brand_cols = self.config_data["headers"][mode]["brand_cols"]
            workload_cols = self.config_data["headers"][mode]["workload_cols"]
            missing = check_missing_columns(df, brand_cols + workload_cols)
            if missing:
                logger.warning(f"Skipping {mode}, missing cols: {missing}")
                continue

            brand_totals = aggregate_analyzer(df, brand_cols, workload_cols, days_per_year)
            if not brand_totals:
                logger.info(f"No data for {mode}, skipping.")
                continue
            market_share = calculate_market_share(brand_totals)

            from aggregator import city_pivot_advanced, class_pivot_advanced

            df_city = None
            if self.checkbox_city.isChecked():
                df_city = city_pivot_advanced(df, brand_cols, workload_cols, days_per_year)
            df_class = None
            if self.checkbox_class.isChecked():
                df_class = class_pivot_advanced(df, brand_cols, workload_cols, days_per_year)

            df_totals = pd.DataFrame(list(brand_totals.items()), columns=["Brand", "Total Yearly"])
            df_share = pd.DataFrame(list(market_share.items()), columns=["Brand", "Market Share (%)"])

            self.save_results(mode, df_totals, df_share, df_city, df_class)

        sites = df["Customer Name"].nunique() if "Customer Name" in df.columns else len(df)
        QMessageBox.information(self, "Done", f"Consolidated run done. Sites processed: {sites}")

    def save_results(self, analyzer_mode, df_totals, df_share, df_city, df_class):
        from openpyxl.utils.exceptions import InvalidFileException
        from openpyxl import load_workbook
        sheet_base = analyzer_mode.upper()

        output_path = self.output_edit.text()
        if os.path.exists(output_path):
            try:
                with pd.ExcelWriter(output_path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
                    df_totals.to_excel(writer, sheet_name=f"{sheet_base}_Totals", index=False)
                    df_share.to_excel(writer, sheet_name=f"{sheet_base}_Share", index=False)
                    if df_city is not None:
                        df_city.to_excel(writer, sheet_name=f"{sheet_base}_CityPivot", index=False)
                    if df_class is not None:
                        df_class.to_excel(writer, sheet_name=f"{sheet_base}_ClassPivot", index=False)
            except InvalidFileException:
                with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                    df_totals.to_excel(writer, sheet_name=f"{sheet_base}_Totals", index=False)
                    df_share.to_excel(writer, sheet_name=f"{sheet_base}_Share", index=False)
                    if df_city is not None:
                        df_city.to_excel(writer, sheet_name=f"{sheet_base}_CityPivot", index=False)
                    if df_class is not None:
                        df_class.to_excel(writer, sheet_name=f"{sheet_base}_ClassPivot", index=False)
        else:
            with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                df_totals.to_excel(writer, sheet_name=f"{sheet_base}_Totals", index=False)
                df_share.to_excel(writer, sheet_name=f"{sheet_base}_Share", index=False)
                if df_city is not None:
                    df_city.to_excel(writer, sheet_name=f"{sheet_base}_CityPivot", index=False)
                if df_class is not None:
                    df_class.to_excel(writer, sheet_name=f"{sheet_base}_ClassPivot", index=False)

    def toggle_theme(self):
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

def main():
    logging.info("Launching Multi-Analyzer with Configurable Headers.")
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
