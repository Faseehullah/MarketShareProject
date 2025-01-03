import sys
import os
import logging
import pandas as pd
import numpy as np
from datetime import datetime
from adjustText import adjust_text

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QComboBox, QMessageBox, QLineEdit,
    QFormLayout, QCheckBox, QSpinBox
)
from PyQt5.QtCore import Qt

###############################################################################
# LOGGING SETUP
###############################################################################
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
console_handler = logging.StreamHandler(sys.stdout)
console_formatter = logging.Formatter(
    "[%(asctime)s] %(levelname)s - %(name)s : %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)
console_handler.setFormatter(console_formatter)
logger.addHandler(console_handler)

###############################################################################
# BUSINESS LOGIC
###############################################################################

def standardize_brand(brand):
    """Convert brand name to uppercase, ignoring NaN or blank/NILL."""
    if pd.isnull(brand):
        return None
    brand_str = str(brand).strip().upper()
    if brand_str in ["NILL", "", "0"]:
        return None
    return brand_str

def load_sheet_names(file_path):
    """
    Return a list of all sheet names in the given Excel file.
    """
    try:
        xls = pd.ExcelFile(file_path)
        return xls.sheet_names
    except Exception as e:
        logger.exception("Error reading sheet names.")
        return []

def load_data(file_path, chosen_sheet):
    """
    Load the chosen_sheet from file_path, replace NILL->NaN.
    """
    logger.info(f"Loading data from {file_path}, sheet: {chosen_sheet}")
    df = pd.read_excel(file_path, sheet_name=chosen_sheet)
    df = df.replace(["NILL", "Nill", "nill"], np.nan)
    return df

def check_missing_columns(df, required_cols):
    """
    Return a list of missing columns from 'required_cols' that are not in df.
    """
    missing = [col for col in required_cols if col not in df.columns]
    return missing

def add_date_time_columns(df):
    """
    Append 'Date' and 'Time' columns.
    """
    now = datetime.now()
    df["Date"] = now.strftime("%Y-%m-%d")
    df["Time"] = now.strftime("%H:%M:%S")
    return df

###############################################################################
# AGGREGATOR: We compute row-level daily sum for the analyzer, multiply by
# 'days_per_year', then do partial allocation by brand daily.
###############################################################################

def allocate_row_brands(row, brand_cols, workload_cols, days_per_year):
    """
    For each row, we sum the daily workloads for the entire analyzer,
    multiply by 'days_per_year' to get total_annual_for_that_row,
    then allocate partial shares per brand.
    Returns a list of (brand, allocated_yearly).
    """
    daily_sum = 0.0
    brand_workloads = []  # (brand, daily_w)

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
    # partial allocation
    allocations = []
    for (brand, w) in brand_workloads:
        proportion = w / daily_sum
        allocated = total_yearly * proportion
        allocations.append((brand, allocated))

    return allocations

def aggregate_analyzer(df, brand_cols, workload_cols, days_per_year):
    """
    Aggregates brand totals for a single analyzer from the consolidated sheet,
    ignoring any 'YEARLY' columns in the data. We compute daily_sum*days and do partial allocations.
    Returns {brand: total_yearly}.
    """
    brand_totals = {}
    for _, row in df.iterrows():
        pairs = allocate_row_brands(row, brand_cols, workload_cols, days_per_year)
        for brand, allocated in pairs:
            brand_totals[brand] = brand_totals.get(brand, 0) + allocated
    return brand_totals

def city_pivot_advanced(df, brand_cols, workload_cols, days_per_year):
    """
    Summation by city using partial daily->yearly allocations.
    """
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
    final_pivot = add_date_time_columns(final_pivot)
    return final_pivot

def class_pivot_advanced(df, brand_cols, workload_cols, days_per_year):
    """
    Summation by class using partial daily->yearly allocations.
    """
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
    final_pivot = add_date_time_columns(final_pivot)
    return final_pivot

def calculate_market_share(brand_totals):
    total = sum(brand_totals.values())
    if total <= 0:
        return {}
    sorted_items = sorted(brand_totals.items(), key=lambda x: x[1], reverse=True)
    return {b: (val / total) * 100 for b, val in sorted_items}

###############################################################################
# MAIN WINDOW (PyQt) 
###############################################################################

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Multi-Analyzer (Consolidated/Individual) Market Share Analyzer")
        self.setGeometry(200, 200, 850, 500)

        self.input_file = ""
        self.output_file = ""
        self.dark_mode = False
        self.sheet_names = []

        widget = QWidget()
        main_layout = QVBoxLayout(widget)

        form_layout = QFormLayout()

        # 1) Input File
        self.input_edit = QLineEdit()
        btn_input = QPushButton("Browse Input Excel")
        btn_input.clicked.connect(self.browse_input)
        hl_in = QHBoxLayout()
        hl_in.addWidget(self.input_edit)
        hl_in.addWidget(btn_input)
        form_layout.addRow(QLabel("Input File:"), hl_in)

        # 2) Output File
        self.output_edit = QLineEdit()
        btn_output = QPushButton("Browse Output Excel")
        btn_output.clicked.connect(self.browse_output)
        hl_out = QHBoxLayout()
        hl_out.addWidget(self.output_edit)
        hl_out.addWidget(btn_output)
        form_layout.addRow(QLabel("Output File:"), hl_out)

        # 3) Analyzer: add "Consolidated" option
        self.combo_analyzer = QComboBox()
        self.combo_analyzer.addItems(["IA", "CBC", "CHEM", "Consolidated"])
        self.combo_analyzer.setToolTip("Select 'IA/CBC/CHEM' or 'Consolidated' for single sheet with all columns.")
        form_layout.addRow(QLabel("Analyzer Mode:"), self.combo_analyzer)

        # 4) Sheet combo
        self.sheet_combo = QComboBox()
        self.sheet_combo.setToolTip("Select the sheet in the Excel file.")
        form_layout.addRow(QLabel("Select Sheet:"), self.sheet_combo)

        # 5) Region Filter
        self.region_filter_edit = QLineEdit()
        self.region_filter_edit.setPlaceholderText("Type region to filter (optional). E.g. 'SOUTH'")
        form_layout.addRow(QLabel("Region Filter:"), self.region_filter_edit)

        # 6) City/Class check
        self.checkbox_city = QCheckBox("City-wise Pivot?")
        self.checkbox_class = QCheckBox("Class-wise Pivot?")
        form_layout.addRow(self.checkbox_city)
        form_layout.addRow(self.checkbox_class)

        # 7) Days per Year
        self.days_label = QLabel("Days per Year:")
        self.days_spin = QSpinBox()
        self.days_spin.setRange(1, 1000)
        self.days_spin.setValue(330)  # default
        form_layout.addRow(self.days_label, self.days_spin)

        main_layout.addLayout(form_layout)

        # Action buttons
        hl_action = QHBoxLayout()
        btn_process = QPushButton("Process Data")
        btn_process.clicked.connect(self.process_data)
        btn_theme = QPushButton("Toggle Theme")
        btn_theme.clicked.connect(self.toggle_theme)
        hl_action.addWidget(btn_process)
        hl_action.addWidget(btn_theme)
        main_layout.addLayout(hl_action)

        self.setCentralWidget(widget)
        self.apply_modern_light_theme()

    def browse_input(self):
        script_dir = os.path.dirname(__file__)
        data_path = os.path.join(script_dir, "data")
        if not os.path.exists(data_path):
            data_path = script_dir
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Input Excel", data_path, "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.input_file = file_path
            self.input_edit.setText(file_path)
            logger.info(f"Input Excel selected: {file_path}")

            # load sheet names
            self.sheet_names = load_sheet_names(file_path)
            self.sheet_combo.clear()
            self.sheet_combo.addItems(self.sheet_names)

    def browse_output(self):
        script_dir = os.path.dirname(__file__)
        data_path = os.path.join(script_dir, "data")
        if not os.path.exists(data_path):
            data_path = script_dir
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Select Output Excel", data_path, "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.output_file = file_path
            self.output_edit.setText(file_path)
            logger.info(f"Output Excel selected: {file_path}")

    def process_data(self):
        logger.info("Process data initiated.")
        if not self.input_file or not self.output_file:
            QMessageBox.warning(self, "Paths Missing", "Please specify both input and output file paths.")
            return

        chosen_sheet = self.sheet_combo.currentText()
        if not chosen_sheet:
            QMessageBox.warning(self, "Sheet Missing", "Please select a sheet from the dropdown.")
            return

        # Load data
        try:
            df = load_data(self.input_file, chosen_sheet)
        except Exception as e:
            logger.exception("Error loading data.")
            QMessageBox.critical(self, "Load Error", str(e))
            return

        # Region filter
        region_filter = self.region_filter_edit.text().strip()
        if region_filter:
            if "Region" not in df.columns:
                QMessageBox.warning(self, "No 'Region' Column",
                                    "You typed a region filter but there is no 'Region' column in the data.")
            else:
                original_len = len(df)
                df = df[df["Region"] == region_filter]
                logger.info(f"Filtered by region '{region_filter}': {original_len} -> {len(df)} rows")

        # Days per year
        days_per_year = self.days_spin.value()
        analyzer_mode = self.combo_analyzer.currentText()  # "IA", "CBC", "CHEM", "Consolidated"

        if analyzer_mode == "Consolidated":
            # We do a single sheet that presumably has ALL columns for IA, CBC, CHEM
            self.process_consolidated(df, days_per_year)
        else:
            # Just do a single analyzer
            self.process_single_analyzer(df, analyzer_mode, days_per_year)

    def process_single_analyzer(self, df, analyzer_mode, days_per_year):
        """
        Process either IA, CBC, or CHEM from the single sheet.
            """
    def process_single_analyzer(self, df, analyzer_mode, days_per_year):
        if analyzer_mode == "IA":
            brand_cols = ["IA Brand 1", "IA Brand 2", "IA Brand 3"]
            workload_cols = ["IA Workload - Brand 1", "IA Workload - Brand 2", "IA Workload - Brand 3"]
        elif analyzer_mode == "CBC":
            brand_cols = ["CBC Brand 1", "CBC Brand 2", "CBC Brand 3", "CBC Brand 4"]
            workload_cols = ["CBC Workload - Brand 1", "CBC Workload - Brand 2",
                            "CBC Workload - Brand 3", "CBC Workload - Brand 4"]
        else:  # "CHEM"
            brand_cols = ["CHEM Brand 1", "CHEM Brand 2", "CHEM Brand 3", "CHEM Brand 4"]
            workload_cols = ["CHEM Workload - Brand 1", "CHEM Workload - Brand 2",
                            "CHEM Workload - Brand 3", "CHEM Workload - Brand 4"]
        # check required columns
        required_cols = set(brand_cols + workload_cols)
        missing = check_missing_columns(df, required_cols)
        if missing:
            QMessageBox.critical(self, "Missing Columns",
                                 f"Missing columns for {analyzer_mode}:\n{missing}")
            return

        # aggregator
        brand_totals = aggregate_analyzer(df, brand_cols, workload_cols, days_per_year)
        if not brand_totals:
            QMessageBox.information(self, "No Data", f"No valid data found for {analyzer_mode}.")
            return

        # market share
        market_share = calculate_market_share(brand_totals)

        # city/class pivot
        df_city_pivot = None
        df_class_pivot = None
        if self.checkbox_city.isChecked():
            df_city_pivot = city_pivot_advanced(df, brand_cols, workload_cols, days_per_year)
        if self.checkbox_class.isChecked():
            df_class_pivot = class_pivot_advanced(df, brand_cols, workload_cols, days_per_year)

        # Convert to DataFrame
        df_totals = pd.DataFrame(list(brand_totals.items()), columns=["Brand", "Total Yearly (computed)"])
        df_totals = add_date_time_columns(df_totals)

        df_share = pd.DataFrame(list(market_share.items()), columns=["Brand", "Market Share (%)"])
        df_share = add_date_time_columns(df_share)

        # Save
        self.save_results(analyzer_mode, df_totals, df_share, df_city_pivot, df_class_pivot)

        # summary
        top_brand = max(market_share, key=market_share.get) if market_share else None
        num_sites = df["Customer Name"].nunique() if "Customer Name" in df.columns else len(df)
        msg = f"{analyzer_mode} analysis completed.\n"
        if top_brand:
            msg += f"Top brand: '{top_brand}' = {market_share[top_brand]:.1f}%.\n"
        msg += f"Sites processed: {num_sites}."
        QMessageBox.information(self, "Done", msg)
        logger.info(msg)

    def process_consolidated(self, df, days_per_year):
        analyzers = {
            "IA": {
                "brand_cols": ["IA Brand 1", "IA Brand 2", "IA Brand 3"],
                "workload_cols": ["IA Workload - Brand 1", "IA Workload - Brand 2", "IA Workload - Brand 3"],
            },
            "CBC": {
                "brand_cols": ["CBC Brand 1", "CBC Brand 2", "CBC Brand 3", "CBC Brand 4"],
                "workload_cols": ["CBC Workload - Brand 1", "CBC Workload - Brand 2",
                                "CBC Workload - Brand 3", "CBC Workload - Brand 4"],
            },
            "CHEM": {
                "brand_cols": ["CHEM Brand 1", "CHEM Brand 2", "CHEM Brand 3", "CHEM Brand 4"],
                "workload_cols": ["CHEM Workload - Brand 1", "CHEM Workload - Brand 2",
                                "CHEM Workload - Brand 3", "CHEM Workload - Brand 4"],
            }
        }

        # We'll process each analyzer in turn
        for analyzer_mode, colmap in analyzers.items():
            brand_cols = colmap["brand_cols"]
            workload_cols = colmap["workload_cols"]
            required_cols = set(brand_cols + workload_cols)
            missing = check_missing_columns(df, required_cols)
            if missing:
                logger.warning(f"Skipping {analyzer_mode} due to missing columns: {missing}")
                # we skip, or we can continue to next
                continue

            brand_totals = aggregate_analyzer(df, brand_cols, workload_cols, days_per_year)
            if not brand_totals:
                logger.info(f"No data for {analyzer_mode} in the consolidated sheet.")
                continue
            market_share = calculate_market_share(brand_totals)

            df_city_pivot = None
            if self.checkbox_city.isChecked():
                df_city_pivot = city_pivot_advanced(df, brand_cols, workload_cols, days_per_year)
            df_class_pivot = None
            if self.checkbox_class.isChecked():
                df_class_pivot = class_pivot_advanced(df, brand_cols, workload_cols, days_per_year)

            df_totals = pd.DataFrame(list(brand_totals.items()), columns=["Brand", "Total Yearly (computed)"])
            df_totals = add_date_time_columns(df_totals)

            df_share = pd.DataFrame(list(market_share.items()), columns=["Brand", "Market Share (%)"])
            df_share = add_date_time_columns(df_share)

            self.save_results(analyzer_mode, df_totals, df_share, df_city_pivot, df_class_pivot)

        # We can also show a final summary
        num_sites = df["Customer Name"].nunique() if "Customer Name" in df.columns else len(df)
        msg = f"Consolidated analysis done for IA, CBC, CHEM.\nSites processed: {num_sites}."
        QMessageBox.information(self, "Done", msg)
        logger.info(msg)

    def save_results(self, analyzer, df_totals, df_share, df_city, df_class):
        from openpyxl.utils.exceptions import InvalidFileException
        from openpyxl import load_workbook

        sheet_base = analyzer.upper()

        if os.path.exists(self.output_file):
            try:
                with pd.ExcelWriter(self.output_file, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
                    df_totals.to_excel(writer, sheet_name=f"{sheet_base}_Totals", index=False)
                    df_share.to_excel(writer, sheet_name=f"{sheet_base}_Share", index=False)
                    if df_city is not None:
                        df_city.to_excel(writer, sheet_name=f"{sheet_base}_CityPivot", index=False)
                    if df_class is not None:
                        df_class.to_excel(writer, sheet_name=f"{sheet_base}_ClassPivot", index=False)
            except InvalidFileException:
                # If file is locked or invalid, we overwrite
                with pd.ExcelWriter(self.output_file, engine="openpyxl") as writer:
                    df_totals.to_excel(writer, sheet_name=f"{sheet_base}_Totals", index=False)
                    df_share.to_excel(writer, sheet_name=f"{sheet_base}_Share", index=False)
                    if df_city is not None:
                        df_city.to_excel(writer, sheet_name=f"{sheet_base}_CityPivot", index=False)
                    if df_class is not None:
                        df_class.to_excel(writer, sheet_name=f"{sheet_base}_ClassPivot", index=False)
        else:
            with pd.ExcelWriter(self.output_file, engine="openpyxl") as writer:
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
    logging.info("Launching Multi-Analyzer (Consolidated/Individual) Market Share App.")
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
