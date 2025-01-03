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
    QFormLayout, QCheckBox
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

def parse_numeric_columns(df, workload_cols, yearly_col):
    """
    Convert the given workload columns and yearly_col to numeric (remove commas, spaces, etc.).
    """
    all_missing = True
    # Convert yearly column
    if yearly_col in df.columns:
        df[yearly_col] = (
            df[yearly_col]
            .astype(str)
            .str.replace(",", "")
            .replace("nan", None)
        )
        df[yearly_col] = pd.to_numeric(df[yearly_col], errors="coerce")
        all_missing = False

    # Convert workload columns
    for wcol in workload_cols:
        if wcol in df.columns:
            df[wcol] = (
                df[wcol].astype(str)
                .str.strip()
                .str.replace(",", "")
                .replace("nan", None)
            )
            df[wcol] = pd.to_numeric(df[wcol], errors="coerce")
            all_missing = False

    return not all_missing  # False means we never found those columns

def check_missing_columns(df, required_cols):
    """
    Return a list of missing columns from 'required_cols' that are not in df.
    """
    missing = [col for col in required_cols if col not in df.columns]
    return missing

def allocate_row_brands(row, brand_cols, workload_cols, yearly_col):
    """
    For advanced partial allocation. Return a list of (brand, allocated_yearly) for this row.
    """
    daily_workloads = []
    total_daily = 0.0
    for bcol, wcol in zip(brand_cols, workload_cols):
        raw_brand = row.get(bcol)
        brand = standardize_brand(raw_brand)
        w = row.get(wcol, 0) or 0
        if brand and w > 0:
            daily_workloads.append((brand, w))
            total_daily += w

    row_yearly = row.get(yearly_col, 0) or 0
    allocated_pairs = []
    if total_daily > 0 and row_yearly > 0:
        for brand, w in daily_workloads:
            proportion = w / total_daily
            allocated = proportion * row_yearly
            allocated_pairs.append((brand, allocated))
    return allocated_pairs

def aggregate_brand_yearly_samples(df, brand_cols, workload_cols, yearly_col):
    """
    Sum up the partial allocations for each brand across all rows.
    Returns {brand: total_annual_samples}.
    """
    logger.info("Aggregating brand yearly samples (partial row allocations).")
    brand_totals = {}

    for _, row in df.iterrows():
        pairs = allocate_row_brands(row, brand_cols, workload_cols, yearly_col)
        for (brand, allocated) in pairs:
            brand_totals[brand] = brand_totals.get(brand, 0) + allocated

    logger.info(f"Finished pro-ration. Brand totals: {brand_totals}")
    return brand_totals

def calculate_market_share(brand_totals):
    """
    Returns {brand: percentage_of_total}.
    """
    total = sum(brand_totals.values())
    if total <= 0:
        logger.warning("Total is zero; cannot calculate market share.")
        return {}
    sorted_items = sorted(brand_totals.items(), key=lambda x: x[1], reverse=True)
    return {b: (val / total) * 100 for b, val in sorted_items}

def add_date_time_columns(df):
    """
    Append 'Date' and 'Time' columns.
    """
    now = datetime.now()
    df["Date"] = now.strftime("%Y-%m-%d")
    df["Time"] = now.strftime("%H:%M:%S")
    return df

def city_pivot_advanced(df, brand_cols, workload_cols, yearly_col):
    """
    Sum partial allocations (yearly) by city. 
    We do row-level brand allocations, then group by city, brand.
    """
    rows = []
    for _, row_data in df.iterrows():
        pairs = allocate_row_brands(row_data, brand_cols, workload_cols, yearly_col)
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

def class_pivot_advanced(df, brand_cols, workload_cols, yearly_col):
    """
    Sum partial allocations (yearly) by class.
    """
    rows = []
    for _, row_data in df.iterrows():
        pairs = allocate_row_brands(row_data, brand_cols, workload_cols, yearly_col)
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

###############################################################################
# MAIN WINDOW (PyQt) WITH ENHANCEMENTS
###############################################################################

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Multi-Analyzer Market Share Analyzer (Enhanced)")
        self.setGeometry(200, 200, 800, 450)

        self.input_file = ""
        self.output_file = ""
        self.dark_mode = False
        self.sheet_names = []  # store sheet names from the chosen file

        widget = QWidget()
        main_layout = QVBoxLayout(widget)

        # Form layout
        form_layout = QFormLayout()

        # 1) Input file
        self.input_edit = QLineEdit()
        btn_input = QPushButton("Browse Input Excel")
        btn_input.clicked.connect(self.browse_input)
        hl_in = QHBoxLayout()
        hl_in.addWidget(self.input_edit)
        hl_in.addWidget(btn_input)
        form_layout.addRow(QLabel("Input File:"), hl_in)

        # 2) Output file
        self.output_edit = QLineEdit()
        btn_output = QPushButton("Browse Output Excel")
        btn_output.clicked.connect(self.browse_output)
        hl_out = QHBoxLayout()
        hl_out.addWidget(self.output_edit)
        hl_out.addWidget(btn_output)
        form_layout.addRow(QLabel("Output File:"), hl_out)

        # 3) Analyzer combo (IA, CBC, CHEM)
        self.combo_analyzer = QComboBox()
        self.combo_analyzer.addItems(["IA", "CBC", "CHEM"])
        self.combo_analyzer.setToolTip("Select which set of columns to use (IA/CBC/CHEM).")
        form_layout.addRow(QLabel("Analyzer Type:"), self.combo_analyzer)

        # 4) Sheet combo (populated after user picks input file)
        self.sheet_combo = QComboBox()
        self.sheet_combo.setToolTip("Select the actual sheet name in the Excel file.")
        form_layout.addRow(QLabel("Select Sheet:"), self.sheet_combo)

        # 5) Region Filter
        self.region_filter_edit = QLineEdit()
        self.region_filter_edit.setPlaceholderText("Type region to filter (optional). E.g. 'SOUTH'")
        form_layout.addRow(QLabel("Region Filter:"), self.region_filter_edit)

        # 6) City/Class check
        self.checkbox_city = QCheckBox("City-wise Pivot (Advanced Partial Allocations)?")
        self.checkbox_class = QCheckBox("Class-wise Pivot (Advanced Partial Allocations)?")
        form_layout.addRow(self.checkbox_city)
        form_layout.addRow(self.checkbox_class)

        main_layout.addLayout(form_layout)

        # Action buttons
        hl_action = QHBoxLayout()
        btn_process = QPushButton("Process Data")
        btn_process.setToolTip("Load data, apply region filter, compute partial allocations, brand share, and save.")
        btn_process.clicked.connect(self.process_data)

        btn_theme = QPushButton("Toggle Theme")
        btn_theme.clicked.connect(self.toggle_theme)

        hl_action.addWidget(btn_process)
        hl_action.addWidget(btn_theme)
        main_layout.addLayout(hl_action)

        self.setCentralWidget(widget)
        self.apply_modern_light_theme()

    def browse_input(self):
        """
        Let user pick an Excel file, then load the sheet names into sheet_combo.
        """
        script_dir = os.path.dirname(__file__)
        data_path = os.path.join(script_dir, "data")
        if not os.path.exists(data_path):
            data_path = script_dir
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Input Excel", data_path, "Excel Files (*.xlsx *.xls)")
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
        file_path, _ = QFileDialog.getSaveFileName(self, "Select Output Excel", data_path, "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.output_file = file_path
            self.output_edit.setText(file_path)
            logger.info(f"Output Excel selected: {file_path}")

    def process_data(self):
        logger.info("Process data initiated.")
        if not self.input_file or not self.output_file:
            QMessageBox.warning(self, "Paths Missing", "Please specify both input and output file paths.")
            return

        chosen_analyzer = self.combo_analyzer.currentText()  # "IA", "CBC", or "CHEM"
        chosen_sheet = self.sheet_combo.currentText()        # actual Excel sheet name

        if not chosen_sheet:
            QMessageBox.warning(self, "Sheet Missing", "Please select a sheet from the dropdown.")
            return

        # Decide columns
        if chosen_analyzer == "IA":
            brand_cols = ["IA Brand 1", "IA Brand 2", "IA Brand 3"]
            workload_cols = ["Workload - Brand 1", "Workload - Brand 2", "Workload - Brand 3"]
            yearly_col = "IA YEARLY SAMPLES"
        elif chosen_analyzer == "CBC":
            brand_cols = ["CBC Brand 1", "CBC Brand 2", "CBC Brand 3", "CBC Brand 4"]
            workload_cols = ["Workload - Brand 1", "Workload - Brand 2", "Workload - Brand 3", "Workload - Brand 4"]
            yearly_col = "HEMATOLOGY  TOTAL YEARLY"
        else:  # CHEM
            brand_cols = ["CHEM Brand 1", "CHEM Brand 2", "CHEM Brand 3", "CHEM Brand 4"]
            workload_cols = ["Workload - Brand 1", "Workload - Brand 2", "Workload - Brand 3", "Workload - Brand 4"]
            yearly_col = "CHEMISTRY TOTAL YEARLY"

        # 1) Load data
        try:
            df = load_data(self.input_file, chosen_sheet)
        except Exception as e:
            logger.exception("Error loading data.")
            QMessageBox.critical(self, "Load Error", str(e))
            return

        # Optional region filter
        region_filter = self.region_filter_edit.text().strip()
        if region_filter:
            if "Region" not in df.columns:
                QMessageBox.warning(self, "No 'Region' Column", 
                                    "You typed a region filter but there is no 'Region' column in the data.")
            else:
                original_len = len(df)
                df = df[df["Region"] == region_filter]
                logger.info(f"Filtered by region '{region_filter}'. Rows from {original_len} -> {len(df)}.")

        # 2) Check required columns
        required_cols = set(brand_cols + workload_cols + [yearly_col])
        missing_cols = check_missing_columns(df, required_cols)
        if missing_cols:
            QMessageBox.critical(self, "Missing Columns", 
                                 f"The following columns are missing from the data:\n{missing_cols}")
            return

        # 3) Parse numeric columns
        parse_numeric_columns(df, workload_cols, yearly_col)

        # 4) Aggregate partial allocations
        brand_totals = aggregate_brand_yearly_samples(df, brand_cols, workload_cols, yearly_col)
        if not brand_totals:
            QMessageBox.information(self, "No Data", f"No valid brand/yearly data found for {chosen_analyzer}.")
            return

        # 5) Market share
        market_share = calculate_market_share(brand_totals)

        # 6) City/Class pivot with advanced partial allocations
        df_city_pivot = None
        if self.checkbox_city.isChecked():
            df_city_pivot = city_pivot_advanced(df, brand_cols, workload_cols, yearly_col)
        df_class_pivot = None
        if self.checkbox_class.isChecked():
            df_class_pivot = class_pivot_advanced(df, brand_cols, workload_cols, yearly_col)

        # Convert results to DataFrame
        df_totals = pd.DataFrame(list(brand_totals.items()), columns=["Brand", "Total Yearly Samples (Pro-Rated)"])
        df_totals = add_date_time_columns(df_totals)

        df_share = pd.DataFrame(list(market_share.items()), columns=["Brand", "Market Share (%)"])
        df_share = add_date_time_columns(df_share)

        # 7) Save
        self.save_results(chosen_analyzer, df_totals, df_share, df_city_pivot, df_class_pivot)

        # 8) Show summary: top brand and # of sites
        top_brand = None
        if market_share:
            top_brand = max(market_share, key=market_share.get)
        num_sites = df["Customer Name"].nunique() if "Customer Name" in df.columns else len(df)
        msg = f"{chosen_analyzer} analysis completed.\n"
        if top_brand:
            msg += f"Top brand is '{top_brand}' at {market_share[top_brand]:.1f}%.\n"
        msg += f"Sites processed: {num_sites}."
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
            QLineEdit, QCheckBox, QComboBox {
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
            QLineEdit, QCheckBox, QComboBox {
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
    logging.info("Launching Multi-Analyzer Market Share App (Enhanced).")
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
