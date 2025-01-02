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

def load_data(file_path, sheet_name):
    """
    Generic loader that replaces NILL->NaN and attempts to parse 
    yearly and workload columns as numeric.
    """
    logger.info(f"Loading {sheet_name} data from {file_path}")
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    df = df.replace(["NILL", "Nill", "nill"], np.nan)
    return df

def parse_numeric_columns(df, workload_cols, yearly_col):
    """
    Convert the given workload columns and yearly_col to numeric (remove commas, spaces, etc.).
    """
    # Convert yearly column
    if yearly_col in df.columns:
        df[yearly_col] = (
            df[yearly_col]
            .astype(str)
            .str.replace(",", "")
            .replace("nan", None)
        )
        df[yearly_col] = pd.to_numeric(df[yearly_col], errors="coerce")

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

def aggregate_brand_yearly_samples(df, brand_cols, workload_cols, yearly_col):
    """
    Pro-rate 'yearly_col' among the brand_cols in proportion to workload_cols.
    Returns {brand: total_annual_samples}.
    """
    logger.info("Aggregating brand yearly samples (pro-ration).")
    brand_totals = {}

    for _, row in df.iterrows():
        daily_workloads = []
        total_daily = 0.0
        for bcol, wcol in zip(brand_cols, workload_cols):
            if bcol not in df.columns or wcol not in df.columns:
                continue
            raw_brand = row.get(bcol)
            brand = standardize_brand(raw_brand)
            w = row.get(wcol, 0) or 0
            if brand and w > 0:
                daily_workloads.append((brand, w))
                total_daily += w

        row_yearly = row.get(yearly_col, 0) or 0
        if total_daily > 0 and row_yearly > 0:
            for brand, w in daily_workloads:
                proportion = w / total_daily
                allocated = proportion * row_yearly
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

def approximate_yearly_from_daily(sum_of_daily):
    """
    Per your requirement: multiply daily sum by 330 to get approximate yearly.
    """
    return sum_of_daily * 330

###############################################################################
# NAIVE PIVOTING (City/Class) with 330 multiplier
###############################################################################

def city_pivot_approx(df, brand_cols, workload_cols):
    """
    For each row, sum daily workloads by brand, then multiply by 330 to get yearly approx.
    Then pivot by city.
    """
    rows = []
    for _, row_data in df.iterrows():
        city = row_data.get("CITY", "UNKNOWN")
        for bcol, wcol in zip(brand_cols, workload_cols):
            brand = standardize_brand(row_data.get(bcol))
            w = row_data.get(wcol, 0) or 0
            if brand and w > 0:
                rows.append({"CITY": city, "BRAND": brand, "DAILY_WORKLOAD": w})

    if not rows:
        return None
    city_df = pd.DataFrame(rows)
    # sum daily workloads
    pivoted = city_df.groupby(["CITY", "BRAND"])["DAILY_WORKLOAD"].sum().reset_index()
    # convert daily to approx yearly
    pivoted["YearlyApprox"] = pivoted["DAILY_WORKLOAD"].apply(approximate_yearly_from_daily)

    # pivot
    final_pivot = pivoted.pivot(index="CITY", columns="BRAND", values="YearlyApprox").fillna(0)
    final_pivot.reset_index(inplace=True)
    final_pivot = add_date_time_columns(final_pivot)
    return final_pivot

def class_pivot_approx(df, brand_cols, workload_cols):
    """
    Same approach as city pivot, but grouping by 'Class'.
    """
    rows = []
    for _, row_data in df.iterrows():
        clss = row_data.get("Class", "UNKNOWN")
        for bcol, wcol in zip(brand_cols, workload_cols):
            brand = standardize_brand(row_data.get(bcol))
            w = row_data.get(wcol, 0) or 0
            if brand and w > 0:
                rows.append({"CLASS": clss, "BRAND": brand, "DAILY_WORKLOAD": w})

    if not rows:
        return None
    class_df = pd.DataFrame(rows)
    pivoted = class_df.groupby(["CLASS", "BRAND"])["DAILY_WORKLOAD"].sum().reset_index()
    pivoted["YearlyApprox"] = pivoted["DAILY_WORKLOAD"].apply(approximate_yearly_from_daily)

    final_pivot = pivoted.pivot(index="CLASS", columns="BRAND", values="YearlyApprox").fillna(0)
    final_pivot.reset_index(inplace=True)
    final_pivot = add_date_time_columns(final_pivot)
    return final_pivot

###############################################################################
# MAIN WINDOW (PyQt)
###############################################################################

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Multi-Analyzer Market Share Analyzer")
        self.setGeometry(200, 200, 650, 400)

        self.input_file = ""
        self.output_file = ""
        self.dark_mode = False

        widget = QWidget()
        main_layout = QVBoxLayout(widget)

        # Form layout
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
        self.combo_analyzer.addItems(["IA", "CBC", "CHEM"])
        self.combo_analyzer.setToolTip("Select which sheet/analyzer to process.")
        form_layout.addRow(QLabel("Analyzer Type:"), self.combo_analyzer)

        # City/Class check
        self.checkbox_city = QCheckBox("City-wise Pivot (x330)?")
        self.checkbox_class = QCheckBox("Class-wise Pivot (x330)?")
        form_layout.addRow(self.checkbox_city)
        form_layout.addRow(self.checkbox_class)

        main_layout.addLayout(form_layout)

        # Action buttons
        hl_action = QHBoxLayout()
        btn_process = QPushButton("Process Data")
        btn_process.setToolTip("Load data from the chosen analyzer, compute share, and save results.")
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
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Input Excel", data_path, "Excel Files (*.xlsx *.xls)")
        if file_path:
            self.input_file = file_path
            self.input_edit.setText(file_path)
            logger.info(f"Input Excel selected: {file_path}")

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
        
        analyzer = self.combo_analyzer.currentText()
        sheet_name = analyzer  # We assume sheet names are exactly "IA", "CBC", "CHEM"

        # Decide columns
        if analyzer == "IA":
            brand_cols = ["IA Brand 1", "IA Brand 2", "IA Brand 3"]
            workload_cols = ["Workload - Brand 1", "Workload - Brand 2", "Workload - Brand 3"]
            yearly_col = "IA YEARLY SAMPLES"
        elif analyzer == "CBC":
            brand_cols = ["CBC Brand 1", "CBC Brand 2", "CBC Brand 3", "CBC Brand 4"]
            workload_cols = ["Workload - Brand 1", "Workload - Brand 2", "Workload - Brand 3", "Workload - Brand 4"]
            yearly_col = "HEMATOLOGY  TOTAL YEARLY"
        else:  # CHEM
            brand_cols = ["CHEM Brand 1", "CHEM Brand 2", "CHEM Brand 3", "CHEM Brand 4"]
            workload_cols = ["Workload - Brand 1", "Workload - Brand 2", "Workload - Brand 3", "Workload - Brand 4"]
            yearly_col = "CHEMISTRY TOTAL YEARLY"

        # 1) Load data
        try:
            df = load_data(self.input_file, sheet_name)
        except ValueError as ve:
            logger.exception("Sheet not found or invalid.")
            QMessageBox.critical(self, "Error", f"Sheet '{sheet_name}' not found in file.\n{ve}")
            return
        except Exception as e:
            logger.exception("Error loading data.")
            QMessageBox.critical(self, "Error", str(e))
            return

        # 2) Parse numeric columns
        parse_numeric_columns(df, workload_cols, yearly_col)

        # 3) Pro-rate
        brand_totals = aggregate_brand_yearly_samples(df, brand_cols, workload_cols, yearly_col)
        if not brand_totals:
            QMessageBox.information(self, "No Data", f"No valid brand/yearly data found for {analyzer}.")
            return

        # 4) Market share
        market_share = calculate_market_share(brand_totals)

        # 5) City/Class pivot with x330 approach
        df_city_pivot = None
        if self.checkbox_city.isChecked():
            df_city_pivot = city_pivot_approx(df, brand_cols, workload_cols)
        df_class_pivot = None
        if self.checkbox_class.isChecked():
            df_class_pivot = class_pivot_approx(df, brand_cols, workload_cols)

        # Convert results to DataFrame
        df_totals = pd.DataFrame(list(brand_totals.items()), columns=["Brand", "Total Yearly Samples (Pro-Rated)"])
        df_totals = add_date_time_columns(df_totals)

        df_share = pd.DataFrame(list(market_share.items()), columns=["Brand", "Market Share (%)"])
        df_share = add_date_time_columns(df_share)

        # 6) Save
        self.save_results(analyzer, df_totals, df_share, df_city_pivot, df_class_pivot)

        QMessageBox.information(self, "Done", f"{analyzer} analysis completed.")
        logger.info(f"{analyzer} analysis completed and saved.")

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
    logging.info("Launching Multi-Analyzer Market Share App with modern UI.")
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
