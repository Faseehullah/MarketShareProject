# modern_dashboard.py

import os
import sys
import pandas as pd
import plotly.express as px
import plotly.io as pio
from fpdf import FPDF

from PyQt5.QtCore import QUrl
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QWidget, QLabel,
    QPushButton, QComboBox, QScrollArea, QMessageBox, QFileDialog, QTabWidget
)
from PyQt5.QtWebEngineWidgets import QWebEngineView


class ModernDataAnalysisApp(QMainWindow):
    def __init__(self):
        super().__init__()

        # Set up main window
        self.setWindowTitle("Modern Data Analysis Dashboard")
        self.setGeometry(100, 100, 1600, 1000)

        # Data variables
        self.data = None
        self.filtered_data = None

        # By default, you might want to match the path your main.py might use.
        # Or leave it blank, so user must "Load Dataset."
        self.file_path = ""

        # Output directory for charts, PDF, etc.
        self.output_dir = os.path.join(os.getcwd(), "Output_Charts_and_Data")
        os.makedirs(self.output_dir, exist_ok=True)

        # Main scroll area
        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_widget = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_widget)
        self.scroll_layout.setSpacing(20)
        self.scroll_layout.setContentsMargins(20, 20, 30, 30)
        self.scroll_area.setWidget(self.scroll_widget)
        self.setCentralWidget(self.scroll_area)

        # Header label
        self.header = QLabel("<h1 style='color: darkblue; text-align: center;'>Modern Data Analysis Dashboard</h1>")
        self.scroll_layout.addWidget(self.header)

        # Export to PDF button
        self.export_button = QPushButton("Export Report to PDF")
        self.export_button.setStyleSheet("font-size: 16px; padding: 10px; background-color: lightgreen; border-radius: 5px;")
        self.export_button.clicked.connect(self.export_to_pdf)
        self.scroll_layout.addWidget(self.export_button)

        # Load dataset button
        self.load_button = QPushButton("Load Dataset")
        self.load_button.setStyleSheet("font-size: 16px; padding: 10px; background-color: lightblue; border-radius: 5px;")
        self.load_button.clicked.connect(self.load_dataset)
        self.scroll_layout.addWidget(self.load_button)

        # Filters layout
        self.filters_layout = QHBoxLayout()
        self.filters_layout.setSpacing(10)

        # Region filter
        self.region_filter = QComboBox()
        self.region_filter.addItem("All Regions")
        self.region_filter.currentIndexChanged.connect(self.apply_filters)
        self.filters_layout.addWidget(QLabel("Region:"))
        self.filters_layout.addWidget(self.region_filter)

        # Type filter
        self.type_filter = QComboBox()
        self.type_filter.addItem("All Types")
        self.type_filter.currentIndexChanged.connect(self.apply_filters)
        self.filters_layout.addWidget(QLabel("Type:"))
        self.filters_layout.addWidget(self.type_filter)

        self.scroll_layout.addLayout(self.filters_layout)

        # Tab widget for charts
        self.tabs = QTabWidget()
        self.scroll_layout.addWidget(self.tabs)

        # Dictionary of charts
        self.charts = {
            "Total Samples by Region (#1)": QWebEngineView(),
            "Test Type Contribution (#2)": QWebEngineView(),
            "Distribution of Samples (#3)": QWebEngineView(),
            "Region-wise Test Contribution (#4)": QWebEngineView(),
            "Type-wise Test Contribution (#5)": QWebEngineView(),
            "Top 10 Chemistry Contributors (#6)": QWebEngineView(),
            "Bottom 10 Chemistry Contributors (#7)": QWebEngineView(),
            "Top 10 CBC Contributors (#8)": QWebEngineView(),
            "Bottom 10 CBC Contributors (#9)": QWebEngineView(),
            "Top 10 Immunoassay Contributors (#10)": QWebEngineView(),
            "Bottom 10 Immunoassay Contributors (#11)": QWebEngineView(),
            "Number of Customers by Region (#12)": QWebEngineView(),
            "Type-wise Test Load (#13)": QWebEngineView(),
            "Class-wise Test Distribution (#14)": QWebEngineView(),  # NEW
            "Heatmap Distribution (#15) (Based on Type)": QWebEngineView(),
            "City-wise Test Distribution (#16)": QWebEngineView(),   # NEW
        }

        # Add each chart as a tab
        for title, widget in self.charts.items():
            chart_box = self.create_chart_box(title, widget)
            self.tabs.addTab(chart_box, title)

    def export_chart_to_png(self, fig, filename):
        png_path = os.path.join(self.output_dir, filename)
        try:
            pio.write_image(fig, png_path)
            print(f"Chart saved as PNG: {png_path}")
        except Exception as e:
            print(f"Error exporting chart to PNG: {e}")

    def create_chart_box(self, title, chart_widget):
        box = QVBoxLayout()
        box.setSpacing(10)
        box.setContentsMargins(10, 10, 10, 10)
        container = QWidget()
        container.setLayout(box)
        container.setStyleSheet("""
            border: 1px solid #ccc;
            border-radius: 8px;
            background-color: #f9f9f9;
            padding: 10px;
        """)
        label = QLabel(f"<h3 style='color: #333;'>{title}</h3>")
        box.addWidget(label)

        chart_widget.setMinimumHeight(600)
        chart_widget.setMaximumHeight(600)
        box.addWidget(chart_widget)

        return container

    def generate_chart(self, fig, html_filename, png_filename, chart_widget):
        html_path = os.path.join(self.output_dir, html_filename)
        try:
            fig.write_html(html_path)
            self.export_chart_to_png(fig, png_filename)
            chart_widget.setUrl(QUrl.fromLocalFile(html_path))
        except Exception as e:
            print(f"Error generating {html_filename}: {e}")
            QMessageBox.critical(self, "Error", f"Error generating {html_filename}: {e}")

    def load_dataset(self):
        """
        By default, tries to open an Excel file with a 'Consolidated' sheet
        that has columns like 'CHEMISTRY_TOTAL', 'IMMUNOASSAY_TOTAL', 'HEMATOLOGY_TOTAL', 'TOTAL_SAMPLES',
        plus 'Region', 'Type', 'CITY', 'Class', 'Customer Name'
        (like your main.py).
        """
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx *.xls)")
        if not file_path:
            return
        self.file_path = file_path
        print(f"Loading dataset from: {self.file_path}")

        try:
            # Reading the "Consolidated" sheet by default
            # Modify if your aggregator writes to a different name
            self.data = pd.read_excel(self.file_path, sheet_name="Consolidated")

            # numeric columns
            numeric_columns = ["CHEMISTRY_TOTAL", "IMMUNOASSAY_TOTAL", "HEMATOLOGY_TOTAL", "TOTAL_SAMPLES"]
            for col in numeric_columns:
                if col in self.data.columns:
                    if self.data[col].dtype == object:
                        self.data[col] = self.data[col].str.replace(',', '').astype(float)
                    else:
                        self.data[col] = pd.to_numeric(self.data[col], errors="coerce")
            # drop rows if any numeric columns are missing
            self.data.dropna(subset=numeric_columns, inplace=True)

            self.filtered_data = self.data.copy()

            # Populate filters
            if "Region" in self.data.columns:
                self.region_filter.clear()
                self.region_filter.addItem("All Regions")
                unique_regions = sorted([r for r in self.data["Region"].dropna().unique() if r])
                for r in unique_regions:
                    self.region_filter.addItem(r)

            if "Type" in self.data.columns:
                self.type_filter.clear()
                self.type_filter.addItem("All Types")
                unique_types = sorted([t for t in self.data["Type"].dropna().unique() if t])
                for t in unique_types:
                    self.type_filter.addItem(t)

            # Show initial charts
            self.show_charts()
            QMessageBox.information(self, "Success", "Dataset loaded successfully!")
        except Exception as e:
            print(f"Error loading dataset: {e}")
            QMessageBox.critical(self, "Error", f"Error loading dataset: {e}")

    def apply_filters(self):
        """
        Applies region/type filters if available, then re-calls show_charts().
        """
        try:
            if self.data is None:
                return

            region = self.region_filter.currentText()
            data_type = self.type_filter.currentText()

            self.filtered_data = self.data.copy()
            if "Region" in self.data.columns and region != "All Regions":
                self.filtered_data = self.filtered_data[self.filtered_data["Region"] == region]
            if "Type" in self.data.columns and data_type != "All Types":
                self.filtered_data = self.filtered_data[self.filtered_data["Type"] == data_type]

            self.show_charts()
        except Exception as e:
            print(f"Error applying filters: {e}")
            QMessageBox.critical(self, "Error", f"Error applying filters: {e}")

    def show_charts(self):
        """
        Refresh each visualization based on the filtered dataset.
        """
        self.show_bar_chart()                          # (#1)
        self.show_pie_chart()                          # (#2)
        self.show_histogram()                          # (#3)
        self.show_region_wise_test_contribution()      # (#4)
        self.show_type_wise_test_contribution()        # (#5)
        self.show_top_chemistry_chart()                # (#6)
        self.show_bottom_chemistry_chart()             # (#7)
        self.show_top_cbc_chart()                      # (#8)
        self.show_bottom_cbc_chart()                   # (#9)
        self.show_top_immunoassay_chart()              # (#10)
        self.show_bottom_immunoassay_chart()           # (#11)
        self.show_region_wise_customer_chart()         # (#12)
        self.show_type_wise_tests_chart()              # (#13)
        self.show_class_wise_distribution()            # (#14)
        self.show_heatmap_chart()                      # (#15)
        self.show_city_wise_distribution()             # (#16)

    # The rest of the code (chart methods) remains the same as you posted,
    # referencing self.filtered_data for plotting.


    def show_bar_chart(self):
        """Displays a stacked bar chart for TOTAL_SAMPLES by Region (#1)."""
        # exactly the same code as you posted
        pass

    # ... etc. (All the same chart methods as in your snippet) ...

    def export_to_pdf(self):
        """
        Exports all dashboard charts to a single PDF report,
        saving it to the `Output_Charts_and_Data` directory.
        """
        try:
            pdf = FPDF()
            pdf.set_auto_page_break(auto=True, margin=15)

            # Title page of the PDF
            pdf.add_page()
            pdf.set_font("helvetica", "B", 16)
            pdf.cell(200, 10, "Modern Data Analysis Report", ln=1, align="C")
            pdf.ln(10)

            chart_files = [
                ("bar_chart.png", "Total Samples by Region (#1)"),
                ("pie_chart.png", "Test Type Contribution (#2)"),
                ("histogram_chart.png", "Distribution of Total Samples (#3)"),
                ("region_wise_test_pie.png", "Region-wise Test Contribution (#4)"),
                ("type_wise_test_pie.png", "Type-wise Test Contribution (#5)"),
                ("top_chemistry_chart.png", "Top 10 Chemistry Contributors (#6)"),
                ("bottom_chemistry_chart.png", "Bottom 10 Chemistry Contributors (#7)"),
                ("top_cbc_chart.png", "Top 10 CBC Contributors (#8)"),
                ("bottom_cbc_chart.png", "Bottom 10 CBC Contributors (#9)"),
                ("top_immunoassay_chart.png", "Top 10 Immunoassay Contributors (#10)"),
                ("bottom_immunoassay_chart.png", "Bottom 10 Immunoassay Contributors (#11)"),
                ("region_customer_chart.png", "Number of Customers by Region (#12)"),
                ("type_tests_chart.png", "Type-wise Test Load (#13)"),
                ("class_wise_distribution.png", "Class-wise Test Distribution (#14)"),
                ("heatmap_chart.png", "Heatmap Distribution (#15)"),
                ("city_wise_distribution.png", "City-wise Test Distribution (#16)"),
            ]

            for png_file, title in chart_files:
                png_path = os.path.join(self.output_dir, png_file)
                if os.path.exists(png_path):
                    pdf.add_page()
                    pdf.set_font("helvetica", "B", 12)
                    pdf.cell(200, 10, title, ln=1, align="C")
                    pdf.ln(5)
                    pdf.image(png_path, x=10, y=30, w=190)
                else:
                    print(f"Warning: PNG file for '{title}' not found at {png_path}.")
                    QMessageBox.warning(self, "Missing Chart", f"PNG file for '{title}' not found at:\n{png_path}")

            pdf_path = os.path.join(self.output_dir, "Data_Analysis_Report.pdf")
            pdf.output(pdf_path)
            print(f"PDF Report saved at: {pdf_path}")
            QMessageBox.information(self, "Success", f"PDF Report saved at:\n{pdf_path}")

        except Exception as e:
            print(f"Error exporting to PDF: {e}")
            QMessageBox.critical(self, "Error", f"Error exporting to PDF: {e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ModernDataAnalysisApp()
    window.show()
    sys.exit(app.exec_())
