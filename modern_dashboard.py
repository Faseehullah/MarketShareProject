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

        # Set up main window properties
        self.setWindowTitle("Modern Data Analysis Dashboard")
        self.setGeometry(100, 100, 1600, 1000)

        # Data-related variables
        self.data = None
        self.filtered_data = None

        # By default, an example path or empty
        self.file_path = ""

        # Create an output directory to store charts and PDF
        self.output_dir = os.path.join(os.getcwd(), "Output_Charts_and_Data")
        os.makedirs(self.output_dir, exist_ok=True)

        # Main scroll area (vertical scrolling)
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

        # Button: Export all charts to a single PDF
        self.export_button = QPushButton("Export Report to PDF")
        self.export_button.setStyleSheet("""
            font-size: 16px;
            padding: 10px;
            background-color: lightgreen;
            border-radius: 5px;
        """)
        self.export_button.clicked.connect(self.export_to_pdf)
        self.scroll_layout.addWidget(self.export_button)

        # Button: Load Dataset
        self.load_button = QPushButton("Load Dataset")
        self.load_button.setStyleSheet("""
            font-size: 16px;
            padding: 10px;
            background-color: lightblue;
            border-radius: 5px;
        """)
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

        # Add filters layout
        self.scroll_layout.addLayout(self.filters_layout)

        # Tabs for charts
        self.tabs = QTabWidget()
        self.scroll_layout.addWidget(self.tabs)

        # Dictionary of chart titles -> QWebEngineView
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
            "Class-wise Test Distribution (#14)": QWebEngineView(),
            "Heatmap Distribution (#15) (Based on Type)": QWebEngineView(),
            "City-wise Test Distribution (#16)": QWebEngineView(),
        }

        # Create a tab for each chart
        for title, widget in self.charts.items():
            chart_box = self.create_chart_box(title, widget)
            self.tabs.addTab(chart_box, title)

    def export_chart_to_png(self, fig, filename):
        """Exports a Plotly figure to PNG in the output directory."""
        try:
            png_path = os.path.join(self.output_dir, filename)
            pio.write_image(fig, png_path)
            print(f"Chart saved as PNG: {png_path}")
        except Exception as e:
            print(f"Error exporting chart to PNG: {e}")

    def create_chart_box(self, title, chart_widget):
        """Create a styled container with a label + chart widget."""
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
        """Generate HTML + PNG for a Plotly figure, load into QWebEngineView."""
        try:
            html_path = os.path.join(self.output_dir, html_filename)
            fig.write_html(html_path)
            # Export PNG
            self.export_chart_to_png(fig, png_filename)
            # Load the HTML into the chart widget
            chart_widget.setUrl(QUrl.fromLocalFile(html_path))
        except Exception as e:
            print(f"Error generating {html_filename}: {e}")
            QMessageBox.critical(self, "Error", f"Error generating {html_filename}: {e}")

    def load_dataset(self):
        """Prompt user to pick an Excel file, read it, and aggregate columns for CHEM/IA/CBC."""
        try:
            file_dialog = QFileDialog()
            file_path, _ = file_dialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx *.xls)")
            if not file_path:
                return

            self.file_path = file_path
            print(f"Loading dataset from: {self.file_path}")

            # Read the data from the single sheet (or "Consolidated"? Up to you)
            # If your file actually has multiple possible sheets, add sheet_name param or logic.
            self.data = pd.read_excel(self.file_path, sheet_name=0)

            # We'll create 3 new columns: CHEM_TOTAL, IA_TOTAL, CBC_TOTAL
            # Summing the workload columns
            self.data["CHEM_TOTAL"] = (
                self.data.get("CHEM Workload - Brand 1", 0).fillna(0) +
                self.data.get("CHEM Workload - Brand 2", 0).fillna(0) +
                self.data.get("CHEM Workload - Brand 3", 0).fillna(0) +
                self.data.get("CHEM Workload - Brand 4", 0).fillna(0)
            )

            self.data["IA_TOTAL"] = (
                self.data.get("IA Workload - Brand 1", 0).fillna(0) +
                self.data.get("IA Workload - Brand 2", 0).fillna(0) +
                self.data.get("IA Workload - Brand 3", 0).fillna(0)
            )

            self.data["CBC_TOTAL"] = (
                self.data.get("CBC Workload - Brand 1", 0).fillna(0) +
                self.data.get("CBC Workload - Brand 2", 0).fillna(0) +
                self.data.get("CBC Workload - Brand 3", 0).fillna(0) +
                self.data.get("CBC Workload - Brand 4", 0).fillna(0)
            )

            # Additionally, we define a TOTAL_SAMPLES col for convenience
            self.data["TOTAL_SAMPLES"] = (
                self.data["CHEM_TOTAL"] + self.data["IA_TOTAL"] + self.data["CBC_TOTAL"]
            )

            # We'll drop rows that have 0 for TOT_SAMPLES if needed
            # but let's keep them for now
            # self.data = self.data[self.data["TOTAL_SAMPLES"] > 0]

            # filtered_data is initially the same
            self.filtered_data = self.data.copy()

            # Set up filters
            self.region_filter.clear()
            self.type_filter.clear()

            # Region filter
            self.region_filter.addItem("All Regions")
            if "Region" in self.data.columns:
                regions = sorted(self.data["Region"].dropna().unique())
                for r in regions:
                    self.region_filter.addItem(r)

            # Type filter
            self.type_filter.addItem("All Types")
            if "Type" in self.data.columns:
                types = sorted(self.data["Type"].dropna().unique())
                for t in types:
                    self.type_filter.addItem(t)

            # Show charts
            self.show_charts()
            QMessageBox.information(self, "Success", "Dataset loaded and aggregated successfully!")

        except Exception as e:
            print(f"Error loading dataset: {e}")
            QMessageBox.critical(self, "Error", f"Error loading dataset: {e}")

    def apply_filters(self):
        """Apply region & type filters to self.filtered_data, then re-show charts."""
        try:
            if self.data is None:
                return

            region = self.region_filter.currentText()
            data_type = self.type_filter.currentText()

            self.filtered_data = self.data.copy()

            if region != "All Regions" and "Region" in self.data.columns:
                self.filtered_data = self.filtered_data[self.filtered_data["Region"] == region]

            if data_type != "All Types" and "Type" in self.data.columns:
                self.filtered_data = self.filtered_data[self.filtered_data["Type"] == data_type]

            self.show_charts()
        except Exception as e:
            print(f"Error applying filters: {e}")
            QMessageBox.critical(self, "Error", f"Error applying filters: {e}")

    def show_charts(self):
        """
        Calls all chart methods referencing 'CHEM_TOTAL', 'IA_TOTAL',
        'CBC_TOTAL', and 'TOTAL_SAMPLES' from self.filtered_data.
        """
        self.show_bar_chart()
        self.show_pie_chart()
        self.show_histogram()
        self.show_region_wise_test_contribution()
        self.show_type_wise_test_contribution()
        self.show_top_chemistry_chart()
        self.show_bottom_chemistry_chart()
        self.show_top_cbc_chart()
        self.show_bottom_cbc_chart()
        self.show_top_immunoassay_chart()
        self.show_bottom_immunoassay_chart()
        self.show_region_wise_customer_chart()
        self.show_type_wise_tests_chart()
        self.show_class_wise_distribution()
        self.show_heatmap_chart()
        self.show_city_wise_distribution()

    # CHART #1: Stacked bar of total samples by Region
    def show_bar_chart(self):
        chart_title = "Total Samples by Region (#1)"
        if self.filtered_data.empty:
            self.charts[chart_title].setHtml("<h3>No data to display.</h3>")
            return

        # Sum CHEM_TOTAL, IA_TOTAL, CBC_TOTAL by region
        region_data = (
            self.filtered_data.groupby("Region")[["CHEM_TOTAL", "IA_TOTAL", "CBC_TOTAL"]]
            .sum()
            .reset_index()
        )

        fig = px.bar(
            region_data,
            x="Region",
            y=["CHEM_TOTAL", "IA_TOTAL", "CBC_TOTAL"],
            title=chart_title,
            barmode="stack",
            color_discrete_sequence=px.colors.qualitative.Dark24
        )
        fig.update_layout(
            xaxis_title="Region",
            yaxis_title="Total Samples",
            legend_title="Test Type",
            template="plotly_white",
            font=dict(size=12),
        )

        self.generate_chart(fig, "bar_chart.html", "bar_chart.png", self.charts[chart_title])

    # CHART #2: Pie chart of test type contributions (CHEM_TOTAL, IA_TOTAL, CBC_TOTAL).
    def show_pie_chart(self):
        chart_title = "Test Type Contribution (#2)"
        if self.filtered_data.empty:
            self.charts[chart_title].setHtml("<h3>No data to display.</h3>")
            return

        # sum of columns
        test_totals = self.filtered_data[["CHEM_TOTAL", "IA_TOTAL", "CBC_TOTAL"]].sum()
        fig = px.pie(
            values=test_totals,
            names=test_totals.index,
            title=chart_title,
            color_discrete_sequence=px.colors.qualitative.Set2,
        )
        fig.update_traces(
            hovertemplate="<b>%{label}</b><br>Total: %{value:,.0f}<extra></extra>",
            textinfo="percent+label",
        )
        fig.update_layout(template="plotly_white", legend_title="Test Type", font=dict(size=12),)

        self.generate_chart(fig, "pie_chart.html", "pie_chart.png", self.charts[chart_title])

    # CHART #3: Distribution histogram of TOTAL_SAMPLES with mean/median lines
    def show_histogram(self):
        chart_title = "Distribution of Samples (#3)"
        if self.filtered_data.empty:
            self.charts[chart_title].setHtml("<h3>No data to display.</h3>")
            return

        fig = px.histogram(
            self.filtered_data,
            x="TOTAL_SAMPLES",
            title=chart_title,
            nbins=30,
            color_discrete_sequence=["#636EFA"],
        )
        mean_value = self.filtered_data["TOTAL_SAMPLES"].mean()
        median_value = self.filtered_data["TOTAL_SAMPLES"].median()

        fig.add_vline(
            x=mean_value, line_width=2, line_dash="dash", line_color="red",
            annotation_text="Mean", annotation_position="top left"
        )
        fig.add_vline(
            x=median_value, line_width=2, line_dash="dash", line_color="green",
            annotation_text="Median", annotation_position="top right"
        )
        fig.update_layout(
            xaxis_title="Total Samples",
            yaxis_title="Frequency",
            template="plotly_white",
            font=dict(size=12),
        )

        self.generate_chart(fig, "histogram_chart.html", "histogram_chart.png", self.charts[chart_title])

    # CHART #4: Region-wise Test Contribution (pie of total region sums)
    def show_region_wise_test_contribution(self):
        chart_title = "Region-wise Test Contribution (#4)"
        if self.filtered_data.empty:
            self.charts[chart_title].setHtml("<h3>No data to display.</h3>")
            return

        # sum up CHEM_TOTAL + IA_TOTAL + CBC_TOTAL for each region
        region_totals = (
            self.filtered_data.groupby("Region")[["CHEM_TOTAL","IA_TOTAL","CBC_TOTAL"]]
            .sum()
            .sum(axis=1)  # sum across columns
            .reset_index()
        )
        region_totals.columns = ["Region", "Total"]

        fig = px.pie(
            region_totals,
            values="Total",
            names="Region",
            title=chart_title,
            color="Region",
            color_discrete_sequence=px.colors.qualitative.Dark24,
        )
        fig.update_traces(
            hovertemplate="<b>%{label}</b><br>Total: %{value:,.0f}<extra></extra>",
            textinfo="percent+label",
        )
        fig.update_layout(template="plotly_white", legend_title="Region", font=dict(size=12),)

        self.generate_chart(fig, "region_wise_test_pie.html", "region_wise_test_pie.png", self.charts[chart_title])

    # CHART #5: Type-wise test contribution (pie of sum of each type's total?)
    def show_type_wise_test_contribution(self):
        chart_title = "Type-wise Test Contribution (#5)"
        if self.filtered_data.empty:
            self.charts[chart_title].setHtml("<h3>No data to display.</h3>")
            return

        # Sum of CHEM_TOTAL + IA_TOTAL + CBC_TOTAL by Type
        type_totals = (
            self.filtered_data.groupby("Type")[["CHEM_TOTAL","IA_TOTAL","CBC_TOTAL"]]
            .sum()
            .sum(axis=1)
            .reset_index()
        )
        type_totals.columns = ["Type", "Total"]

        fig = px.pie(
            type_totals,
            values="Total",
            names="Type",
            title=chart_title,
            color="Type",
            color_discrete_sequence=px.colors.qualitative.Set3,
        )
        fig.update_traces(
            hovertemplate="<b>%{label}</b><br>Total: %{value:,.0f}<extra></extra>",
            textinfo="percent+label",
        )
        fig.update_layout(template="plotly_white", legend_title="Type", font=dict(size=12),)

        self.generate_chart(fig, "type_wise_test_pie.html", "type_wise_test_pie.png", self.charts[chart_title])

    # CHART #6: Top 10 Chemistry
    def show_top_chemistry_chart(self):
        chart_title = "Top 10 Chemistry Contributors (#6)"
        if self.filtered_data.empty:
            self.charts[chart_title].setHtml("<h3>No data.</h3>")
            return

        valid_data = self.filtered_data[self.filtered_data["CHEM_TOTAL"]>0]
        if valid_data.empty:
            self.charts[chart_title].setHtml("<h3>No valid data (all zeros).</h3>")
            return

        top_chem = valid_data.nlargest(10, "CHEM_TOTAL")
        fig = px.bar(
            top_chem,
            x="Customer Name",
            y="CHEM_TOTAL",
            title=chart_title,
            color="Region",
            color_discrete_sequence=px.colors.qualitative.Pastel,
            text="CHEM_TOTAL",
        )
        fig.update_traces(texttemplate='%{text:,}', textposition='outside')
        fig.update_layout(
            xaxis_title="Customer Name",
            yaxis_title="Chemistry Total",
            template="plotly_white",
            font=dict(size=12),
            xaxis_tickangle=-45
        )
        self.generate_chart(fig, "top_chemistry_chart.html", "top_chemistry_chart.png", self.charts[chart_title])

    # CHART #7: Bottom 10 Chemistry
    def show_bottom_chemistry_chart(self):
        chart_title = "Bottom 10 Chemistry Contributors (#7)"
        if self.filtered_data.empty:
            self.charts[chart_title].setHtml("<h3>No data.</h3>")
            return

        valid_data = self.filtered_data[self.filtered_data["CHEM_TOTAL"]>0]
        if valid_data.empty:
            self.charts[chart_title].setHtml("<h3>No valid data (all zeros).</h3>")
            return

        bottom_chem = valid_data.nsmallest(10, "CHEM_TOTAL")
        fig = px.bar(
            bottom_chem,
            x="Customer Name",
            y="CHEM_TOTAL",
            title=chart_title,
            color="Region",
            color_discrete_sequence=px.colors.qualitative.Set3,
            text="CHEM_TOTAL",
        )
        fig.update_traces(texttemplate='%{text:,}', textposition='outside')
        fig.update_layout(
            xaxis_title="Customer Name",
            yaxis_title="Chemistry Total",
            template="plotly_white",
            font=dict(size=12),
            xaxis_tickangle=-45
        )
        self.generate_chart(fig, "bottom_chemistry_chart.html", "bottom_chemistry_chart.png", self.charts[chart_title])

    # CHART #8: Top 10 CBC
    def show_top_cbc_chart(self):
        chart_title = "Top 10 CBC Contributors (#8)"
        if self.filtered_data.empty:
            self.charts[chart_title].setHtml("<h3>No data.</h3>")
            return

        valid_data = self.filtered_data[self.filtered_data["CBC_TOTAL"]>0]
        if valid_data.empty:
            self.charts[chart_title].setHtml("<h3>No valid data (zeros).</h3>")
            return

        top_cbc = valid_data.nlargest(10, "CBC_TOTAL")
        fig = px.bar(
            top_cbc,
            x="Customer Name",
            y="CBC_TOTAL",
            title=chart_title,
            color="Region",
            color_discrete_sequence=px.colors.qualitative.Set2,
            text="CBC_TOTAL",
        )
        fig.update_traces(texttemplate='%{text:,}', textposition='outside')
        fig.update_layout(
            xaxis_title="Customer Name",
            yaxis_title="CBC Total",
            template="plotly_white",
            font=dict(size=12),
            xaxis_tickangle=-45
        )
        self.generate_chart(fig, "top_cbc_chart.html", "top_cbc_chart.png", self.charts[chart_title])

    # CHART #9: Bottom 10 CBC
    def show_bottom_cbc_chart(self):
        chart_title = "Bottom 10 CBC Contributors (#9)"
        if self.filtered_data.empty:
            self.charts[chart_title].setHtml("<h3>No data.</h3>")
            return

        valid_data = self.filtered_data[self.filtered_data["CBC_TOTAL"]>0]
        if valid_data.empty:
            self.charts[chart_title].setHtml("<h3>No valid data (zeros).</h3>")
            return

        bottom_cbc = valid_data.nsmallest(10, "CBC_TOTAL")
        fig = px.bar(
            bottom_cbc,
            x="Customer Name",
            y="CBC_TOTAL",
            title=chart_title,
            color="Region",
            color_discrete_sequence=px.colors.qualitative.Set3,
            text="CBC_TOTAL",
        )
        fig.update_traces(texttemplate='%{text:,}', textposition='outside')
        fig.update_layout(
            xaxis_title="Customer Name",
            yaxis_title="CBC Total",
            template="plotly_white",
            font=dict(size=12),
            xaxis_tickangle=-45
        )
        self.generate_chart(fig, "bottom_cbc_chart.html", "bottom_cbc_chart.png", self.charts[chart_title])

    # CHART #10: Top 10 Immunoassay
    def show_top_immunoassay_chart(self):
        chart_title = "Top 10 Immunoassay Contributors (#10)"
        if self.filtered_data.empty:
            self.charts[chart_title].setHtml("<h3>No data.</h3>")
            return

        valid_data = self.filtered_data[self.filtered_data["IA_TOTAL"]>0]
        if valid_data.empty:
            self.charts[chart_title].setHtml("<h3>No valid data (zeros).</h3>")
            return

        top_ia = valid_data.nlargest(10, "IA_TOTAL")
        fig = px.bar(
            top_ia,
            x="Customer Name",
            y="IA_TOTAL",
            title=chart_title,
            color="Region",
            color_discrete_sequence=px.colors.qualitative.Pastel,
            text="IA_TOTAL",
        )
        fig.update_traces(texttemplate='%{text:,}', textposition='outside')
        fig.update_layout(
            xaxis_title="Customer Name",
            yaxis_title="Immunoassay Total",
            template="plotly_white",
            font=dict(size=12),
            xaxis_tickangle=-45
        )
        self.generate_chart(fig, "top_immunoassay_chart.html", "top_immunoassay_chart.png", self.charts[chart_title])

    # CHART #11: Bottom 10 Immunoassay
    def show_bottom_immunoassay_chart(self):
        chart_title = "Bottom 10 Immunoassay Contributors (#11)"
        if self.filtered_data.empty:
            self.charts[chart_title].setHtml("<h3>No data.</h3>")
            return

        valid_data = self.filtered_data[self.filtered_data["IA_TOTAL"]>0]
        if valid_data.empty:
            self.charts[chart_title].setHtml("<h3>No valid data (zeros).</h3>")
            return

        bottom_ia = valid_data.nsmallest(10, "IA_TOTAL")
        fig = px.bar(
            bottom_ia,
            x="Customer Name",
            y="IA_TOTAL",
            title=chart_title,
            color="Region",
            color_discrete_sequence=px.colors.qualitative.Set1,
            text="IA_TOTAL",
        )
        fig.update_traces(texttemplate='%{text:,}', textposition='outside')
        fig.update_layout(
            xaxis_title="Customer Name",
            yaxis_title="Immunoassay Total",
            template="plotly_white",
            font=dict(size=12),
            xaxis_tickangle=-45
        )
        self.generate_chart(fig, "bottom_immunoassay_chart.html", "bottom_immunoassay_chart.png", self.charts[chart_title])

    # CHART #12: Number of Customers by Region
    def show_region_wise_customer_chart(self):
        chart_title = "Number of Customers by Region (#12)"
        if self.filtered_data.empty:
            self.charts[chart_title].setHtml("<h3>No data.</h3>")
            return

        if "Region" not in self.filtered_data.columns or "Customer Name" not in self.filtered_data.columns:
            self.charts[chart_title].setHtml("<h3>Missing 'Region' or 'Customer Name' columns.</h3>")
            return

        region_customer_data = (
            self.filtered_data.groupby("Region")["Customer Name"]
            .nunique()
            .reset_index()
            .rename(columns={"Customer Name": "Unique Customers"})
            .sort_values(by="Unique Customers", ascending=False)
        )
        fig = px.bar(
            region_customer_data,
            x="Region",
            y="Unique Customers",
            title=chart_title,
            text="Unique Customers",
            color="Region",
            color_discrete_sequence=px.colors.qualitative.Pastel,
        )
        fig.update_traces(texttemplate='%{text}', textposition='outside')
        fig.update_layout(
            xaxis_title="Region",
            yaxis_title="Unique Customers",
            template="plotly_white",
            font=dict(size=12),
            xaxis_tickangle=-45
        )

        self.generate_chart(fig, "region_customer_chart.html", "region_customer_chart.png", self.charts[chart_title])

    # CHART #13: Type-wise test load (grouped bar for CHEM_TOTAL, IA_TOTAL, CBC_TOTAL by Type)
    def show_type_wise_tests_chart(self):
        chart_title = "Type-wise Test Load (#13)"
        if self.filtered_data.empty:
            self.charts[chart_title].setHtml("<h3>No data.</h3>")
            return

        if "Type" not in self.filtered_data.columns:
            self.charts[chart_title].setHtml("<h3>Missing 'Type' column.</h3>")
            return

        type_tests = (
            self.filtered_data.groupby("Type")[["CHEM_TOTAL","IA_TOTAL","CBC_TOTAL"]]
            .sum()
            .reset_index()
        )
        fig = px.bar(
            type_tests,
            x="Type",
            y=["CHEM_TOTAL","IA_TOTAL","CBC_TOTAL"],
            title=chart_title,
            barmode="group",
            color_discrete_sequence=px.colors.qualitative.Set3,
            text_auto=True,
        )
        fig.update_layout(
            xaxis_title="Type",
            yaxis_title="Total Tests",
            legend_title="Test Type",
            template="plotly_white",
            font=dict(size=12),
        )
        self.generate_chart(fig, "type_tests_chart.html", "type_tests_chart.png", self.charts[chart_title])

    # CHART #14: Class-wise distribution (grouped bar of CHEM_TOTAL, IA_TOTAL, CBC_TOTAL by Class)
    def show_class_wise_distribution(self):
        chart_title = "Class-wise Test Distribution (#14)"
        if self.filtered_data.empty:
            self.charts[chart_title].setHtml("<h3>No data.</h3>")
            return

        if "Class" not in self.filtered_data.columns:
            self.charts[chart_title].setHtml("<h3>Missing 'Class' column.</h3>")
            return

        class_data = (
            self.filtered_data.groupby("Class")[["CHEM_TOTAL","IA_TOTAL","CBC_TOTAL"]]
            .sum()
            .reset_index()
        )
        fig = px.bar(
            class_data,
            x="Class",
            y=["CHEM_TOTAL","IA_TOTAL","CBC_TOTAL"],
            barmode="group",
            title=chart_title,
            color_discrete_sequence=px.colors.qualitative.Set2,
            text_auto=True
        )
        fig.update_layout(
            xaxis_title="Class",
            yaxis_title="Total Tests",
            template="plotly_white",
            font=dict(size=12),
            legend_title="Test Type"
        )
        self.generate_chart(fig, "class_wise_distribution.html", "class_wise_distribution.png", self.charts[chart_title])

    # CHART #15: Heatmap of TOTAL_SAMPLES by region & type
    def show_heatmap_chart(self):
        chart_title = "Heatmap Distribution (#15) (Based on Type)"
        if self.filtered_data.empty:
            self.charts[chart_title].setHtml("<h3>No data.</h3>")
            return

        if "Region" not in self.filtered_data.columns or "Type" not in self.filtered_data.columns:
            self.charts[chart_title].setHtml("<h3>Missing 'Region' or 'Type' columns.</h3>")
            return

        region_pivot = self.filtered_data.pivot_table(
            index="Region", columns="Type", values="TOTAL_SAMPLES", aggfunc="sum"
        )
        fig = px.imshow(
            region_pivot,
            text_auto=True,
            title=chart_title,
            color_continuous_scale="Viridis",
        )
        fig.update_layout(
            xaxis_title="Type",
            yaxis_title="Region",
            template="plotly_white",
            font=dict(size=12),
        )
        self.generate_chart(fig, "heatmap_chart.html", "heatmap_chart.png", self.charts[chart_title])

    # CHART #16: City-wise distribution (stacked bar of CHEM_TOTAL, IA_TOTAL, CBC_TOTAL by CITY)
    def show_city_wise_distribution(self):
        chart_title = "City-wise Test Distribution (#16)"
        if self.filtered_data.empty:
            self.charts[chart_title].setHtml("<h3>No data.</h3>")
            return

        if "CITY" not in self.filtered_data.columns:
            self.charts[chart_title].setHtml("<h3>Missing 'CITY' column.</h3>")
            return

        city_data = (
            self.filtered_data.groupby("CITY")[["CHEM_TOTAL","IA_TOTAL","CBC_TOTAL"]]
            .sum()
            .reset_index()
        )
        fig = px.bar(
            city_data,
            x="CITY",
            y=["CHEM_TOTAL","IA_TOTAL","CBC_TOTAL"],
            barmode="stack",
            title=chart_title,
            color_discrete_sequence=px.colors.qualitative.Dark24,
            text_auto=True
        )
        fig.update_layout(
            xaxis_title="City",
            yaxis_title="Total Tests",
            template="plotly_white",
            font=dict(size=12),
            legend_title="Test Type",
            xaxis_tickangle=-45
        )
        self.generate_chart(fig, "city_wise_distribution.html", "city_wise_distribution.png", self.charts[chart_title])

    def export_to_pdf(self):
        """
        Exports all charts to a single PDF in self.output_dir
        """
        try:
            pdf = FPDF()
            pdf.set_auto_page_break(auto=True, margin=15)

            # Title page
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
                    QMessageBox.warning(self, "Missing Chart", f"PNG file for '{title}' not found:\n{png_path}")

            pdf_path = os.path.join(self.output_dir, "Data_Analysis_Report.pdf")
            pdf.output(pdf_path)
            print(f"PDF Report saved at: {pdf_path}")
            QMessageBox.information(self, "Success", f"PDF Report saved at:\n{pdf_path}")

        except Exception as e:
            print(f"Error exporting to PDF: {e}")
            QMessageBox.critical(self, "Error", f"Error exporting to PDF: {e}")


# If running standalone:
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ModernDataAnalysisApp()
    window.show()
    sys.exit(app.exec_())
