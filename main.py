# src/main.py

"""
Market Share Analysis Tool
-------------------------
A comprehensive GUI application for analyzing market share data from Excel files.
Features include multi-analyzer support, interactive visualizations, and detailed analytics.
"""

import sys
import os
import logging
from datetime import datetime
from typing import Optional, Dict, Any, List, Union
from pathlib import Path

# Data processing imports
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import seaborn as sns
from matplotlib.figure import Figure

# PyQt imports
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QComboBox, QMessageBox, QLineEdit,
    QFormLayout, QCheckBox, QSpinBox, QTabWidget, QGroupBox, QScrollArea,
    QSizePolicy, QProgressBar
)
from PyQt5.QtCore import Qt, QSize
from PyQt5.QtGui import QFont, QPalette, QColor, QIntValidator

# Local imports
from aggregator import WorkingAggregator, AnalysisResult
from config import MarketAnalysisConfig
from settings_dialog import SettingsDialog  # Importing the SettingsDialog

# Attempt to import xlsxwriter and handle ImportError gracefully
try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None  # Will handle this in save_results

# Configure logging
logger = logging.getLogger(__name__)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler('market_analysis.log')
    ]
)

class MarketAnalysisUI:
    """Base class containing shared UI styling and utilities."""

    STYLE_SHEET = """
        QMainWindow {
            background-color: #f5f6fa;
        }
        QTabWidget::pane {
            border: 1px solid #ddd;
            border-radius: 8px;
            background-color: white;
        }
        QPushButton {
            background-color: #1976d2;
            color: white;
            padding: 8px 16px;
            border: none;
            border-radius: 4px;
            font-size: 13px;
        }
        QPushButton:hover {
            background-color: #1565c0;
        }
        QGroupBox {
            background-color: white;
            border: 1px solid #ddd;
            border-radius: 8px;
            margin-top: 12px;
            padding: 15px;
        }
        QLabel {
            color: #2c3e50;
            font-size: 13px;
        }
        QLineEdit, QComboBox {
            padding: 6px;
            border: 1px solid #ddd;
            border-radius: 4px;
            min-height: 25px;
        }
    """

    @staticmethod
    def create_group_box(title: str, layout: Any) -> QGroupBox:
        """Create a styled group box with the given title and layout."""
        group = QGroupBox(title)
        group.setLayout(layout)
        return group

class FileHandler:
    """Manages file operations for the Market Share Analysis Tool."""

    @staticmethod
    def load_excel_sheets(file_path: str) -> List[str]:
        """Load sheet names from an Excel file."""
        try:
            return pd.ExcelFile(file_path).sheet_names
        except Exception as e:
            logger.error(f"Error reading Excel sheets: {str(e)}")
            raise ValueError(f"Failed to read Excel file: {str(e)}")

    @staticmethod
    def read_excel_data(file_path: str, sheet_name: str) -> pd.DataFrame:
        """Read and preprocess Excel data."""
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            return df.replace(["NILL", "Nill", "nill", "NIL", ""], np.nan)
        except Exception as e:
            logger.error(f"Error reading Excel data: {str(e)}")
            raise ValueError(f"Failed to read data: {str(e)}")

class ModernMarketAnalyzer(QMainWindow, MarketAnalysisUI):
    """Main application window for market share analysis."""

    def __init__(self):
        super().__init__()
        self.config = MarketAnalysisConfig()
        self.file_handler = FileHandler()
        self.latest_results: Optional[AnalysisResult] = None  # To store the latest analysis results
        self.init_ui()

    def init_ui(self):
        """Initialize the main user interface components."""
        self.setWindowTitle("Market Share Analysis Tool")
        self.setGeometry(100, 100, 1400, 900)
        self.setStyleSheet(self.STYLE_SHEET)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)

        # Create main tab widget
        tab_widget = QTabWidget()
        tab_widget.addTab(self.create_input_tab(), "Data Input")
        tab_widget.addTab(self.create_analysis_tab(), "Analysis")
        tab_widget.addTab(self.create_visualization_tab(), "Visualization")

        main_layout.addWidget(tab_widget)

        # Add progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)

        # Initialize status bar
        self.statusBar().showMessage("Ready")

    def create_input_tab(self) -> QWidget:
        """Create the data input interface."""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # File selection section
        file_layout = QFormLayout()

        # Input file selection
        input_group = QHBoxLayout()
        self.input_edit = QLineEdit()
        browse_input_btn = QPushButton("Browse")
        browse_input_btn.clicked.connect(self.browse_input)
        input_group.addWidget(self.input_edit)
        input_group.addWidget(browse_input_btn)
        file_layout.addRow("Input Excel:", input_group)

        # Sheet selection
        self.sheet_combo = QComboBox()
        file_layout.addRow("Sheet:", self.sheet_combo)

        # Output file selection
        output_group = QHBoxLayout()
        self.output_edit = QLineEdit()
        browse_output_btn = QPushButton("Browse")
        browse_output_btn.clicked.connect(self.browse_output)
        output_group.addWidget(self.output_edit)
        output_group.addWidget(browse_output_btn)
        file_layout.addRow("Output Excel:", output_group)

        layout.addWidget(self.create_group_box("File Selection", file_layout))

        # Analysis options section
        analysis_layout = QFormLayout()

        self.analyzer_combo = QComboBox()
        self.analyzer_combo.addItems(["IA", "CBC", "CHEM", "Consolidated"])
        analysis_layout.addRow("Analyzer:", self.analyzer_combo)

        self.region_combo = QComboBox()
        # Fetch regions from configuration if available
        regions = self.config.config_data.get("metadata", {}).get("regions", ["Region1", "Region2"])  # Default regions if not set
        self.region_combo.addItems(["All"] + regions)
        analysis_layout.addRow("Region:", self.region_combo)

        # Analysis checkboxes
        self.city_check = QCheckBox("Include City Analysis")
        self.class_check = QCheckBox("Include Class Analysis")
        self.value_check = QCheckBox("Include Value Analysis")

        analysis_layout.addRow(self.city_check)
        analysis_layout.addRow(self.class_check)
        analysis_layout.addRow(self.value_check)

        layout.addWidget(self.create_group_box("Analysis Options", analysis_layout))

        # Action buttons
        button_layout = QHBoxLayout()
        process_btn = QPushButton("Process Data")
        process_btn.clicked.connect(self.process_data)
        settings_btn = QPushButton("Settings")
        settings_btn.clicked.connect(self.open_settings)

        button_layout.addWidget(process_btn)
        button_layout.addWidget(settings_btn)
        layout.addLayout(button_layout)

        return tab

    def create_analysis_tab(self) -> QWidget:
        """Create the analysis results interface with real-time updates."""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Results overview section
        results_group = QGroupBox("Analysis Results")
        results_layout = QVBoxLayout()

        # Market share results
        self.volume_results = QLabel("Volume Market Share: Awaiting analysis")
        self.value_results = QLabel("Value Market Share: Awaiting analysis")
        self.city_results = QLabel("City Analysis: Not available")
        self.class_results = QLabel("Class Analysis: Not available")

        for label in [self.volume_results, self.value_results,
                     self.city_results, self.class_results]:
            label.setWordWrap(True)
            results_layout.addWidget(label)

        results_group.setLayout(results_layout)
        layout.addWidget(results_group)

        # Detailed metrics section
        metrics_group = QGroupBox("Detailed Metrics")
        metrics_layout = QVBoxLayout()

        self.metrics_display = QLabel("Detailed metrics will appear here after analysis")
        self.metrics_display.setWordWrap(True)
        metrics_layout.addWidget(self.metrics_display)

        metrics_group.setLayout(metrics_layout)
        layout.addWidget(metrics_group)

        return tab

    def create_visualization_tab(self) -> QWidget:
        """Create the data visualization interface with interactive charts."""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        viz_tabs = QTabWidget()

        # Market Share visualization
        self.market_share_widget = self.create_market_share_view()
        viz_tabs.addTab(self.market_share_widget, "Market Share")

        # Regional Analysis visualization
        self.regional_widget = self.create_regional_view()
        viz_tabs.addTab(self.regional_widget, "Regional Analysis")

        # Trend Analysis visualization
        self.trend_widget = self.create_trend_view()
        viz_tabs.addTab(self.trend_widget, "Trends")

        layout.addWidget(viz_tabs)
        return tab

    def browse_input(self):
        """Handle browsing and selecting the input Excel file."""
        try:
            path, _ = QFileDialog.getOpenFileName(
                self,
                "Select Input Excel File",
                "",
                "Excel Files (*.xlsx *.xls)"
            )
            if path:
                self.input_edit.setText(path)
                sheets = self.file_handler.load_excel_sheets(path)
                self.sheet_combo.clear()
                self.sheet_combo.addItems(sheets)
                logger.info(f"Selected input file: {path}")
        except Exception as e:
            logger.error(f"Error browsing input file: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to select input file:\n{str(e)}")

    def browse_output(self):
        """Handle browsing and selecting the output Excel file."""
        try:
            path, _ = QFileDialog.getSaveFileName(
                self,
                "Select Output Excel File",
                "",
                "Excel Files (*.xlsx *.xls)"
            )
            if path:
                self.output_edit.setText(path)
                logger.info(f"Selected output file: {path}")
        except Exception as e:
            logger.error(f"Error browsing output file: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to select output file:\n{str(e)}")

    def process_data(self):
        """Process the input data and update all visualizations."""
        try:
            if not self.validate_inputs():
                return

            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            self.statusBar().showMessage("Processing data...")

            # Read and validate input data
            df = self.file_handler.read_excel_data(
                file_path=self.input_edit.text(),
                sheet_name=self.sheet_combo.currentText()
            )
            self.progress_bar.setValue(20)

            # Apply data filters
            df = self.apply_filters(df)
            self.progress_bar.setValue(40)

            # Process based on analyzer selection
            analyzer_type = self.analyzer_combo.currentText()
            if analyzer_type == "Consolidated":
                self.process_consolidated(df)
            else:
                self.process_single_analyzer(df, analyzer_type)

            self.progress_bar.setValue(100)
            self.statusBar().showMessage("Processing complete", 5000)

        except Exception as e:
            logger.error(f"Error processing data: {str(e)}")
            QMessageBox.critical(self, "Error", f"Error processing data: {str(e)}")
            self.statusBar().showMessage("Processing failed")
        finally:
            self.progress_bar.setVisible(False)

    def process_single_analyzer(self, df: pd.DataFrame, analyzer_type: str):
        """Process data for a single analyzer type with comprehensive analysis."""
        try:
            # Perform market share analysis
            results = WorkingAggregator.analyze_market_data(
                df=df,
                config=self.config.config_data,
                analyzer_type=analyzer_type
            )

            if results is None:
                logger.warning(f"No results for analyzer type: {analyzer_type}")
                QMessageBox.warning(self, "No Data", f"No results found for {analyzer_type}.")
                return

            # Store results and update displays
            self.latest_results = results
            self.update_results_display(results, analyzer_type)
            self.save_results(results, analyzer_type)
            self.update_visualizations(results, analyzer_type)

            # Calculate and display additional metrics
            self.calculate_advanced_metrics(df, results, analyzer_type)

        except Exception as e:
            logger.error(f"Error processing {analyzer_type}: {str(e)}")
            raise

    def process_consolidated(self, df: pd.DataFrame):
        """Process data in consolidated mode for all analyzers."""
        try:
            for analyzer_type in ["IA", "CBC", "CHEM"]:
                results = WorkingAggregator.analyze_market_data(
                    df=df,
                    config=self.config.config_data,
                    analyzer_type=analyzer_type
                )

                if results is None:
                    logger.warning(f"No results for analyzer type: {analyzer_type}")
                    continue

                # Store the latest results (can be overwritten; consider storing separately if needed)
                self.latest_results = results

                # Update results display
                self.update_results_display(results, analyzer_type)
                self.save_results(results, analyzer_type)
                self.update_visualizations(results, analyzer_type)

            QMessageBox.information(self, "Processing Complete", "Consolidated processing completed successfully.")

        except Exception as e:
            logger.error(f"Error in consolidated processing: {str(e)}")
            QMessageBox.critical(self, "Error", f"Error in consolidated processing: {str(e)}")

    def update_results_display(self, results: AnalysisResult, analyzer_type: str):
        """Update the GUI with analysis results."""
        # Update volume market share
        volume_text = f"Volume Market Share ({analyzer_type}):\n"
        for brand, share in results.market_share.items():
            volume_text += f"{brand}: {share:.1f}%\n"
        self.volume_results.setText(volume_text)

        # Update value market share if available
        if results.brand_values:
            value_text = f"Value Market Share ({analyzer_type}):\n"
            value_share = WorkingAggregator.calculate_market_share(results.brand_values)
            for brand, share in value_share.items():
                value_text += f"{brand}: {share:.1f}%\n"
            self.value_results.setText(value_text)
        else:
            self.value_results.setText("Value Market Share: Not analyzed")

        # Update pivot analysis results
        if results.city_pivot is not None:
            self.city_results.setText("City Analysis: Available (see visualizations)")
        else:
            self.city_results.setText("City Analysis: Not included")

        if results.class_pivot is not None:
            self.class_results.setText("Class Analysis: Available (see visualizations)")
        else:
            self.class_results.setText("Class Analysis: Not included")

    def update_visualizations(self, results: AnalysisResult, analyzer_type: str):
        """Update visualization charts with new data."""
        # Update Market Share Chart
        self.update_market_share_chart(results, analyzer_type)

        # Update Regional Analysis Chart
        if self.city_check.isChecked() and results.city_pivot is not None:
            self.update_regional_chart(results.city_pivot, "City")
        if self.class_check.isChecked() and results.class_pivot is not None:
            self.update_regional_chart(results.class_pivot, "Class")

    def calculate_advanced_metrics(self, df: pd.DataFrame,
                                 results: AnalysisResult,
                                 analyzer_type: str):
        """Calculate advanced metrics for deeper market analysis."""
        metrics_text = f"Advanced Metrics for {analyzer_type}:\n\n"

        # Market concentration metrics
        total_market = sum(results.brand_totals.values())
        top_brands = dict(sorted(results.market_share.items(),
                         key=lambda x: x[1], reverse=True)[:3])

        metrics_text += "Market Concentration:\n"
        metrics_text += f"Top 3 Brands Market Share: {sum(top_brands.values()):.1f}%\n"

        # Regional distribution
        if "Region" in df.columns:
            region_counts = df["Region"].value_counts()
            metrics_text += f"\nRegional Distribution:\n"
            for region, count in region_counts.items():
                metrics_text += f"{region}: {count} sites\n"

        self.metrics_display.setText(metrics_text)

    def create_market_share_view(self) -> QWidget:
        """Create the market share visualization interface with interactive controls."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Visualization controls
        controls_group = QGroupBox("Display Options")
        controls_layout = QHBoxLayout()

        self.show_volume = QCheckBox("Volume Market Share")
        self.show_volume.setChecked(True)
        self.show_volume.stateChanged.connect(self.refresh_visualizations)

        self.show_value = QCheckBox("Value Market Share")
        self.show_value.setChecked(True)
        self.show_value.stateChanged.connect(self.refresh_visualizations)

        controls_layout.addWidget(self.show_volume)
        controls_layout.addWidget(self.show_value)
        controls_group.setLayout(controls_layout)
        layout.addWidget(controls_group)

        # Chart area
        self.market_share_chart_area = QWidget()
        self.market_share_layout = QVBoxLayout(self.market_share_chart_area)
        layout.addWidget(self.market_share_chart_area)

        return widget

    def refresh_visualizations(self):
        """Update all visualizations with current analysis results."""
        if self.latest_results:
            self.update_market_share_chart(
                self.latest_results,
                self.analyzer_combo.currentText()
            )
            self.update_regional_analysis()
            self.update_trend_analysis()
        else:
            self.statusBar().showMessage("No analysis results available")

    def update_market_share_chart(self, results: AnalysisResult, analyzer_type: str):
        """Create and display market share visualization charts."""
        self.clear_layout(self.market_share_layout)

        # Gather all brands from volume and value market share
        all_brands = set(results.market_share.keys())
        if results.brand_values:
            all_brands.update(results.brand_values.keys())
        all_brands = sorted(all_brands)
        colors = plt.cm.Set3(np.linspace(0, 1, len(all_brands)))
        brand_color_map = dict(zip(all_brands, colors))

        # Volume Market Share Chart
        if self.show_volume.isChecked() and results.market_share:
            fig, ax = plt.subplots(figsize=(8, 6))

            brands = list(results.market_share.keys())
            shares = list(results.market_share.values())
            volume_colors = [brand_color_map[brand] for brand in brands]

            wedges, texts, autotexts = ax.pie(
                shares,
                labels=brands,
                colors=volume_colors,
                autopct='%1.1f%%',
                startangle=90
            )

            ax.set_title(f'{analyzer_type} Volume Market Share')

            # Enhance text visibility
            plt.setp(autotexts, size=9, weight="bold")
            plt.setp(texts, size=10)

            canvas = FigureCanvas(fig)
            self.market_share_layout.addWidget(canvas)

        # Value Market Share Chart
        if self.show_value.isChecked() and results.brand_values:
            fig, ax = plt.subplots(figsize=(8, 6))

            value_share = WorkingAggregator.calculate_market_share(results.brand_values)
            value_brands = list(value_share.keys())
            value_shares = list(value_share.values())
            value_colors = [brand_color_map[brand] for brand in value_brands]

            wedges, texts, autotexts = ax.pie(
                value_shares,
                labels=value_brands,
                colors=value_colors,
                autopct='%1.1f%%',
                startangle=90
            )

            ax.set_title(f'{analyzer_type} Value Market Share')

            plt.setp(autotexts, size=9, weight="bold")
            plt.setp(texts, size=10)

            canvas = FigureCanvas(fig)
            self.market_share_layout.addWidget(canvas)

        plt.close('all')

    def create_regional_view(self) -> QWidget:
        """Create the regional analysis visualization interface."""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Analysis type selection
        controls_group = QGroupBox("Regional Analysis Controls")
        controls_layout = QFormLayout()

        self.regional_analysis_combo = QComboBox()
        self.regional_analysis_combo.addItems(['City Analysis', 'Class Analysis'])
        self.regional_analysis_combo.currentTextChanged.connect(self.update_regional_analysis)

        self.chart_type = QComboBox()
        self.chart_type.addItems(['Bar Chart', 'Pie Chart', 'Heatmap'])
        self.chart_type.currentTextChanged.connect(self.update_regional_analysis)

        controls_layout.addRow("Analysis Type:", self.regional_analysis_combo)
        controls_layout.addRow("Chart Type:", self.chart_type)

        controls_group.setLayout(controls_layout)
        layout.addWidget(controls_group)

        # Chart area
        self.regional_chart_area = QWidget()
        self.regional_chart_layout = QVBoxLayout(self.regional_chart_area)
        layout.addWidget(self.regional_chart_area)

        return tab

    def update_regional_analysis(self):
        """Update regional analysis visualization based on selected options."""
        self.clear_layout(self.regional_chart_layout)

        if not self.latest_results:
            self.regional_chart_layout.addWidget(
                QLabel("Please process data to view regional analysis")
            )
            return

        analysis_type = self.regional_analysis_combo.currentText()
        chart_type = self.chart_type.currentText()

        try:
            if analysis_type == 'City Analysis':
                data = self.latest_results.city_pivot
                category = 'City'
            else:
                data = self.latest_results.class_pivot
                category = 'Class'

            if data is None or data.empty:
                self.regional_chart_layout.addWidget(
                    QLabel(f"No {analysis_type} data available")
                )
                return

            fig = self.create_regional_chart(data, category, chart_type)
            canvas = FigureCanvas(fig)
            self.regional_chart_layout.addWidget(canvas)

        except Exception as e:
            logger.error(f"Error updating regional analysis: {str(e)}")
            self.regional_chart_layout.addWidget(
                QLabel(f"Error creating visualization: {str(e)}")
            )

    def create_regional_chart(self, data: pd.DataFrame,
                              category: str, chart_type: str) -> Figure:
        """Create regional analysis chart based on selected type."""
        fig, ax = plt.subplots(figsize=(10, 6))

        if chart_type == 'Bar Chart':
            data.plot(kind='bar', ax=ax)
            plt.xticks(rotation=45)
        elif chart_type == 'Pie Chart':
            # For Pie Chart, sum across brands if multiple
            if 'BRAND' in data.columns:
                # Sum allocated yearly per category
                summed = data.groupby(category).sum()
                summed.plot(kind='pie', y=summed.columns[0], autopct='%1.1f%%', ax=ax)
            else:
                data.plot(kind='pie', autopct='%1.1f%%', ax=ax)
        else:  # Heatmap
            sns.heatmap(data, annot=True, fmt='.1f', ax=ax)

        ax.set_title(f'{category} Distribution Analysis')
        plt.tight_layout()

        return fig

    def create_trend_view(self) -> QWidget:
        """Create the trend analysis visualization interface."""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Trend analysis controls
        controls_group = QGroupBox("Trend Analysis Controls")
        controls_layout = QFormLayout()

        self.trend_period = QComboBox()
        self.trend_period.addItems(['Monthly', 'Quarterly', 'Yearly'])
        self.trend_period.currentTextChanged.connect(self.update_trend_analysis)

        self.trend_metric = QComboBox()
        self.trend_metric.addItems([
            'Market Share',
            'Total Volume',
            'Value Distribution'
        ])
        self.trend_metric.currentTextChanged.connect(self.update_trend_analysis)

        controls_layout.addRow("Time Period:", self.trend_period)
        controls_layout.addRow("Metric:", self.trend_metric)

        controls_group.setLayout(controls_layout)
        layout.addWidget(controls_group)

        # Chart area
        self.trend_chart_area = QWidget()
        self.trend_chart_layout = QVBoxLayout(self.trend_chart_area)
        layout.addWidget(self.trend_chart_area)

        return tab

    def update_trend_analysis(self):
        """Update the trend analysis chart based on selected time period and metric."""
        self.clear_layout(self.trend_chart_layout)
        # Implement trend analysis visualization based on 'self.latest_results'
        if not self.latest_results:
            self.trend_chart_layout.addWidget(
                QLabel("Please process data to view trend analysis")
            )
            return

        # Placeholder implementation
        trend_period = self.trend_period.currentText()
        trend_metric = self.trend_metric.currentText()

        try:
            if trend_metric == 'Market Share':
                data = self.latest_results.market_share
                fig, ax = plt.subplots(figsize=(10, 6))
                sns.lineplot(x=list(range(len(data))), y=list(data.values()), marker='o', ax=ax)
                ax.set_title(f'{trend_period} Market Share Trend')
                ax.set_xlabel('Time')
                ax.set_ylabel('Market Share (%)')
                plt.tight_layout()
                canvas = FigureCanvas(fig)
                self.trend_chart_layout.addWidget(canvas)

            elif trend_metric == 'Total Volume':
                data = self.latest_results.brand_totals
                fig, ax = plt.subplots(figsize=(10, 6))
                sns.lineplot(x=list(range(len(data))), y=list(data.values()), marker='o', ax=ax)
                ax.set_title(f'{trend_period} Total Volume Trend')
                ax.set_xlabel('Time')
                ax.set_ylabel('Total Volume')
                plt.tight_layout()
                canvas = FigureCanvas(fig)
                self.trend_chart_layout.addWidget(canvas)

            elif trend_metric == 'Value Distribution':
                data = self.latest_results.brand_values
                fig, ax = plt.subplots(figsize=(10, 6))
                sns.lineplot(x=list(range(len(data))), y=list(data.values()), marker='o', ax=ax)
                ax.set_title(f'{trend_period} Value Distribution Trend')
                ax.set_xlabel('Time')
                ax.set_ylabel('Value Distribution')
                plt.tight_layout()
                canvas = FigureCanvas(fig)
                self.trend_chart_layout.addWidget(canvas)

            plt.close('all')

        except Exception as e:
            logger.error(f"Error updating trend analysis: {str(e)}")
            self.trend_chart_layout.addWidget(
                QLabel(f"Error creating trend visualization: {str(e)}")
            )

    def validate_inputs(self) -> bool:
        """Validate all input parameters before processing."""
        if not self.input_edit.text():
            QMessageBox.warning(self, "Validation Error", "Please select an input file")
            return False

        if not self.output_edit.text():
            QMessageBox.warning(self, "Validation Error", "Please select an output file")
            return False

        if not self.sheet_combo.currentText():
            QMessageBox.warning(self, "Validation Error", "Please select a worksheet")
            return False

        return True

    def apply_filters(self, df: pd.DataFrame) -> pd.DataFrame:
        """Apply selected filters to the dataset."""
        # Region filter
        selected_region = self.region_combo.currentText()
        if selected_region != "All" and "Region" in df.columns:
            df = df[df["Region"] == selected_region]

        return df

    def clear_layout(self, layout):
        """Clear all widgets from a layout."""
        if layout is not None:
            while layout.count():
                item = layout.takeAt(0)
                widget = item.widget()
                if widget is not None:
                    widget.deleteLater()

    def update_regional_chart(self, pivot_df: pd.DataFrame, analysis_type: str):
        """Update the regional chart with the provided pivot DataFrame."""
        self.clear_layout(self.regional_chart_layout)

        if pivot_df is None or pivot_df.empty:
            self.regional_chart_layout.addWidget(QLabel("No data available for this analysis."))
            return

        fig, ax = plt.subplots(figsize=(10, 6))
        pivot_df.plot(kind='bar', ax=ax)
        ax.set_title(f'{analysis_type} Distribution')
        ax.set_xlabel(analysis_type)
        ax.set_ylabel('Allocated Yearly')

        plt.tight_layout()
        canvas = FigureCanvas(fig)
        self.regional_chart_layout.addWidget(canvas)

        plt.close('all')

    def update_trend_analysis(self):
        """Placeholder method if additional processing is needed."""
        pass  # Already handled in `update_trend_analysis` method above

    def update_regional_analysis(self):
        """Placeholder method to ensure regional analysis updates correctly."""
        # This method can be implemented similarly to update_trend_analysis if needed
        pass

    def save_results(self, results: AnalysisResult, analyzer_type: str):
        """Export analysis results to Excel with formatted worksheets."""
        try:
            if xlsxwriter is None:
                raise ImportError("xlsxwriter module is not installed.")

            output_path = self.output_edit.text()

            # Create Excel writer object
            with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                workbook = writer.book

                # Define formats
                header_format = workbook.add_format({
                    'bold': True,
                    'bg_color': '#4F81BD',
                    'font_color': 'white',
                    'border': 1
                })

                number_format = workbook.add_format({
                    'num_format': '#,##0.00',
                    'border': 1
                })

                percent_format = workbook.add_format({
                    'num_format': '0.00%',
                    'border': 1
                })

                # Create summary dataframe
                summary_data = {
                    'Brand': list(results.market_share.keys()),
                    'Market Share (%)': list(results.market_share.values()),
                    'Total Volume': [results.brand_totals[brand] for brand in results.market_share.keys()]
                }

                if results.brand_values:
                    summary_data['Value'] = [results.brand_values[brand] for brand in results.market_share.keys()]

                summary_df = pd.DataFrame(summary_data)

                # Write summary sheet
                summary_df.to_excel(writer, sheet_name=f'{analyzer_type}_Summary', index=False)

                # Get worksheet object
                summary_ws = writer.sheets[f'{analyzer_type}_Summary']

                # Apply header format
                for col_num, value in enumerate(summary_df.columns.values):
                    summary_ws.write(0, col_num, value, header_format)

                # Apply number formats
                summary_ws.set_column('B:B', 18, percent_format)
                summary_ws.set_column('C:C', 15, number_format)
                if results.brand_values:
                    summary_ws.set_column('D:D', 15, number_format)

                # Add pivot tables if available
                if results.city_pivot is not None:
                    results.city_pivot.to_excel(writer, sheet_name=f'{analyzer_type}_City_Analysis', index=False)
                    city_ws = writer.sheets[f'{analyzer_type}_City_Analysis']
                    for col_num, value in enumerate(results.city_pivot.columns.values):
                        city_ws.write(0, col_num, value, header_format)
                    city_ws.set_column('B:B', 18, number_format)

                if results.class_pivot is not None:
                    results.class_pivot.to_excel(writer, sheet_name=f'{analyzer_type}_Class_Analysis', index=False)
                    class_ws = writer.sheets[f'{analyzer_type}_Class_Analysis']
                    for col_num, value in enumerate(results.class_pivot.columns.values):
                        class_ws.write(0, col_num, value, header_format)
                    class_ws.set_column('B:B', 18, number_format)

            self.statusBar().showMessage(f"Results saved to {output_path}", 5000)
            logger.info(f"Results saved to {output_path}")

            QMessageBox.information(self, "Success", f"Results saved to {output_path}")

        except ImportError as ie:
            logger.error(f"Error saving results: {ie}")
            QMessageBox.critical(self, "Missing Dependency", "The 'xlsxwriter' module is required but not installed.\nPlease install it using 'pip install XlsxWriter'.")
        except Exception as e:
            logger.error(f"Error saving results: {e}")
            QMessageBox.critical(self, "Error", f"Failed to save results: {str(e)}")

    def open_settings(self):
        """Open the settings dialog."""
        dialog = SettingsDialog(self.config)
        if dialog.exec_():
            # Reload configuration if needed
            self.config.load_config()
            # Update region_combo in case regions have changed
            regions = self.config.config_data.get("metadata", {}).get("regions", ["Region1", "Region2"])
            self.region_combo.clear()
            self.region_combo.addItems(["All"] + regions)
            QMessageBox.information(self, "Settings Updated", "Settings have been updated successfully.")

def setup_application():
    """Initialize the application with proper error handling."""
    try:
        app = QApplication(sys.argv)
        app.setStyle('Fusion')  # Apply modern style

        # Create and show the main window
        window = ModernMarketAnalyzer()
        window.show()

        return app, window
    except Exception as e:
        logger.error(f"Error initializing application: {e}")
        raise

def main():
    """Main entry point for the Market Share Analysis Tool."""
    try:
        # Configure logging (already configured at the top)
        logger.info("Starting Market Share Analysis Tool")

        # Initialize application
        app, window = setup_application()

        # Execute application
        sys.exit(app.exec_())

    except Exception as e:
        logger.error(f"Application failed to start: {e}")
        print(f"Error starting application: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
