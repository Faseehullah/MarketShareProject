# main.py

import sys
import os
import logging
from datetime import datetime
from typing import Optional, Dict, Any

# Data processing imports
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

# PyQt imports
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QComboBox, QMessageBox, QLineEdit,
    QFormLayout, QCheckBox, QSpinBox, QTabWidget, QGroupBox, QScrollArea
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

# Local imports
from aggregator import WorkingAggregator, AnalysisResult
from config import MarketAnalysisConfig

logger = logging.getLogger(__name__)

def load_sheet_names(file_path: str) -> list:
    """Load and return all sheet names from an Excel file."""
    try:
        xls = pd.ExcelFile(file_path)
        return xls.sheet_names
    except Exception as e:
        logger.exception("Error reading sheet names.")
        QMessageBox.critical(None, "Error", f"Failed to read sheet names: {str(e)}")
        return []

class ModernMarketAnalyzer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config = MarketAnalysisConfig()
        self.setup_logger()
        self.latest_results: Optional[AnalysisResult] = None  # To store the latest analysis results
        self.init_ui()

    def setup_logger(self):
        """Configure logging for the application."""
        logger.setLevel(logging.INFO)
        handler = logging.StreamHandler(sys.stdout)
        handler.setFormatter(logging.Formatter(
            '[%(asctime)s] %(levelname)s: %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        ))
        logger.addHandler(handler)

    def init_ui(self):
        """Initialize the main user interface."""
        self.setWindowTitle("Market Share Analysis Tool")
        self.setGeometry(100, 100, 1200, 800)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)

        # Create tabs for different sections
        tab_widget = QTabWidget()
        tab_widget.addTab(self.create_input_tab(), "Data Input")
        tab_widget.addTab(self.create_analysis_tab(), "Analysis")
        tab_widget.addTab(self.create_visualization_tab(), "Visualization")

        main_layout.addWidget(tab_widget)

        # Status bar for feedback
        self.statusBar().showMessage("Ready")

        self.apply_modern_style()

    def create_input_tab(self) -> QWidget:
        """Create the data input tab."""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # File selection group
        file_group = QGroupBox("File Selection")
        file_layout = QFormLayout()

        # Input file selection
        self.input_edit = QLineEdit()
        browse_input_btn = QPushButton("Browse")
        browse_input_btn.clicked.connect(self.browse_input)
        input_layout = QHBoxLayout()
        input_layout.addWidget(self.input_edit)
        input_layout.addWidget(browse_input_btn)
        file_layout.addRow("Input Excel:", input_layout)

        # Sheet selection
        self.sheet_combo = QComboBox()
        file_layout.addRow("Sheet:", self.sheet_combo)

        # Output file selection
        self.output_edit = QLineEdit()
        browse_output_btn = QPushButton("Browse")
        browse_output_btn.clicked.connect(self.browse_output)
        output_layout = QHBoxLayout()
        output_layout.addWidget(self.output_edit)
        output_layout.addWidget(browse_output_btn)
        file_layout.addRow("Output Excel:", output_layout)

        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        # Analysis options group
        analysis_group = QGroupBox("Analysis Options")
        analysis_layout = QFormLayout()

        # Analyzer selection
        self.analyzer_combo = QComboBox()
        self.analyzer_combo.addItems(["IA", "CBC", "CHEM", "Consolidated"])
        analysis_layout.addRow("Analyzer:", self.analyzer_combo)

        # Region filter
        self.region_combo = QComboBox()
        regions = self.config.config_data.get("metadata", {}).get("regions", ["Region1", "Region2"])  # Default regions if not set
        self.region_combo.addItems(["All"] + regions)
        analysis_layout.addRow("Region:", self.region_combo)

        # Additional options
        self.city_check = QCheckBox("Include City Analysis")
        self.class_check = QCheckBox("Include Class Analysis")
        self.value_check = QCheckBox("Include Value Analysis")

        analysis_layout.addRow(self.city_check)
        analysis_layout.addRow(self.class_check)
        analysis_layout.addRow(self.value_check)

        analysis_group.setLayout(analysis_layout)
        layout.addWidget(analysis_group)

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
        """Create the analysis tab with real-time results."""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Results display area
        self.results_group = QGroupBox("Analysis Results")
        results_layout = QVBoxLayout()

        # Market share results
        self.volume_results = QLabel("Volume Market Share: Not analyzed yet")
        self.value_results = QLabel("Value Market Share: Not analyzed yet")
        self.city_results = QLabel("City Analysis: Not available")
        self.class_results = QLabel("Class Analysis: Not available")

        results_layout.addWidget(self.volume_results)
        results_layout.addWidget(self.value_results)
        results_layout.addWidget(self.city_results)
        results_layout.addWidget(self.class_results)

        self.results_group.setLayout(results_layout)
        layout.addWidget(self.results_group)

        return tab

    def create_visualization_tab(self) -> QWidget:
        """Create the visualization tab with charts and graphs."""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Create sub-tabs for different visualizations
        viz_tabs = QTabWidget()

        # Market Share Charts
        self.market_share_widget = self.create_market_share_view()
        viz_tabs.addTab(self.market_share_widget, "Market Share")

        # Regional Analysis
        self.regional_widget = self.create_regional_view()
        viz_tabs.addTab(self.regional_widget, "Regional Analysis")

        # Trend Analysis
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
                sheets = load_sheet_names(path)
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
        """Process the input data and update results."""
        try:
            if not self.validate_inputs():
                return

            self.statusBar().showMessage("Processing data...")

            # Read input data
            df = self.read_input_data()
            if df is None:
                return

            # Apply filters
            df = self.apply_filters(df)

            # Process based on analyzer type
            analyzer_type = self.analyzer_combo.currentText()
            if analyzer_type == "Consolidated":
                self.process_consolidated(df)
            else:
                self.process_single_analyzer(df, analyzer_type)

            self.statusBar().showMessage("Processing complete", 5000)

        except Exception as e:
            logger.error(f"Error processing data: {str(e)}")
            QMessageBox.critical(self, "Error", f"Error processing data: {str(e)}")
            self.statusBar().showMessage("Processing failed")

    def process_single_analyzer(self, df: pd.DataFrame, analyzer_type: str):
        """Process data for a single analyzer type."""
        try:
            # Get analysis results
            results = WorkingAggregator.analyze_market_data(
                df=df,
                config=self.config.config_data,
                analyzer_type=analyzer_type
            )

            # Store the latest results
            self.latest_results = results

            # Update results display
            self.update_results_display(results, analyzer_type)

            # Save results
            self.save_results(results, analyzer_type)

            # Update visualizations
            self.update_visualizations(results, analyzer_type)

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

                # Save results
                self.save_results(results, analyzer_type)

                # Update visualizations
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

    def update_market_share_chart(self, results: AnalysisResult, analyzer_type: str):
        """Update the market share pie chart."""
        self.clear_layout(self.market_share_layout)

        # Volume Market Share
        if self.show_volume.isChecked() and results.market_share:
            fig1, ax1 = plt.subplots(figsize=(6, 6))
            brands = list(results.market_share.keys())
            shares = list(results.market_share.values())
            ax1.pie(shares, labels=brands, autopct='%1.1f%%', startangle=140)
            ax1.set_title(f'{analyzer_type} Volume Market Share')
            canvas1 = FigureCanvas(fig1)
            self.market_share_layout.addWidget(canvas1)

        # Value Market Share
        if self.show_value.isChecked() and results.brand_values:
            value_share = WorkingAggregator.calculate_market_share(results.brand_values)
            if value_share:
                fig2, ax2 = plt.subplots(figsize=(6, 6))
                brands = list(value_share.keys())
                shares = list(value_share.values())
                ax2.pie(shares, labels=brands, autopct='%1.1f%%', startangle=140)
                ax2.set_title(f'{analyzer_type} Value Market Share')
                canvas2 = FigureCanvas(fig2)
                self.market_share_layout.addWidget(canvas2)

        plt.close('all')

    def create_market_share_view(self) -> QWidget:
        """Create the market share visualization widget."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # Add chart options
        options_group = QGroupBox("Display Options")
        options_layout = QHBoxLayout()

        self.show_volume = QCheckBox("Volume Share")
        self.show_volume.setChecked(True)
        self.show_volume.stateChanged.connect(self.toggle_market_share_visibility)
        self.show_value = QCheckBox("Value Share")
        self.show_value.setChecked(True)
        self.show_value.stateChanged.connect(self.toggle_market_share_visibility)

        options_layout.addWidget(self.show_volume)
        options_layout.addWidget(self.show_value)
        options_group.setLayout(options_layout)

        layout.addWidget(options_group)

        # Placeholder for charts
        self.market_share_chart_area = QWidget()
        self.market_share_layout = QVBoxLayout(self.market_share_chart_area)
        layout.addWidget(self.market_share_chart_area)

        return widget

    def create_regional_view(self) -> QWidget:
        """Create the regional analysis visualization widget."""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Analysis type selection
        analysis_group = QGroupBox("Regional Analysis Options")
        analysis_layout = QFormLayout()

        self.regional_analysis_combo = QComboBox()
        self.regional_analysis_combo.addItems(['City Analysis', 'Class Analysis'])
        self.regional_analysis_combo.currentTextChanged.connect(self.update_regional_chart_view)

        analysis_layout.addRow("Analysis Type:", self.regional_analysis_combo)
        analysis_group.setLayout(analysis_layout)

        layout.addWidget(analysis_group)

        # Placeholder for regional chart
        self.regional_chart_area = QWidget()
        self.regional_chart_layout = QVBoxLayout(self.regional_chart_area)
        layout.addWidget(self.regional_chart_area)

        return tab

    def create_trend_view(self) -> QWidget:
        """Create the trend analysis visualization widget."""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Trend options
        trend_group = QGroupBox("Trend Analysis Options")
        trend_layout = QFormLayout()

        self.trend_type_combo = QComboBox()
        self.trend_type_combo.addItems(['Monthly', 'Quarterly', 'Yearly'])
        self.trend_type_combo.currentTextChanged.connect(self.update_trend_chart)

        trend_layout.addRow("Time Period:", self.trend_type_combo)
        trend_group.setLayout(trend_layout)
        layout.addWidget(trend_group)

        # Placeholder for trend chart
        self.trend_chart_area = QWidget()
        self.trend_chart_layout = QVBoxLayout(self.trend_chart_area)
        layout.addWidget(self.trend_chart_area)

        return tab

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

    def read_input_data(self) -> Optional[pd.DataFrame]:
        """Read and validate input Excel file."""
        try:
            df = pd.read_excel(
                self.input_edit.text(),
                sheet_name=self.sheet_combo.currentText()
            )

            # Replace null values
            df = df.replace(["NILL", "Nill", "nill", "NIL"], np.nan)

            return df

        except Exception as e:
            logger.error(f"Error reading input file: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to read input file: {str(e)}")
            return None

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

    def apply_modern_style(self):
        """Apply modern styling to the application."""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f2f5;
            }
            QTabWidget::pane {
                border: 1px solid #cccccc;
                background: white;
                border-radius: 4px;
            }
            QTabBar::tab {
                background: #e0e0e0;
                padding: 8px 16px;
                margin-right: 2px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }
            QTabBar::tab:selected {
                background: white;
                border-bottom: none;
            }
            QPushButton {
                background-color: #1976d2;
                color: white;
                padding: 8px 16px;
                border: none;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1565c0;
            }
            QGroupBox {
                background-color: white;
                border: 1px solid #cccccc;
                border-radius: 4px;
                margin-top: 1em;
                padding-top: 1em;
            }
            QLineEdit, QComboBox {
                padding: 6px;
                border: 1px solid #cccccc;
                border-radius: 4px;
            }
        """)

    def toggle_market_share_visibility(self):
        """Toggle the visibility of volume and value market share charts."""
        if self.latest_results:
            self.update_market_share_chart(self.latest_results, self.analyzer_combo.currentText())
        else:
            QMessageBox.warning(self, "No Data", "Please process data before toggling visualization options.")

    def update_regional_chart_view(self, analysis_type: str):
        """Update the regional analysis chart based on selected type."""
        self.clear_layout(self.regional_chart_layout)

        if not self.latest_results:
            QMessageBox.warning(self, "No Data", "Please process data before viewing visualizations.")
            return

        if analysis_type == 'City Analysis':
            if self.latest_results.city_pivot is not None and not self.latest_results.city_pivot.empty:
                fig, ax = plt.subplots(figsize=(10, 6))
                self.latest_results.city_pivot.plot(kind='bar', ax=ax)
                ax.set_title('City Analysis Distribution')
                ax.set_xlabel('City')
                ax.set_ylabel('Allocated Yearly')
                plt.tight_layout()
                canvas = FigureCanvas(fig)
                self.regional_chart_layout.addWidget(canvas)
                plt.close('all')
            else:
                self.regional_chart_layout.addWidget(QLabel("No City Analysis data available."))
        elif analysis_type == 'Class Analysis':
            if self.latest_results.class_pivot is not None and not self.latest_results.class_pivot.empty:
                fig, ax = plt.subplots(figsize=(10, 6))
                self.latest_results.class_pivot.plot(kind='bar', ax=ax)
                ax.set_title('Class Analysis Distribution')
                ax.set_xlabel('Class')
                ax.set_ylabel('Allocated Yearly')
                plt.tight_layout()
                canvas = FigureCanvas(fig)
                self.regional_chart_layout.addWidget(canvas)
                plt.close('all')
            else:
                self.regional_chart_layout.addWidget(QLabel("No Class Analysis data available."))

    def update_trend_chart(self):
        """Update the trend analysis chart based on selected time period."""
        self.clear_layout(self.trend_chart_layout)
        # Implement trend analysis visualization based on 'self.latest_results'
        if not self.latest_results:
            QMessageBox.warning(self, "No Data", "Please process data before viewing trend visualizations.")
            return

        # Placeholder implementation
        trend_type = self.trend_type_combo.currentText()
        if trend_type == 'Monthly':
            # Implement monthly trend analysis
            QMessageBox.information(self, "Info", "Monthly Trend Analysis not implemented yet.")
        elif trend_type == 'Quarterly':
            # Implement quarterly trend analysis
            QMessageBox.information(self, "Info", "Quarterly Trend Analysis not implemented yet.")
        elif trend_type == 'Yearly':
            # Implement yearly trend analysis
            QMessageBox.information(self, "Info", "Yearly Trend Analysis not implemented yet.")

    def save_results(self, results: AnalysisResult, analyzer_type: str):
        """Save the analysis results to the output Excel file."""
        try:
            output_path = self.output_edit.text()
            aggregator = WorkingAggregator()
            aggregator.save_analysis_results(results, output_path, analyzer_type)
            logger.info(f"Results saved to {output_path}")
            QMessageBox.information(self, "Success", f"Results saved to {output_path}")
        except Exception as e:
            logger.error(f"Error saving results: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to save results:\n{str(e)}")

    def open_settings(self):
        """Open the settings dialog."""
        # Implement your settings dialog logic here
        QMessageBox.information(self, "Settings", "Settings dialog not implemented yet.")

def main():
    print("DEBUG: About to create QApplication")
    logging.info("Launching Market Share Analysis Tool.")
    app = QApplication(sys.argv)
    window = ModernMarketAnalyzer()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
