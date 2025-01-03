# modern_dashboard.py
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QComboBox, QPushButton, QFrame, QScrollArea, QTabWidget,
    QGroupBox, QSplitter
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QPalette, QColor
from visualization import MarketShareVisualizer
from modern_ui import ModernFrame, ModernButton
import pandas as pd

class AnalysisDashboard(QMainWindow):
    """Main dashboard window for market share analysis."""

    def __init__(self, config_manager):
        super().__init__()
        self.config_manager = config_manager
        self.visualizer = MarketShareVisualizer()
        self.init_ui()

    def init_ui(self):
        """Initialize the dashboard UI."""
        self.setWindowTitle("Market Share Analysis Dashboard")
        self.setMinimumSize(1400, 800)

        # Create central widget with main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Add control panel
        control_panel = self.create_control_panel()
        main_layout.addWidget(control_panel)

        # Create tab widget for different views
        tab_widget = QTabWidget()
        tab_widget.addTab(self.create_overview_tab(), "Overview")
        tab_widget.addTab(self.create_detailed_analysis_tab(), "Detailed Analysis")
        tab_widget.addTab(self.create_regional_analysis_tab(), "Regional Analysis")
        tab_widget.addTab(self.create_trends_tab(), "Trends")

        main_layout.addWidget(tab_widget)

    def create_control_panel(self) -> QWidget:
        """Create the control panel with analysis options."""
        panel = ModernFrame("Analysis Controls")
        layout = QHBoxLayout()

        # Analyzer selection
        analyzer_group = QGroupBox("Analyzer")
        analyzer_layout = QVBoxLayout()
        self.analyzer_combo = QComboBox()
        self.analyzer_combo.addItems(["IA", "CBC", "CHEM", "Consolidated"])
        self.analyzer_combo.currentTextChanged.connect(self.update_dashboard)
        analyzer_layout.addWidget(self.analyzer_combo)
        analyzer_group.setLayout(analyzer_layout)
        layout.addWidget(analyzer_group)

        # Region selection
        region_group = QGroupBox("Region")
        region_layout = QVBoxLayout()
        self.region_combo = QComboBox()
        self.region_combo.addItems(["All"] + self.config_manager.config_data["metadata"]["regions"])
        self.region_combo.currentTextChanged.connect(self.update_dashboard)
        region_layout.addWidget(self.region_combo)
        region_group.setLayout(region_layout)
        layout.addWidget(region_group)

        # Time period selection
        period_group = QGroupBox("Time Period")
        period_layout = QVBoxLayout()
        self.period_combo = QComboBox()
        self.period_combo.addItems(["Current", "Last Month", "Last Quarter", "YTD"])
        period_layout.addWidget(self.period_combo)
        period_group.setLayout(period_layout)
        layout.addWidget(period_group)

        # Action buttons
        button_group = QGroupBox("Actions")
        button_layout = QHBoxLayout()

        refresh_btn = ModernButton("Refresh")
        refresh_btn.clicked.connect(self.refresh_data)

        export_btn = ModernButton("Export Report")
        export_btn.clicked.connect(self.export_report)

        button_layout.addWidget(refresh_btn)
        button_layout.addWidget(export_btn)
        button_group.setLayout(button_layout)
        layout.addWidget(button_group)

        panel.setLayout(layout)
        return panel

    def create_overview_tab(self) -> QWidget:
        """Create the overview dashboard tab."""
        tab = QWidget()
        layout = QGridLayout(tab)

        # Market Share Overview
        market_share_frame = ModernFrame("Market Share Overview")
        self.market_share_widget = QWidget()
        market_share_frame.layout().addWidget(self.market_share_widget)
        layout.addWidget(market_share_frame, 0, 0)

        # Regional Distribution
        regional_frame = ModernFrame("Regional Distribution")
        self.regional_widget = QWidget()
        regional_frame.layout().addWidget(self.regional_widget)
        layout.addWidget(regional_frame, 0, 1)

        # Brand Performance
        brand_frame = ModernFrame("Brand Performance")
        self.brand_widget = QWidget()
        brand_frame.layout().addWidget(self.brand_widget)
        layout.addWidget(brand_frame, 1, 0)

        # Class Distribution
        class_frame = ModernFrame("Class Distribution")
        self.class_widget = QWidget()
        class_frame.layout().addWidget(self.class_widget)
        layout.addWidget(class_frame, 1, 1)

        return tab

    def create_detailed_analysis_tab(self) -> QWidget:
        """Create the detailed analysis tab."""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        splitter = QSplitter(Qt.Horizontal)

        # Data grid
        data_frame = ModernFrame("Detailed Data")
        self.data_widget = QWidget()
        data_frame.layout().addWidget(self.data_widget)
        splitter.addWidget(data_frame)

        # Analysis charts
        chart_frame = ModernFrame("Analysis Charts")
        self.chart_widget = QWidget()
        chart_frame.layout().addWidget(self.chart_widget)
        splitter.addWidget(chart_frame)

        layout.addWidget(splitter)
        return tab

    def update_dashboard(self):
        """Update all dashboard components with new data."""
        try:
            analyzer_type = self.analyzer_combo.currentText()
            region = self.region_combo.currentText()

            # Update visualizations
            self.update_market_share_chart(analyzer_type, region)
            self.update_regional_chart(analyzer_type, region)
            self.update_brand_chart(analyzer_type, region)
            self.update_class_chart(analyzer_type, region)

        except Exception as e:
            self.show_error_message(f"Error updating dashboard: {str(e)}")

    def show_error_message(self, message: str):
        """Show error message to user."""
        from PyQt5.QtWidgets import QMessageBox
        QMessageBox.critical(self, "Error", message)