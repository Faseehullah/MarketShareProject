# settings_dialog.py
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTabWidget, QWidget,
    QLabel, QLineEdit, QPushButton, QFormLayout, QSpinBox,
    QDoubleSpinBox, QGroupBox, QMessageBox, QScrollArea
)
from PyQt5.QtCore import Qt
from typing import Dict
from config import AnalyzerConfig  # Adjust based on project structure
import logging

logger = logging.getLogger(__name__)

class AnalyzerSettingsGroup(QGroupBox):
    """Group box for analyzer-specific settings."""
    def __init__(self, analyzer_type: str, config: Dict, parent=None):
        super().__init__(f"{analyzer_type} Settings", parent)
        self.analyzer_type = analyzer_type
        self.config = config
        self.init_ui()

    def init_ui(self):
        layout = QFormLayout()

        # Brand columns
        self.brand_inputs = []
        for i in range(len(self.config["brand_cols"])):  # Dynamic based on config
            line_edit = QLineEdit()
            line_edit.setText(self.config["brand_cols"][i])
            self.brand_inputs.append(line_edit)
            layout.addRow(f"Brand {i+1}:", line_edit)

        # Workload columns
        self.workload_inputs = []
        for i in range(len(self.config["workload_cols"])):  # Dynamic based on config
            line_edit = QLineEdit()
            line_edit.setText(self.config["workload_cols"][i])
            self.workload_inputs.append(line_edit)
            layout.addRow(f"Workload {i+1}:", line_edit)

        # Cost per Test
        self.price_spin = QDoubleSpinBox()
        self.price_spin.setRange(0, 10000)
        self.price_spin.setValue(self.config.get("test_price", 0))
        self.price_spin.setDecimals(2)
        layout.addRow("Cost per Test:", self.price_spin)

        # Add buttons to add/remove brand and workload columns
        button_layout = QHBoxLayout()
        add_brand_btn = QPushButton("Add Brand")
        add_brand_btn.clicked.connect(self.add_brand_field)
        remove_brand_btn = QPushButton("Remove Brand")
        remove_brand_btn.clicked.connect(self.remove_brand_field)
        button_layout.addWidget(add_brand_btn)
        button_layout.addWidget(remove_brand_btn)
        layout.addRow(button_layout)

        add_workload_btn = QPushButton("Add Workload")
        add_workload_btn.clicked.connect(self.add_workload_field)
        remove_workload_btn = QPushButton("Remove Workload")
        remove_workload_btn.clicked.connect(self.remove_workload_field)
        button_layout_workload = QHBoxLayout()
        button_layout_workload.addWidget(add_workload_btn)
        button_layout_workload.addWidget(remove_workload_btn)
        layout.addRow(button_layout_workload)

        self.setLayout(layout)

    def add_brand_field(self):
        """Add a new brand input field."""
        index = len(self.brand_inputs) + 1
        line_edit = QLineEdit()
        self.brand_inputs.append(line_edit)
        self.layout().addRow(f"Brand {index}:", line_edit)

    def remove_brand_field(self):
        """Remove the last brand input field."""
        if self.brand_inputs:
            line_edit = self.brand_inputs.pop()
            self.layout().removeRow(line_edit)
            line_edit.deleteLater()

    def add_workload_field(self):
        """Add a new workload input field."""
        index = len(self.workload_inputs) + 1
        line_edit = QLineEdit()
        self.workload_inputs.append(line_edit)
        self.layout().addRow(f"Workload {index}:", line_edit)

    def remove_workload_field(self):
        """Remove the last workload input field."""
        if self.workload_inputs:
            line_edit = self.workload_inputs.pop()
            self.layout().removeRow(line_edit)
            line_edit.deleteLater()

    def get_settings(self) -> Dict:
        """Get current settings from the group."""
        # Enforce at least one brand and one workload column
        brand_cols = [edit.text().strip() for edit in self.brand_inputs if edit.text().strip()]
        workload_cols = [edit.text().strip() for edit in self.workload_inputs if edit.text().strip()]

        if not brand_cols:
            raise ValueError(f"At least one brand column is required for {self.analyzer_type}.")
        if not workload_cols:
            raise ValueError(f"At least one workload column is required for {self.analyzer_type}.")

        return {
            "brand_cols": brand_cols,
            "workload_cols": workload_cols,
            "test_price": self.price_spin.value()
        }

class SettingsDialog(QDialog):
    """Dialog to view and edit application settings."""

    def __init__(self, config_manager, parent=None):
        super().__init__(parent)
        self.config_manager = config_manager
        self.analyzer_groups = {}
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Market Analysis Settings")
        self.setMinimumWidth(800)
        self.setMinimumHeight(600)

        layout = QVBoxLayout(self)

        # Create tab widget
        tab_widget = QTabWidget()

        # Analyzers tab
        analyzer_tab = self.create_analyzer_tab()
        tab_widget.addTab(analyzer_tab, "Analyzers")

        # General settings tab
        general_tab = self.create_general_tab()
        tab_widget.addTab(general_tab, "General Settings")

        # Add tabs to layout
        layout.addWidget(tab_widget)

        # Buttons
        button_layout = QHBoxLayout()
        reset_button = QPushButton("Reset to Default")
        reset_button.clicked.connect(self.reset_to_default)
        save_button = QPushButton("Save")
        save_button.clicked.connect(self.save_settings)
        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)

        button_layout.addWidget(reset_button)
        button_layout.addStretch()
        button_layout.addWidget(save_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)

    def create_analyzer_tab(self) -> QWidget:
        """Create the analyzers configuration tab."""
        scroll = QScrollArea()
        widget = QWidget()
        layout = QVBoxLayout(widget)

        headers = self.config_manager.get_headers()
        for analyzer_type, config in headers.items():
            group = AnalyzerSettingsGroup(analyzer_type, config.__dict__)
            self.analyzer_groups[analyzer_type] = group
            layout.addWidget(group)

        # Add stretch to push groups to the top
        layout.addStretch()

        scroll.setWidget(widget)
        scroll.setWidgetResizable(True)
        return scroll

    def create_general_tab(self) -> QWidget:
        """Create the general settings tab."""
        widget = QWidget()
        layout = QFormLayout(widget)

        # Days per year setting
        self.days_spin = QSpinBox()
        self.days_spin.setRange(1, 366)
        self.days_spin.setValue(
            self.config_manager.get_days_per_year()
        )
        layout.addRow("Days per Year:", self.days_spin)

        # Additional settings can be added here

        return widget

    def save_settings(self):
        """Save all settings to configuration."""
        try:
            # Update analyzer settings
            headers = {}
            cost_per_test = {}
            for analyzer_type, group in self.analyzer_groups.items():
                settings = group.get_settings()
                headers[analyzer_type] = AnalyzerConfig(
                    brand_cols=settings["brand_cols"],
                    workload_cols=settings["workload_cols"]
                )
                cost_per_test[analyzer_type] = settings["test_price"]

            self.config_manager.set_headers(headers)
            self.config_manager.set_cost_per_test(cost_per_test)

            # Update general settings
            self.config_manager.set_days_per_year(self.days_spin.value())

            # Save to file
            self.config_manager.save_config()

            QMessageBox.information(self, "Success", "Settings saved successfully")
            self.accept()

        except ValueError as ve:
            logger.error(f"Validation error saving settings: {ve}")
            QMessageBox.critical(self, "Validation Error", str(ve))
        except KeyError as ke:
            logger.error(f"Configuration key error: {ke}")
            QMessageBox.critical(self, "Configuration Error", f"Missing configuration key: {ke}")
        except Exception as e:
            logger.error(f"Unexpected error saving settings: {e}")
            QMessageBox.critical(self, "Error", f"Failed to save settings: {str(e)}")

    def reset_to_default(self):
        """Reset settings to default values."""
        confirm = QMessageBox.question(
            self, "Reset Confirmation",
            "Are you sure you want to reset all settings to default?",
            QMessageBox.Yes | QMessageBox.No
        )
        if confirm == QMessageBox.Yes:
            self.config_manager.load_config()
            self.close()
            self.__init__(self.config_manager)
            self.exec_()
