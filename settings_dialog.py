# settings_dialog.py
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTabWidget, QWidget,
    QLabel, QLineEdit, QPushButton, QFormLayout, QSpinBox,
    QDoubleSpinBox, QGroupBox, QMessageBox, QScrollArea
)
from PyQt5.QtCore import Qt
from typing import Dict, List

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
        for i in range(4):  # Maximum 4 brands
            line_edit = QLineEdit()
            if i < len(self.config["brand_columns"]):
                line_edit.setText(self.config["brand_columns"][i])
            self.brand_inputs.append(line_edit)
            layout.addRow(f"Brand {i+1}:", line_edit)

        # Test price
        self.price_spin = QDoubleSpinBox()
        self.price_spin.setRange(0, 10000)
        self.price_spin.setValue(self.config.get("test_price", 0))
        layout.addRow("Test Price:", self.price_spin)

        self.setLayout(layout)

    def get_settings(self) -> Dict:
        """Get current settings from the group."""
        return {
            "brand_columns": [edit.text() for edit in self.brand_inputs if edit.text()],
            "test_price": self.price_spin.value()
        }

class SettingsDialog(QDialog):
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
        save_button = QPushButton("Save")
        save_button.clicked.connect(self.save_settings)
        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)

        button_layout.addWidget(save_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)

    def create_analyzer_tab(self) -> QWidget:
        """Create the analyzers configuration tab."""
        scroll = QScrollArea()
        widget = QWidget()
        layout = QVBoxLayout(widget)

        for analyzer_type, config in self.config_manager.config_data["analyzers"].items():
            group = AnalyzerSettingsGroup(analyzer_type, config)
            self.analyzer_groups[analyzer_type] = group
            layout.addWidget(group)

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
            self.config_manager.config_data["analysis_settings"].get("days_per_year", 330)
        )
        layout.addRow("Days per Year:", self.days_spin)

        # Additional settings can be added here

        return widget

    def save_settings(self):
        """Save all settings to configuration."""
        try:
            # Update analyzer settings
            for analyzer_type, group in self.analyzer_groups.items():
                self.config_manager.config_data["analyzers"][analyzer_type].update(
                    group.get_settings()
                )

            # Update general settings
            self.config_manager.config_data["analysis_settings"]["days_per_year"] = self.days_spin.value()

            # Save to file
            self.config_manager.save_config()

            QMessageBox.information(self, "Success", "Settings saved successfully")
            self.accept()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save settings: {str(e)}")