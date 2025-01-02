import logging
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QFormLayout, QSpinBox, QMessageBox
)
from PyQt5.QtCore import Qt
from config import ConfigManager, ConfigurationError

logger = logging.getLogger(__name__)

class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.config_manager = ConfigManager()
        self.config_data = self.config_manager.get_config()
        self.setup_ui()

    def setup_ui(self):
        """Initialize and setup the user interface."""
        self.setWindowTitle("Settings")
        self.setMinimumWidth(400)

        main_layout = QVBoxLayout()
        form_layout = self.create_form_layout()
        button_layout = self.create_button_layout()

        main_layout.addLayout(form_layout)
        main_layout.addLayout(button_layout)
        self.setLayout(main_layout)

    def create_form_layout(self):
        """Create and return the form layout with all input fields."""
        form = QFormLayout()

        # Create department input fields
        self.department_fields = {}
        for department in ["IA", "CBC", "CHEM"]:
            fields = self.create_department_fields(department)
            self.department_fields[department] = fields
            form.addRow(QLabel(f"{department} brand cols:"), fields['brand'])
            form.addRow(QLabel(f"{department} workload cols:"), fields['workload'])

        # Days per year spinner
        self.days_spin = QSpinBox()
        self.days_spin.setRange(1, 1000)
        self.days_spin.setValue(self.config_manager.get_days_per_year())
        form.addRow(QLabel("Days per Year:"), self.days_spin)

        return form

    def create_department_fields(self, department):
        """Create and initialize input fields for a department."""
        brand_cols = self.config_manager.get_brand_columns(department)
        workload_cols = self.config_manager.get_workload_columns(department)

        brand_edit = QLineEdit(", ".join(brand_cols))
        workload_edit = QLineEdit(", ".join(workload_cols))

        brand_edit.setToolTip("Enter comma-separated column names")
        workload_edit.setToolTip("Enter comma-separated column names")

        return {'brand': brand_edit, 'workload': workload_edit}

    def create_button_layout(self):
        """Create and return the button layout."""
        button_layout = QHBoxLayout()

        save_button = QPushButton("Save")
        save_button.clicked.connect(self.save_settings)
        save_button.setDefault(True)

        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)

        button_layout.addWidget(save_button)
        button_layout.addWidget(cancel_button)

        return button_layout

    def validate_inputs(self):
        """Validate user inputs before saving."""
        for department, fields in self.department_fields.items():
            brand_cols = [x.strip() for x in fields['brand'].text().split(",") if x.strip()]
            workload_cols = [x.strip() for x in fields['workload'].text().split(",") if x.strip()]

            if len(brand_cols) != len(workload_cols):
                raise ValueError(
                    f"{department}: Number of brand columns ({len(brand_cols)}) "
                    f"does not match workload columns ({len(workload_cols)})"
                )

            if not brand_cols or not workload_cols:
                raise ValueError(f"{department}: Brand and workload columns cannot be empty")

    def save_settings(self):
        """Save the configuration if valid."""
        try:
            self.validate_inputs()

            # Update configuration data
            for department, fields in self.department_fields.items():
                brand_cols = [x.strip() for x in fields['brand'].text().split(",") if x.strip()]
                workload_cols = [x.strip() for x in fields['workload'].text().split(",") if x.strip()]

                self.config_data["headers"][department]["brand_cols"] = brand_cols
                self.config_data["headers"][department]["workload_cols"] = workload_cols

            self.config_data["days_per_year"] = self.days_spin.value()

            # Save configuration
            self.config_manager.save_config(self.config_data)
            QMessageBox.information(self, "Success", "Settings saved successfully")
            self.accept()

        except (ValueError, ConfigurationError) as e:
            QMessageBox.critical(self, "Error", str(e))
            logger.error(f"Settings validation failed: {e}")