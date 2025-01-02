# settings_dialog.py
import logging
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QFormLayout, QSpinBox
)
from PyQt5.QtCore import Qt
from config import save_config

logger = logging.getLogger(__name__)

class SettingsDialog(QDialog):
    def __init__(self, config_data, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Settings")
        self.config_data = config_data  # dictionary from load_config()

        layout = QVBoxLayout()

        form = QFormLayout()

        # IA brand cols
        ia_brand_str = ", ".join(self.config_data["headers"]["IA"]["brand_cols"])
        ia_workload_str = ", ".join(self.config_data["headers"]["IA"]["workload_cols"])
        self.ia_brand_edit = QLineEdit(ia_brand_str)
        self.ia_workload_edit = QLineEdit(ia_workload_str)
        form.addRow(QLabel("IA brand cols (comma-separated):"), self.ia_brand_edit)
        form.addRow(QLabel("IA workload cols:"), self.ia_workload_edit)

        # CBC
        cbc_brand_str = ", ".join(self.config_data["headers"]["CBC"]["brand_cols"])
        cbc_workload_str = ", ".join(self.config_data["headers"]["CBC"]["workload_cols"])
        self.cbc_brand_edit = QLineEdit(cbc_brand_str)
        self.cbc_workload_edit = QLineEdit(cbc_workload_str)
        form.addRow(QLabel("CBC brand cols:"), self.cbc_brand_edit)
        form.addRow(QLabel("CBC workload cols:"), self.cbc_workload_edit)

        # CHEM
        chem_brand_str = ", ".join(self.config_data["headers"]["CHEM"]["brand_cols"])
        chem_workload_str = ", ".join(self.config_data["headers"]["CHEM"]["workload_cols"])
        self.chem_brand_edit = QLineEdit(chem_brand_str)
        self.chem_workload_edit = QLineEdit(chem_workload_str)
        form.addRow(QLabel("CHEM brand cols:"), self.chem_brand_edit)
        form.addRow(QLabel("CHEM workload cols:"), self.chem_workload_edit)

        # Days per year
        self.days_spin = QSpinBox()
        self.days_spin.setRange(1, 1000)
        self.days_spin.setValue(self.config_data.get("days_per_year", 330))
        form.addRow(QLabel("Days per Year:"), self.days_spin)

        layout.addLayout(form)

        hl = QHBoxLayout()
        btn_save = QPushButton("Save")
        btn_save.clicked.connect(self.save_settings)
        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)
        hl.addWidget(btn_save)
        hl.addWidget(btn_cancel)

        layout.addLayout(hl)
        self.setLayout(layout)

    def save_settings(self):
        # Convert text fields into lists
        ia_brands = [x.strip() for x in self.ia_brand_edit.text().split(",")]
        ia_workloads = [x.strip() for x in self.ia_workload_edit.text().split(",")]
        cbc_brands = [x.strip() for x in self.cbc_brand_edit.text().split(",")]
        cbc_workloads = [x.strip() for x in self.cbc_workload_edit.text().split(",")]
        chem_brands = [x.strip() for x in self.chem_brand_edit.text().split(",")]
        chem_workloads = [x.strip() for x in self.chem_workload_edit.text().split(",")]

        self.config_data["headers"]["IA"]["brand_cols"] = ia_brands
        self.config_data["headers"]["IA"]["workload_cols"] = ia_workloads

        self.config_data["headers"]["CBC"]["brand_cols"] = cbc_brands
        self.config_data["headers"]["CBC"]["workload_cols"] = cbc_workloads

        self.config_data["headers"]["CHEM"]["brand_cols"] = chem_brands
        self.config_data["headers"]["CHEM"]["workload_cols"] = chem_workloads

        self.config_data["days_per_year"] = self.days_spin.value()

        save_config(self.config_data)
        self.accept()
