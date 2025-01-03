# modern_ui.py
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFrame
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QPalette, QColor

class ModernUI:
    """Base class for modern UI styling and common components."""

    @staticmethod
    def get_base_style() -> str:
        return """
            QMainWindow {
                background-color: #f5f6fa;
            }
            QLabel {
                color: #2c3e50;
                font-size: 13px;
            }
            QPushButton {
                background-color: #3498db;
                color: white;
                padding: 8px 15px;
                border: none;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QComboBox {
                padding: 5px;
                border: 1px solid #ddd;
                border-radius: 4px;
                background: white;
            }
            QLineEdit {
                padding: 5px;
                border: 1px solid #ddd;
                border-radius: 4px;
            }
            QGroupBox {
                background-color: white;
                border: 1px solid #ddd;
                border-radius: 8px;
                margin-top: 1em;
                padding: 15px;
            }
            QTabWidget::pane {
                border: 1px solid #ddd;
                border-radius: 4px;
                background: white;
            }
            QTabBar::tab {
                padding: 8px 15px;
                margin-right: 2px;
            }
        """

class ModernFrame(QFrame):
    """Modern styled frame for grouping components."""

    def __init__(self, title: str = "", parent=None):
        super().__init__(parent)
        self.title = title
        self.init_ui()

    def init_ui(self):
        self.setFrameStyle(QFrame.StyledPanel | QFrame.Raised)
        self.setStyleSheet("""
            ModernFrame {
                background-color: white;
                border: 1px solid #ddd;
                border-radius: 8px;
                padding: 15px;
            }
        """)

        layout = QVBoxLayout(self)

        if self.title:
            title_label = QLabel(self.title)
            title_label.setStyleSheet("""
                font-weight: bold;
                font-size: 14px;
                color: #2c3e50;
            """)
            layout.addWidget(title_label)

class ModernButton(QPushButton):
    """Modern styled button with hover effects."""

    def __init__(self, text: str, parent=None):
        super().__init__(text, parent)
        self.init_ui()

    def init_ui(self):
        self.setStyleSheet("""
            ModernButton {
                background-color: #3498db;
                color: white;
                padding: 8px 15px;
                border: none;
                border-radius: 4px;
                font-weight: bold;
            }
            ModernButton:hover {
                background-color: #2980b9;
            }
        """)