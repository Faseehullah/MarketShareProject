# **MarketShareProject**

A Python application for analyzing market share data using PyQt5.

---

## **Features**
- Configurable settings via GUI.
- Data aggregation and analysis.
- Market share calculation.
- Excel file processing.

---

## **Dependencies**
- **Python 3.9**
- **PyQt5**
- **pandas**
- **openpyxl**

---

## **Setup**

1. Clone the repository:
   ```bash
   git clone https://github.com/username/MarketShareProject.git
   cd MarketShareProject
1. Clone the repository
2. Install dependencies:
   ```pip install PyQt5 pandas openpyxl```
3. Run main.py

## Project Structure
- main.py: Main application entry point
- settings_dialog.py: Configuration GUI
- aggregator.py: Data processing logic
- config.py: Configuration handling
- config.json: Settings storage
- modern_ui.py
- modern_dashboard.py
- export_manager.py


## Project Structure

```plaintext
MarketShareProject/
│
├── .vscode/                    # VS Code configuration
│   ├── settings.json           # Editor settings
│   └── launch.json             # Debugger configuration
│
├── src/                        # Source code
│   ├── __init__.py             # Package initializer
│   ├── main.py                 # Main application entry
│   ├── config.py               # Configuration management
│   ├── config.json             # Setting / configurations
│   ├── aggregator.py           # Data processing logic
│   ├── visualization.py        # Visualization tools
│   ├── modern_ui.py            # Base UI components
│   ├── modern_dashboard.py     # Dashboard implementation
│   ├── export_manager.py       # Export functionality
│   └── settings_dialog.py      # Settings UI
│
├── data/                       # Data directories
│   ├── Input Jan 2025/         # Input data directory
│   ├── Output Jan 2025/        # Output data directory
│   └── LastWorkings/           # Previous versions directory
├── .gitignore                  # Git ignore file
├── README.md                   # Project documentation
|── marketsurvey.bat            # Startup script
├── tests/
│   ├── test_aggregator.py
│   └── test_config.py
└── requirements.txt