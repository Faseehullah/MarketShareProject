# MarketShareProject abc

A Python application for analyzing market share data using PyQt5.

## Features
- Configurable settings via GUI
- Data aggregation and analysis
- Market share calculation
- Excel file processing

## Dependencies
- Python 3.x
- PyQt5
- pandas
- openpyxl

## Setup
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

##MarketShareProject/
│
├── .vscode/                    # VS Code configuration
│   ├── settings.json
│   └── launch.json
│
├── src/                        # Source code
│   ├── __init__.py
│   ├── main.py                # Main application entry
│   ├── config.py              # Configuration management
│   ├── aggregator.py          # Data processing logic
│   ├── visualization.py       # Visualization tools
│   ├── modern_ui.py          # Base UI components
│   ├── modern_dashboard.py   # Dashboard implementation
│   ├── export_manager.py     # Export functionality
│   └── settings_dialog.py    # Settings UI
│
├── data/                       # Data directories
│   ├── Input Jan 2025/        # Input data
│   ├── Output Jan 2025/       # Output data
│   └── LastWorkings/          # Previous versions
│
├── tests/                     # Test files
│   └── __init__.py
│
├── .gitignore                # Git ignore file
├── README.md                 # Project documentation
├── config.json               # Configuration file
└── marketsurvery.bat        # Startup script