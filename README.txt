Sales Report Builder
=====================

How to run (Windows):
1. Make sure Python is installed (python.org).
2. Double-click run.bat
3. A browser tab will open with the app.

How to run (Mac/Linux):
1. Open a terminal in this folder.
2. Run: bash run.sh
3. A browser tab will open with the app.

Manual steps (any OS):
1. pip install -r requirements.txt
2. streamlit run app.py

Usage:
- Upload one or more raw export files (CSV or XLSX) using the file uploader.
- The app cleans, filters, and merges them automatically.
- Download the final report as an Excel file with two sheets:
  "UAE & Oman" and "Rest of Gulf".
