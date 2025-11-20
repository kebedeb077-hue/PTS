# Publications Tracking System - Desktop App (Python)

This is a native Windows desktop rewrite of the classic ASP Publications Tracking System using Python and PyQt5.

## Features
- View list of forms from the MS Access database (`Forms.mdb`)
- Modern desktop UI (PyQt5)

## Setup
1. Install Python 3.x
2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
3. Ensure you have the Microsoft Access ODBC driver installed (Windows).
4. Run the app:
   ```
   python ui/main_window.py
   ```

## Packaging
To create a standalone Windows executable:
```
pip install pyinstaller
pyinstaller --onefile ui/main_window.py
```

## Next Steps
- Recreate additional forms and business logic as needed.
- Enhance UI and add more features from the original ASP app.
