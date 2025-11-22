@echo off
REM Build script for Windows - run from repo root (double-click or run from cmd)
echo Building PTS_App...

python -m pip install --upgrade pip
pip install -r pts_desktop_app\requirements.txt
pip install pyinstaller

REM Use --add-data to include the Database/Forms.mdb file alongside the exe at runtime
REM Adjust --hidden-import flags if PyInstaller complains about missing modules
pyinstaller --noconfirm --onefile -n PTS_App --add-data "Database\Forms.mdb;Database" pts_desktop_app\ui\main_window.py

echo Build complete. The executable will be in the dist folder.
pause
