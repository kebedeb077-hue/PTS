#!/usr/bin/env python
"""Simple launcher for development: run from project root.

Usage:
    python run_app.py
"""
import sys
import os

# Ensure project root is on sys.path
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

try:
    from pts_desktop_app.ui.main_window import MainWindow
    from PyQt5.QtWidgets import QApplication
except Exception as e:
    print('Import error:', e)
    raise


def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
