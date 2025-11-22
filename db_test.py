"""Database test script - run from project root on Windows to verify ODBC and DB access.

Usage:
    python db_test.py
"""
import os
import sys

print('Python:', sys.version)
try:
    import pyodbc
except Exception as e:
    print('pyodbc import error:', e)
    raise

print('ODBC drivers:', pyodbc.drivers())

db_path = os.path.abspath('Database/Forms.mdb')
print('DB path:', db_path)

from pts_desktop_app.db.access_db import AccessDB

print('Connection string sample:', AccessDB(db_path).conn_str)

try:
    db = AccessDB(db_path)
    rows = db.fetch_forms()
    print('Fetched', len(rows), 'rows; sample:', rows[0] if rows else 'no rows')
    db.close()
except Exception as e:
    print('DB access error:', repr(e))
    raise
