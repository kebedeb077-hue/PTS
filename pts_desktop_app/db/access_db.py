import pyodbc
import os

class AccessDB:
    def __init__(self, db_path=None):
        if db_path is None:
            db_path = os.path.join(os.path.dirname(__file__), '../../Database/Forms.mdb')
        self.conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            f'DBQ={os.path.abspath(db_path)};'
            'UID=admin;PWD=serbof;'
        )
        self.conn = None

    def connect(self):
        self.conn = pyodbc.connect(self.conn_str)

    def fetch_forms(self):
        if self.conn is None:
            self.connect()
        cursor = self.conn.cursor()
        cursor.execute('SELECT Form_ID, Form_Name, Model_No FROM Forms ORDER BY Form_Name;')
        return cursor.fetchall()

    def close(self):
        if self.conn:
            self.conn.close()
            self.conn = None
