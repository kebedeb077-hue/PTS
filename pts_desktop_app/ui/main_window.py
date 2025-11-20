import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget, QMessageBox
from db.access_db import AccessDB

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Form List - Publications Tracking System')
        self.setGeometry(100, 100, 800, 600)
        self.db = AccessDB()
        self.init_ui()
        self.load_data()

    def init_ui(self):
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(['Form ID', 'Form Name', 'Model No'])
        layout = QVBoxLayout()
        layout.addWidget(self.table)
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def load_data(self):
        try:
            forms = self.db.fetch_forms()
            self.table.setRowCount(len(forms))
            for row_idx, row in enumerate(forms):
                for col_idx, value in enumerate(row):
                    self.table.setItem(row_idx, col_idx, QTableWidgetItem(str(value)))
        except Exception as e:
            QMessageBox.critical(self, 'Database Error', str(e))

    def closeEvent(self, event):
        self.db.close()
        event.accept()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
