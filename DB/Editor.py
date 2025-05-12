import sys
from PyQt5 import QtWidgets, QtSql
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QComboBox, QTableView, QFileDialog, QMessageBox)

class SQLiteEditor(QWidget):
    def __init__(self):
        super().__init__()
        self.db = None
        self.model = None
        self.init_ui()

    def init_ui(self):
        # Open database button and table selector
        open_btn = QPushButton("Open Database")
        open_btn.clicked.connect(self.open_database)

        self.table_combo = QComboBox()
        self.table_combo.currentIndexChanged.connect(self.load_table)

        # Add and Delete record buttons
        add_btn = QPushButton("Add Record")
        add_btn.clicked.connect(self.add_record)
        delete_btn = QPushButton("Delete Record")
        delete_btn.clicked.connect(self.delete_record)

        # Layout for controls
        controls_layout = QHBoxLayout()
        controls_layout.addWidget(open_btn)
        controls_layout.addWidget(self.table_combo)
        controls_layout.addWidget(add_btn)
        controls_layout.addWidget(delete_btn)

        # Table view
        self.table_view = QTableView()
        self.table_view.setSortingEnabled(True)

        # Main layout
        main_layout = QVBoxLayout()
        main_layout.addLayout(controls_layout)
        main_layout.addWidget(self.table_view)
        self.setLayout(main_layout)

        self.setWindowTitle("SQLite Editor")
        self.resize(800, 600)

    def open_database(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Open SQLite Database", "", "SQLite Files (*.db *.sqlite *.sqlite3);;All Files (*)"
        )
        if path:
            # Close existing connection
            if self.db and self.db.isOpen():
                self.db.close()
                QtSql.QSqlDatabase.removeDatabase('qt_sql_default_connection')

            # Connect to new database
            self.db = QtSql.QSqlDatabase.addDatabase('QSQLITE')
            self.db.setDatabaseName(path)
            if not self.db.open():
                QMessageBox.critical(self, "Error", f"Could not open database: {self.db.lastError().text()}")
                return

            self.load_tables()

    def load_tables(self):
        tables = self.db.tables()
        self.table_combo.clear()
        self.table_combo.addItems(tables)

    def load_table(self, index):
        table_name = self.table_combo.currentText()
        if not table_name:
            return

        # Clear old model
        if self.model:
            self.model.clear()

        # Set up table model
        self.model = QtSql.QSqlTableModel(self, self.db)
        self.model.setTable(table_name)
        self.model.setEditStrategy(QtSql.QSqlTableModel.OnFieldChange)
        self.model.select()

        # Show in view
        self.table_view.setModel(self.model)
        self.table_view.resizeColumnsToContents()

    def add_record(self):
        if not self.model:
            return
        row = self.model.rowCount()
        if self.model.insertRow(row):
            # Select new row
            index = self.model.index(row, 0)
            self.table_view.setCurrentIndex(index)
        else:
            QMessageBox.warning(self, "Add Record", "Could not insert new record.")

    def delete_record(self):
        if not self.model:
            return
        selection = self.table_view.selectionModel()
        if not selection.hasSelection():
            QMessageBox.warning(self, "Delete Record", "No row selected to delete.")
            return

        # Remove each selected row (reverse order)
        rows = sorted({idx.row() for idx in selection.selectedIndexes()}, reverse=True)
        for row in rows:
            self.model.removeRow(row)
        self.model.submitAll()
        self.model.select()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    editor = SQLiteEditor()
    editor.show()
    sys.exit(app.exec_())

# Usage:
# 1. Install dependencies: pip install PyQt5
# 2. Run: python sqlite_editor.py
# 3. Click "Open Database" to select any .db/.sqlite file
# 4. Select a table, then use "Add Record" and "Delete Record" buttons to modify rows.
