import sys
import os
from PyQt5 import QtWidgets, QtSql
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QComboBox,
    QTableView,
    QFileDialog,
    QMessageBox,
    QDateEdit,
    QStyledItemDelegate,
    QDialog,
    QLabel,
    QFormLayout,
    QTimeEdit,
    QLineEdit,
    QTabWidget, QSplitter, QHeaderView, QAbstractItemView, QSpacerItem, QSizePolicy, QGridLayout, QGroupBox, QRadioButton
)
from PyQt5.QtCore import Qt, QDate, QModelIndex, QTime, QSettings, QSize, QDateTime, QVariant
from PyQt5.QtGui import QColor, QPalette, QIcon
from fill_test_data import fill_test_data
from DB import create_db
import re
from docx import Document
from word import extract_placeholders, replace_placeholders
from datetime import datetime, timedelta
import pandas as pd
import random
import string

# Паттерн для функциональных маркеров: функция(аргумент)
func_pattern = re.compile(r"^([a-zA-Zа-яА-Я_]+)\(([^)]+)\)$")

class DateDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        editor = QDateEdit(parent)
        editor.setCalendarPopup(True)
        editor.setDate(QDate.currentDate())
        editor.setDisplayFormat("dd.MM.yyyy")
        editor.setKeyboardTracking(True)
        editor.setReadOnly(False)
        return editor

    def setEditorData(self, editor, index):
        value = index.model().data(index, Qt.EditRole)
        if value:
            editor.setDate(QDate.fromString(value, "yyyy-MM-dd"))

    def setModelData(self, editor, model, index):
        value = editor.date().toString("yyyy-MM-dd")
        model.setData(index, value, Qt.EditRole)

class ReadOnlyRelationalTableModel(QtSql.QSqlRelationalTableModel):
    def __init__(self, parent=None, db=None, read_only_columns_by_name=None):
        super().__init__(parent, db)
        self.read_only_columns_by_name = read_only_columns_by_name if read_only_columns_by_name else []

    def flags(self, index):
        default_flags = super().flags(index)
        if not index.isValid():
            return default_flags
        
        col_name = self.record().fieldName(index.column())
        if col_name in self.read_only_columns_by_name:
            return default_flags & ~Qt.ItemIsEditable
        return default_flags

class AddWorkDialog(QDialog):
    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("Добавление выполненной работы")
        self.setup_ui()

    def setup_ui(self):
        layout = QFormLayout(self)

        # Выбор договора
        contract_layout = QHBoxLayout()
        self.contract_combo = QComboBox()
        self.load_contracts()
        manage_services_btn = QPushButton("...")
        manage_services_btn.setToolTip("Управление услугами по договору")
        manage_services_btn.setFixedWidth(30)
        manage_services_btn.clicked.connect(self.manage_contract_services)
        contract_layout.addWidget(self.contract_combo)
        contract_layout.addWidget(manage_services_btn)
        layout.addRow("Номер договора:", contract_layout)
        self.contract_combo.currentIndexChanged.connect(self.load_services)

        # Выбор услуги
        self.service_combo = QComboBox()
        layout.addRow("Услуга:", self.service_combo)

        # Выбор вагона
        self.wagon_combo = QComboBox()
        self.load_wagons()
        layout.addRow("Номер вагона:", self.wagon_combo)

        # Выбор исполнителя
        self.worker_combo = QComboBox()
        self.load_workers()
        layout.addRow("Исполнитель:", self.worker_combo)

        # Дата выполнения
        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(QDate.currentDate())
        layout.addRow("Дата выполнения:", self.date_edit)

        # Время начала и окончания
        time_layout = QHBoxLayout()
        self.time_start = QTimeEdit()
        self.time_start.setTime(QTime(8, 0))
        self.time_end = QTimeEdit()
        self.time_end.setTime(QTime(17, 0))
        time_layout.addWidget(self.time_start)
        time_layout.addWidget(QLabel("до"))
        time_layout.addWidget(self.time_end)
        layout.addRow("Время выполнения:", time_layout)

        # Подписант
        self.signer_edit = QLineEdit()
        layout.addRow("Подписант:", self.signer_edit)

        # Кнопки
        buttons_layout = QHBoxLayout()
        save_btn = QPushButton("Сохранить")
        save_btn.clicked.connect(self.save_work)
        cancel_btn = QPushButton("Отмена")
        cancel_btn.clicked.connect(self.reject)
        buttons_layout.addWidget(save_btn)
        buttons_layout.addWidget(cancel_btn)
        layout.addRow(buttons_layout)
        
    def manage_contract_services(self):
        contract_id = self.contract_combo.currentData()
        if not contract_id:
            QMessageBox.warning(self, "Ошибка", "Выберите договор")
            return
            
        dialog = ManageContractServicesDialog(self.db, contract_id, self)
        if dialog.exec_() == QDialog.Accepted:
            # Reload services after management
            self.load_services()

    def load_contracts(self):
        query = QtSql.QSqlQuery(self.db)
        query.exec_("SELECT id, номер FROM договоры ORDER BY номер")
        self.contract_combo.clear()
        while query.next():
            self.contract_combo.addItem(query.value(1), query.value(0))

    def load_services(self):
        contract_id = self.contract_combo.currentData()
        if not contract_id:
            return
        query = QtSql.QSqlQuery(self.db)
        query.prepare("""
            SELECT у.id, у.наименование 
            FROM услуги у
            JOIN договорные_услуги ду ON у.id = ду.id_услуги
            WHERE ду.id_договора = ?
        """)
        query.addBindValue(contract_id)
        success = query.exec_()
        print(f"DEBUG: load_services query executed with success = {success}")
        if not success:
            print(f"DEBUG: SQL Error: {query.lastError().text()}")
        
        # Count results for debugging
        count = 0
        self.service_combo.clear()
        while query.next():
            count += 1
            self.service_combo.addItem(query.value(1), query.value(0))
        print(f"DEBUG: Found {count} services for contract_id = {contract_id}")
        
        # If no services found, check if we need to add them
        if count == 0:
            self.check_and_setup_contract_services(contract_id)
    
    def check_and_setup_contract_services(self, contract_id):
        # Check if the contract exists
        contract_query = QtSql.QSqlQuery(self.db)
        contract_query.prepare("SELECT номер FROM договоры WHERE id = ?")
        contract_query.addBindValue(contract_id)
        if not contract_query.exec_() or not contract_query.next():
            print(f"DEBUG: Contract with id {contract_id} not found")
            return
        
        contract_number = contract_query.value(0)
        
        # Check if any services exist at all
        services_query = QtSql.QSqlQuery(self.db)
        services_query.exec_("SELECT COUNT(*) FROM услуги")
        if services_query.next() and services_query.value(0) == 0:
            QMessageBox.warning(self, "Нет услуг", 
                               "В базе данных нет услуг. Пожалуйста, добавьте услуги в таблицу 'Услуги'.")
            return
        
        reply = QMessageBox.question(self, "Услуги не найдены", 
                                   f"Для договора {contract_number} не найдено связанных услуг.\n"
                                   f"Хотите добавить все доступные услуги для этого договора?",
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        
        if reply == QMessageBox.Yes:
            # Get all available services
            all_services_query = QtSql.QSqlQuery(self.db)
            all_services_query.exec_("SELECT id, наименование FROM услуги ORDER BY наименование")
            
            # Insert relationships in the договорные_услуги table
            self.db.transaction()
            try:
                insert_query = QtSql.QSqlQuery(self.db)
                insert_query.prepare("""
                    INSERT INTO договорные_услуги (id_договора, id_услуги)
                    VALUES (?, ?)
                """)
                
                service_count = 0
                while all_services_query.next():
                    service_id = all_services_query.value(0)
                    insert_query.bindValue(0, contract_id)
                    insert_query.bindValue(1, service_id)
                    if insert_query.exec_():
                        service_count += 1
                    else:
                        print(f"DEBUG: Failed to insert contract-service relation: {insert_query.lastError().text()}")
                
                if service_count > 0:
                    self.db.commit()
                    QMessageBox.information(self, "Успех", 
                                          f"Добавлено {service_count} услуг для договора {contract_number}")
                    # Reload services
                    self.load_services()
                else:
                    self.db.rollback()
                    QMessageBox.warning(self, "Ошибка", "Не удалось добавить услуги")
            except Exception as e:
                self.db.rollback()
                print(f"DEBUG: Exception while adding services: {e}")
                QMessageBox.critical(self, "Ошибка", f"Ошибка при добавлении услуг: {e}")

    def load_wagons(self):
        query = QtSql.QSqlQuery(self.db)
        query.exec_("SELECT id, номер FROM вагоны ORDER BY номер")
        self.wagon_combo.clear()
        while query.next():
            self.wagon_combo.addItem(query.value(1), query.value(0))

    def load_workers(self):
        query = QtSql.QSqlQuery(self.db)
        query.exec_("SELECT id, фио FROM исполнители ORDER BY фио")
        self.worker_combo.clear()
        while query.next():
            self.worker_combo.addItem(query.value(1), query.value(0))

    def save_work(self):
        # Validate required fields
        if not self.contract_combo.currentData():
            QMessageBox.warning(self, "Ошибка", "Выберите номер договора")
            return
            
        if not self.service_combo.currentData():
            QMessageBox.warning(self, "Ошибка", "Выберите услугу")
            return
            
        if not self.wagon_combo.currentData():
            QMessageBox.warning(self, "Ошибка", "Выберите номер вагона")
            return
            
        if not self.worker_combo.currentData():
            QMessageBox.warning(self, "Ошибка", "Выберите исполнителя")
            return
        
        query = QtSql.QSqlQuery(self.db)
        query.prepare("""
            INSERT INTO выполненные_работы 
            (id_вагона, id_договора, id_услуги, id_исполнителя, 
             дата_начала, дата_окончания, подписант)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """)
        
        date = self.date_edit.date().toString("yyyy-MM-dd")
        time_start = self.time_start.time().toString("HH:mm")
        time_end = self.time_end.time().toString("HH:mm")
        
        # Debug output
        print(f"DEBUG: Saving work with values:")
        print(f"  id_вагона: {self.wagon_combo.currentData()} ({self.wagon_combo.currentText()})")
        print(f"  id_договора: {self.contract_combo.currentData()} ({self.contract_combo.currentText()})")
        print(f"  id_услуги: {self.service_combo.currentData()} ({self.service_combo.currentText()})")
        print(f"  id_исполнителя: {self.worker_combo.currentData()} ({self.worker_combo.currentText()})")
        print(f"  дата_начала: {date} {time_start}")
        print(f"  дата_окончания: {date} {time_end}")
        print(f"  подписант: {self.signer_edit.text()}")
        
        query.addBindValue(self.wagon_combo.currentData())
        query.addBindValue(self.contract_combo.currentData())
        query.addBindValue(self.service_combo.currentData())
        query.addBindValue(self.worker_combo.currentData())
        query.addBindValue(f"{date} {time_start}")
        query.addBindValue(f"{date} {time_end}")
        query.addBindValue(self.signer_edit.text())

        success = query.exec_()
        if success:
            QMessageBox.information(self, "Успех", "Работа успешно добавлена")
            self.accept()
        else:
            error_text = query.lastError().text()
            print(f"DEBUG: SQL Error: {error_text}")
            QMessageBox.critical(self, "Ошибка", f"Ошибка при добавлении работы: {error_text}")

class WorkerPaymentDialog(QDialog):
    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("Расчет оплаты работника")
        self.style().unpolish(QApplication.instance())
        self.style().polish(QApplication.instance())
        self.setup_ui()
        self.setMinimumSize(800, 600)

    def setup_ui(self):
        layout = QFormLayout(self)

        # Выбор работника
        self.worker_combo = QComboBox()
        self.load_workers()
        layout.addRow("Исполнитель:", self.worker_combo)

        # Выбор периода
        period_layout = QHBoxLayout()
        self.date_start = QDateEdit()
        self.date_start.setCalendarPopup(True)
        self.date_start.setDate(QDate.currentDate().addMonths(-1))
        self.date_end = QDateEdit()
        self.date_end.setCalendarPopup(True)
        self.date_end.setDate(QDate.currentDate())
        
        period_layout.addWidget(self.date_start)
        period_layout.addWidget(QLabel("до"))
        period_layout.addWidget(self.date_end)
        layout.addRow("Период:", period_layout)

        # Кнопка расчета
        calculate_btn = QPushButton("Рассчитать")
        calculate_btn.clicked.connect(self.calculate_payment)
        layout.addRow(calculate_btn)

        # Таблица с результатами
        self.result_table = QTableView()
        self.result_table.horizontalHeader().setStretchLastSection(True)
        self.result_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.result_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        # Двойной клик для редактирования записи
        self.result_table.doubleClicked.connect(self.edit_selected_record)
        layout.addRow(self.result_table)

        # Кнопка редактирования
        edit_btn = QPushButton("Изменить запись")
        edit_btn.clicked.connect(self.edit_selected_record)
        layout.addRow(edit_btn)

        # Итоговая сумма
        self.total_label = QLabel("Итого: 0.00 руб.")
        font = self.total_label.font()
        font.setPointSize(12)
        font.setBold(True)
        self.total_label.setFont(font)
        layout.addRow(self.total_label)

    def edit_selected_record(self):
        selection = self.result_table.selectionModel()
        if not selection.hasSelection():
            QMessageBox.warning(self, "Редактирование записи", "Пожалуйста, выберите строку для редактирования.")
            return
            
        # Получаем индекс выбранной строки
        selected_rows = selection.selectedRows()
        if not selected_rows:
            rows = sorted({idx.row() for idx in selection.selectedIndexes()})
            if not rows:
                return
            row = rows[0]
        else:
            row = selected_rows[0].row()
        
        # Получаем модель таблицы
        model = self.result_table.model()
        if not model:
            return
            
        # Создаем и отображаем диалог редактирования
        dialog = EditRecordDialog(model, row, self)
        if dialog.exec_() == QDialog.Accepted:
            # После успешного редактирования обновляем отображение
            self.calculate_payment()  # Перезагружаем данные
    
    def calculate_payment(self):
        worker_id = self.worker_combo.currentData()
        if not worker_id:
            return

        start_date = self.date_start.date().toString("yyyy-MM-dd")
        end_date = self.date_end.date().toString("yyyy-MM-dd")

        # Создаем модель для отображения результатов
        # Replace QSqlQueryModel with QSqlTableModel to allow editing
        self.work_model = QtSql.QSqlTableModel(self, self.db)
        self.work_model.setTable("выполненные_работы")
        self.work_model.setFilter(f"id_исполнителя = {worker_id} AND дата_начала BETWEEN '{start_date}' AND '{end_date}'")
        self.work_model.setEditStrategy(QtSql.QSqlTableModel.OnFieldChange)
        self.work_model.select()
        
        # Set headers for the editable model
        self.work_model.setHeaderData(0, Qt.Horizontal, "ID")
        self.work_model.setHeaderData(1, Qt.Horizontal, "ID вагона")
        self.work_model.setHeaderData(2, Qt.Horizontal, "ID договора")
        self.work_model.setHeaderData(3, Qt.Horizontal, "ID услуги")
        self.work_model.setHeaderData(4, Qt.Horizontal, "ID исполнителя")
        self.work_model.setHeaderData(5, Qt.Horizontal, "Дата начала")
        self.work_model.setHeaderData(6, Qt.Horizontal, "Дата окончания")
        self.work_model.setHeaderData(7, Qt.Horizontal, "Подписант")
        
        self.result_table.setModel(self.work_model)
        self.result_table.resizeColumnsToContents()

        # Рассчитываем итоговую сумму
        sum_query = QtSql.QSqlQuery(self.db)
        sum_query.prepare("""
            SELECT SUM(у.стоимость_без_ндс)
            FROM выполненные_работы в
            JOIN услуги у ON в.id_услуги = у.id
            WHERE в.id_исполнителя = ? 
            AND в.дата_начала BETWEEN ? AND ?
        """)
        sum_query.addBindValue(worker_id)
        sum_query.addBindValue(start_date)
        sum_query.addBindValue(end_date)
        
        total = 0.0
        if sum_query.exec_() and sum_query.next():
            total = float(sum_query.value(0) or 0)
        
        self.total_label.setText(f"Итого: {total:.2f} руб.")

    def load_workers(self):
        query = QtSql.QSqlQuery(self.db)
        query.exec_("SELECT id, фио FROM исполнители ORDER BY фио")
        self.worker_combo.clear()
        while query.next():
            self.worker_combo.addItem(query.value(1), query.value(0))

class ExcelReportDialog(QDialog):
    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("Формирование Акта выполненных работ (Excel)")
        self.style().unpolish(QApplication.instance())
        self.style().polish(QApplication.instance())
        self.setup_ui()
        self.resize(600, 400)
        
    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        form_layout = QFormLayout()
        
        # Выбор договора
        self.contract_combo = QComboBox()
        self.load_contracts()
        form_layout.addRow("Номер договора:", self.contract_combo)
        
        # Выбор объема ТО
        self.to_volume_combo = QComboBox()
        self.to_volume_combo.addItems(["250", "500", "1000"])
        form_layout.addRow("Объем ТО:", self.to_volume_combo)
        
        # Порядковый номер акта
        self.act_number = QLineEdit()
        self.act_number.setText(str(random.randint(1, 999)))
        form_layout.addRow("Порядковый № акта:", self.act_number)
        
        # Период для акта
        period_layout = QHBoxLayout()
        self.date_start_edit = QDateEdit()
        self.date_start_edit.setCalendarPopup(True)
        self.date_start_edit.setDate(QDate.currentDate().addMonths(-1)) # Default to one month ago
        self.date_end_edit = QDateEdit()
        self.date_end_edit.setCalendarPopup(True)
        self.date_end_edit.setDate(QDate.currentDate()) # Default to today
        
        period_layout.addWidget(QLabel("Период с:"))
        period_layout.addWidget(self.date_start_edit)
        period_layout.addWidget(QLabel("по:"))
        period_layout.addWidget(self.date_end_edit)
        form_layout.addRow("Период акта:", period_layout)
        
        layout.addLayout(form_layout)
        
        # Add table view to preview and edit data
        preview_label = QLabel("Предварительный просмотр данных:")
        layout.addWidget(preview_label)
        
        self.preview_table = QTableView()
        self.preview_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.preview_table.horizontalHeader().setStretchLastSection(True)
        # Двойной клик для редактирования записи
        self.preview_table.doubleClicked.connect(self.edit_selected_record)
        layout.addWidget(self.preview_table)
        
        # Add buttons for data management
        buttons_layout = QHBoxLayout()
        
        load_preview_btn = QPushButton("Загрузить данные для предпросмотра")
        load_preview_btn.clicked.connect(self.load_preview_data)
        buttons_layout.addWidget(load_preview_btn)
        
        edit_btn = QPushButton("Изменить запись")
        edit_btn.clicked.connect(self.edit_selected_record)
        buttons_layout.addWidget(edit_btn)
        
        layout.addLayout(buttons_layout)
        
        # Buttons for report generation
        report_buttons_layout = QHBoxLayout()
        generate_btn = QPushButton("Сформировать отчет")
        generate_btn.clicked.connect(self.generate_report)
        cancel_btn = QPushButton("Отмена")
        cancel_btn.clicked.connect(self.reject)
        report_buttons_layout.addWidget(generate_btn)
        report_buttons_layout.addWidget(cancel_btn)
        layout.addLayout(report_buttons_layout)
    
    def edit_selected_record(self):
        selection = self.preview_table.selectionModel()
        if not selection.hasSelection():
            QMessageBox.warning(self, "Редактирование записи", "Пожалуйста, выберите строку для редактирования.")
            return
            
        # Получаем индекс выбранной строки
        selected_rows = selection.selectedRows()
        if not selected_rows:
            rows = sorted({idx.row() for idx in selection.selectedIndexes()})
            if not rows:
                return
            row = rows[0]
        else:
            row = selected_rows[0].row()
        
        # Получаем модель таблицы
        model = self.preview_table.model()
        if not model:
            return
            
        # Создаем и отображаем диалог редактирования
        dialog = EditRecordDialog(model, row, self)
        if dialog.exec_() == QDialog.Accepted:
            # После успешного редактирования обновляем отображение
            self.load_preview_data()
    
    def load_preview_data(self):
        contract_id = self.contract_combo.currentData()
        if not contract_id:
            QMessageBox.warning(self, "Ошибка", "Выберите договор")
            return
            
        # Create an editable model for preview
        self.preview_model = QtSql.QSqlRelationalTableModel(self, self.db)
        self.preview_model.setTable("выполненные_работы")
        self.preview_model.setFilter(f"id_договора = {contract_id}")
        self.preview_model.setRelation(1, QtSql.QSqlRelation("вагоны", "id", "номер"))
        self.preview_model.setRelation(2, QtSql.QSqlRelation("договоры", "id", "номер"))
        self.preview_model.setRelation(3, QtSql.QSqlRelation("услуги", "id", "наименование"))
        self.preview_model.setRelation(4, QtSql.QSqlRelation("исполнители", "id", "фио"))
        self.preview_model.setEditStrategy(QtSql.QSqlTableModel.OnFieldChange)
        self.preview_model.select()
        
        # Set headers
        self.preview_model.setHeaderData(0, Qt.Horizontal, "ID")
        self.preview_model.setHeaderData(1, Qt.Horizontal, "Вагон")
        self.preview_model.setHeaderData(2, Qt.Horizontal, "Договор")
        self.preview_model.setHeaderData(3, Qt.Horizontal, "Услуга")
        self.preview_model.setHeaderData(4, Qt.Horizontal, "Исполнитель") 
        self.preview_model.setHeaderData(5, Qt.Horizontal, "Дата начала")
        self.preview_model.setHeaderData(6, Qt.Horizontal, "Дата окончания")
        self.preview_model.setHeaderData(7, Qt.Horizontal, "Подписант")
        
        self.preview_table.setModel(self.preview_model)
        self.preview_table.hideColumn(0)  # Hide ID column
        self.preview_table.resizeColumnsToContents()

    def load_contracts(self):
        query = QtSql.QSqlQuery(self.db)
        query.exec_("SELECT id, номер FROM договоры ORDER BY номер")
        self.contract_combo.clear()
        while query.next():
            self.contract_combo.addItem(query.value(1), query.value(0))
            
    def generate_report(self):
        try:
            # Получаем выбранные данные
            contract_id = self.contract_combo.currentData()
            contract_number = self.contract_combo.currentText()
            to_volume = self.to_volume_combo.currentText()
            act_number = self.act_number.text()
            report_date_start = self.date_start_edit.date().toString("yyyy-MM-dd")
            report_date_end = self.date_end_edit.date().toString("yyyy-MM-dd")
            
            if not contract_id:
                QMessageBox.warning(self, "Ошибка", "Выберите договор")
                return
                
            # Создаем папку для отчетов, если её нет
            output_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "Отчеты")
            os.makedirs(output_dir, exist_ok=True)
            
            # Получаем необработанные данные о выполненных работах и услугах по договору
            query = QtSql.QSqlQuery(self.db)
            
            # Упрощенный запрос без агрегации и DISTINCT
            # Добавляем фильтр по дате
            sql = f"""
                SELECT 
                    у.id AS id_услуги,
                    у.наименование AS наименование_услуги,
                    у.стоимость_без_ндс AS стоимость_без_ндс,
                    у.стоимость_с_ндс AS стоимость_с_ндс,
                    в.номер AS номер_вагона
                FROM 
                    выполненные_работы вр
                JOIN услуги у ON вр.id_услуги = у.id
                JOIN вагоны в ON вр.id_вагона = в.id
                WHERE 
                    вр.id_договора = {contract_id}
                    AND вр.дата_начала BETWEEN '{report_date_start}' AND '{report_date_end}'
                ORDER BY 
                    у.наименование, в.номер;
            """
            
            print(f"Выполняется SQL запрос (без агрегации): {sql}")
            
            # --- Добавлено логирование --- 
            print("DEBUG: Перед query.exec_(sql)")
            
            success = query.exec_(sql)
            
            # --- Добавлено логирование --- 
            print(f"DEBUG: После query.exec_(sql). Успех: {success}")
            
            if not success:
                # --- Добавлено логирование ошибки --- 
                error_text = query.lastError().text()
                print(f"DEBUG: Ошибка выполнения SQL: {error_text}")
                QMessageBox.critical(self, "Ошибка", f"Ошибка при получении данных: {error_text}")
                return
                
            # --- Добавлено логирование --- 
            print("DEBUG: SQL выполнен успешно, начинаем сбор данных.")
            
            # Собираем данные в список словарей
            raw_data = []
            while query.next():
                raw_data.append({
                    "id_услуги": query.value(0),
                    "Наименование услуги": query.value(1),
                    "Стоимость за ед. без НДС": float(query.value(2) or 0),
                    "Стоимость за ед. с НДС": float(query.value(3) or 0),
                    "Номер вагона": query.value(4)
                })
            
            if not raw_data:
                QMessageBox.warning(self, "Предупреждение", "Нет данных о выполненных работах по выбранному договору")
                return
                
            # Создаем DataFrame из необработанных данных
            raw_df = pd.DataFrame(raw_data)
            
            # Выполняем агрегацию с помощью pandas
            grouped = raw_df.groupby([
                'id_услуги', 
                'Наименование услуги', 
                'Стоимость за ед. без НДС', 
                'Стоимость за ед. с НДС'
            ])
            
            # Считаем количество уникальных вагонов и собираем их номера
            aggregated_data = grouped['Номер вагона'].agg(
                Количество='nunique', 
                Номера_вагонов=lambda x: ', '.join(x.unique())
            ).reset_index()
            
            # Переименовываем колонку для соответствия отчету
            aggregated_data.rename(columns={'Номера_вагонов': 'Номера вагонов'}, inplace=True)
            
            # Вычисляем итоговые суммы для каждой услуги
            aggregated_data['Итого без НДС'] = aggregated_data['Стоимость за ед. без НДС'] * aggregated_data['Количество']
            aggregated_data['Итого с НДС'] = aggregated_data['Стоимость за ед. с НДС'] * aggregated_data['Количество']
            
            # Убираем id_услуги, так как он не нужен в финальном отчете
            df = aggregated_data.drop(columns=['id_услуги'])
            
            # Вычисляем общие итоговые суммы
            total_without_vat = df['Итого без НДС'].sum()
            total_with_vat = df['Итого с НДС'].sum()
            
            # Добавляем строку с итоговой суммой
            summary_row = pd.DataFrame([{
                "Наименование услуги": "ИТОГО:",
                "Стоимость за ед. без НДС": "",
                "Стоимость за ед. с НДС": "",
                "Количество": "",
                "Номера вагонов": "",
                "Итого без НДС": total_without_vat,
                "Итого с НДС": total_with_vat
            }])
            
            # Объединяем основной DataFrame с итоговой строкой
            df = pd.concat([df, summary_row], ignore_index=True)
            
            # Предлагаем пользователю выбрать место сохранения и имя файла
            default_filename = f"Акт_{contract_number.replace('.', '_')}_{report_date_start}_по_{report_date_end}_{act_number}.xlsx"
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Сохранить акт выполненных работ",
                os.path.join(os.path.expanduser("~"), "Documents", default_filename),
                "Excel файлы (*.xlsx);;Все файлы (*.*)"
            )
            
            if not file_path:  # Если пользователь отменил сохранение
                return
                
            # Если пользователь не указал расширение .xlsx, добавляем его
            if not file_path.lower().endswith('.xlsx'):
                file_path += '.xlsx'
            
            # Записываем в Excel
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Акт выполненных работ', index=False)
                
                # Получаем рабочий лист для форматирования
                worksheet = writer.sheets['Акт выполненных работ']
                
                # Настраиваем ширину столбцов (приблизительно)
                worksheet.column_dimensions['A'].width = 20  # Наименование услуги
                worksheet.column_dimensions['B'].width = 20  # Дата выполнения
                worksheet.column_dimensions['C'].width = 20  # Порядковый № акта
                worksheet.column_dimensions['D'].width = 20  # Номер вагона
                worksheet.column_dimensions['E'].width = 20  # Номер договора
                worksheet.column_dimensions['F'].width = 30  # Подразделение приписки
                worksheet.column_dimensions['G'].width = 40  # Услуга
                worksheet.column_dimensions['H'].width = 15  # Стоимость
                worksheet.column_dimensions['I'].width = 15  # Объем ТО
                worksheet.column_dimensions['J'].width = 30  # ФИО исполнителя
                
            QMessageBox.information(self, "Успех", f"Отчет успешно сформирован и сохранен в:\n{file_path}")
            self.accept()
            
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при формировании отчета: {str(e)}")

class ContractReportDialog(QDialog):
    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("Формирование Выписки по договорам (Excel)")
        self.style().unpolish(QApplication.instance())
        self.style().polish(QApplication.instance())
        self.setup_ui()
        self.resize(600, 400)
        
    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        # Выбор договора
        contract_layout = QFormLayout()
        self.contract_combo = QComboBox()
        self.load_contracts()
        contract_layout.addRow("Выберите договор:", self.contract_combo)
        layout.addLayout(contract_layout)
        
        # Add table view to preview and edit data
        preview_label = QLabel("Предварительный просмотр данных:")
        layout.addWidget(preview_label)
        
        self.preview_table = QTableView()
        self.preview_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.preview_table.horizontalHeader().setStretchLastSection(True)
        # Двойной клик для редактирования записи
        self.preview_table.doubleClicked.connect(self.edit_selected_record)
        layout.addWidget(self.preview_table)
        
        # Add buttons for data management
        buttons_layout = QHBoxLayout()
        
        load_preview_btn = QPushButton("Загрузить данные для предпросмотра")
        load_preview_btn.clicked.connect(self.load_preview_data)
        buttons_layout.addWidget(load_preview_btn)
        
        edit_btn = QPushButton("Изменить запись")
        edit_btn.clicked.connect(self.edit_selected_record)
        buttons_layout.addWidget(edit_btn)
        
        layout.addLayout(buttons_layout)
        
        # Кнопки
        report_buttons_layout = QHBoxLayout()
        generate_btn = QPushButton("Сформировать отчет")
        generate_btn.clicked.connect(self.generate_report)
        cancel_btn = QPushButton("Отмена")
        cancel_btn.clicked.connect(self.reject)
        report_buttons_layout.addWidget(generate_btn)
        report_buttons_layout.addWidget(cancel_btn)
        layout.addLayout(report_buttons_layout)
        
    def edit_selected_record(self):
        selection = self.preview_table.selectionModel()
        if not selection.hasSelection():
            QMessageBox.warning(self, "Редактирование записи", "Пожалуйста, выберите строку для редактирования.")
            return
            
        # Получаем индекс выбранной строки
        selected_rows = selection.selectedRows()
        if not selected_rows:
            rows = sorted({idx.row() for idx in selection.selectedIndexes()})
            if not rows:
                return
            row = rows[0]
        else:
            row = selected_rows[0].row()
        
        # Получаем модель таблицы
        model = self.preview_table.model()
        if not model:
            return
            
        # Создаем и отображаем диалог редактирования
        dialog = EditRecordDialog(model, row, self)
        if dialog.exec_() == QDialog.Accepted:
            # После успешного редактирования обновляем отображение
            self.load_preview_data()
        
    def load_preview_data(self):
        contract_id = self.contract_combo.currentData()
        if not contract_id:
            QMessageBox.warning(self, "Ошибка", "Выберите договор")
            return
            
        # Create an editable model for preview
        self.preview_model = QtSql.QSqlRelationalTableModel(self, self.db)
        self.preview_model.setTable("выполненные_работы")
        self.preview_model.setFilter(f"id_договора = {contract_id}")
        self.preview_model.setRelation(1, QtSql.QSqlRelation("вагоны", "id", "номер"))
        self.preview_model.setRelation(2, QtSql.QSqlRelation("договоры", "id", "номер"))
        self.preview_model.setRelation(3, QtSql.QSqlRelation("услуги", "id", "наименование"))
        self.preview_model.setRelation(4, QtSql.QSqlRelation("исполнители", "id", "фио"))
        self.preview_model.setEditStrategy(QtSql.QSqlTableModel.OnFieldChange)
        self.preview_model.select()
        
        # Set headers
        self.preview_model.setHeaderData(0, Qt.Horizontal, "ID")
        self.preview_model.setHeaderData(1, Qt.Horizontal, "Вагон")
        self.preview_model.setHeaderData(2, Qt.Horizontal, "Договор")
        self.preview_model.setHeaderData(3, Qt.Horizontal, "Услуга")
        self.preview_model.setHeaderData(4, Qt.Horizontal, "Исполнитель")
        self.preview_model.setHeaderData(5, Qt.Horizontal, "Дата начала")
        self.preview_model.setHeaderData(6, Qt.Horizontal, "Дата окончания")
        self.preview_model.setHeaderData(7, Qt.Horizontal, "Подписант")
        
        self.preview_table.setModel(self.preview_model)
        self.preview_table.hideColumn(0)  # Hide ID column
        self.preview_table.resizeColumnsToContents()

    def load_contracts(self):
        query = QtSql.QSqlQuery(self.db)
        query.exec_("SELECT id, номер FROM договоры ORDER BY номер")
        self.contract_combo.clear()
        while query.next():
            self.contract_combo.addItem(query.value(1), query.value(0))
            
    def generate_report(self):
        try:
            # Получаем выбранный договор
            contract_id = self.contract_combo.currentData()
            contract_number = self.contract_combo.currentText()
            
            if not contract_id:
                QMessageBox.warning(self, "Ошибка", "Выберите договор")
                return
                
            # Создаем папку для отчетов, если её нет
            output_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "Отчеты")
            os.makedirs(output_dir, exist_ok=True)
            
            # Получаем список услуг по договору и статистику по ним
            query = QtSql.QSqlQuery(self.db)
            
            # Упрощенный запрос без агрегации и DISTINCT
            sql = f"""
                SELECT 
                    у.id AS id_услуги,
                    у.наименование AS наименование_услуги,
                    у.стоимость_без_ндс AS стоимость_без_ндс,
                    у.стоимость_с_ндс AS стоимость_с_ндс,
                    в.номер AS номер_вагона
                FROM 
                    выполненные_работы вр
                JOIN услуги у ON вр.id_услуги = у.id
                JOIN вагоны в ON вр.id_вагона = в.id
                WHERE 
                    вр.id_договора = {contract_id}
                ORDER BY 
                    у.наименование, в.номер;
            """
            
            print(f"Выполняется SQL запрос (без агрегации): {sql}")
            
            # --- Добавлено логирование --- 
            print("DEBUG: Перед query.exec_(sql)")
            
            success = query.exec_(sql)
            
            # --- Добавлено логирование --- 
            print(f"DEBUG: После query.exec_(sql). Успех: {success}")
            
            if not success:
                # --- Добавлено логирование ошибки --- 
                error_text = query.lastError().text()
                print(f"DEBUG: Ошибка выполнения SQL: {error_text}")
                QMessageBox.critical(self, "Ошибка", f"Ошибка при получении данных: {error_text}")
                return
                
            # --- Добавлено логирование --- 
            print("DEBUG: SQL выполнен успешно, начинаем сбор данных.")
            
            # Собираем данные в список словарей
            raw_data = []
            while query.next():
                raw_data.append({
                    "id_услуги": query.value(0),
                    "Наименование услуги": query.value(1),
                    "Стоимость за ед. без НДС": float(query.value(2) or 0),
                    "Стоимость за ед. с НДС": float(query.value(3) or 0),
                    "Номер вагона": query.value(4)
                })
            
            if not raw_data:
                QMessageBox.warning(self, "Предупреждение", "Нет данных о выполненных работах по выбранному договору")
                return
                
            # Создаем DataFrame из необработанных данных
            raw_df = pd.DataFrame(raw_data)
            
            # Выполняем агрегацию с помощью pandas
            grouped = raw_df.groupby([
                'id_услуги', 
                'Наименование услуги', 
                'Стоимость за ед. без НДС', 
                'Стоимость за ед. с НДС'
            ])
            
            # Считаем количество уникальных вагонов и собираем их номера
            aggregated_data = grouped['Номер вагона'].agg(
                Количество='nunique', 
                Номера_вагонов=lambda x: ', '.join(x.unique())
            ).reset_index()
            
            # Переименовываем колонку для соответствия отчету
            aggregated_data.rename(columns={'Номера_вагонов': 'Номера вагонов'}, inplace=True)
            
            # Вычисляем итоговые суммы для каждой услуги
            aggregated_data['Итого без НДС'] = aggregated_data['Стоимость за ед. без НДС'] * aggregated_data['Количество']
            aggregated_data['Итого с НДС'] = aggregated_data['Стоимость за ед. с НДС'] * aggregated_data['Количество']
            
            # Убираем id_услуги, так как он не нужен в финальном отчете
            df = aggregated_data.drop(columns=['id_услуги'])
            
            # Вычисляем общие итоговые суммы
            total_without_vat = df['Итого без НДС'].sum()
            total_with_vat = df['Итого с НДС'].sum()
            
            # Добавляем строку с итоговой суммой
            summary_row = pd.DataFrame([{
                "Наименование услуги": "ИТОГО:",
                "Стоимость за ед. без НДС": "",
                "Стоимость за ед. с НДС": "",
                "Количество": "",
                "Номера вагонов": "",
                "Итого без НДС": total_without_vat,
                "Итого с НДС": total_with_vat
            }])
            
            # Объединяем основной DataFrame с итоговой строкой
            df = pd.concat([df, summary_row], ignore_index=True)
            
            # Формируем имя файла
            current_date = datetime.now().strftime("%Y-%m-%d")
            filename = f"Отчет_по_договору_{contract_number.replace('.', '_')}_{current_date}.xlsx"
            file_path = os.path.join(output_dir, filename)
            
            # Записываем в Excel
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Отчет по договору', index=False)
                
                # Получаем рабочий лист для форматирования
                worksheet = writer.sheets['Отчет по договору']
                
                # Настраиваем ширину столбцов
                worksheet.column_dimensions['A'].width = 40  # Наименование услуги
                worksheet.column_dimensions['B'].width = 20  # Стоимость без НДС
                worksheet.column_dimensions['C'].width = 20  # Стоимость с НДС
                worksheet.column_dimensions['D'].width = 15  # Количество
                worksheet.column_dimensions['E'].width = 40  # Номера вагонов
                worksheet.column_dimensions['F'].width = 20  # Итого без НДС
                worksheet.column_dimensions['G'].width = 20  # Итого с НДС
                
            QMessageBox.information(self, "Успех", f"Отчет успешно сформирован и сохранен в:\n{file_path}")
            self.accept()
            
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при формировании отчета: {str(e)}")

class ManageOwnersDialog(QDialog):
    def __init__(self, settings, parent=None):
        super().__init__(parent)
        self.settings = settings
        self.setWindowTitle("Управление списком собственников")
        self.setup_ui()
        self.load_owners()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        # Список собственников
        self.owners_list = QtWidgets.QListWidget()
        layout.addWidget(self.owners_list)
        
        # Поле ввода нового собственника
        input_layout = QHBoxLayout()
        self.new_owner_edit = QLineEdit()
        self.new_owner_edit.setPlaceholderText("Введите название собственника")
        add_btn = QPushButton("Добавить")
        add_btn.clicked.connect(self.add_owner)
        input_layout.addWidget(self.new_owner_edit)
        input_layout.addWidget(add_btn)
        layout.addLayout(input_layout)
        
        # Кнопки управления
        buttons_layout = QHBoxLayout()
        delete_btn = QPushButton("Удалить")
        delete_btn.clicked.connect(self.delete_owner)
        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(self.accept)
        buttons_layout.addWidget(delete_btn)
        buttons_layout.addWidget(close_btn)
        layout.addLayout(buttons_layout)

    def load_owners(self):
        owners = self.settings.value("owners", ["ДОСС", "ФПК", "Гранд Экспресс"], type=list)
        self.owners_list.clear()
        self.owners_list.addItems(owners)

    def add_owner(self):
        new_owner = self.new_owner_edit.text().strip()
        if new_owner:
            if self.owners_list.findItems(new_owner, Qt.MatchExactly):
                QMessageBox.warning(self, "Предупреждение", "Такой собственник уже существует")
                return
            self.owners_list.addItem(new_owner)
            self.new_owner_edit.clear()
            self.save_owners()

    def delete_owner(self):
        current_item = self.owners_list.currentItem()
        if current_item:
            reply = QMessageBox.question(
                self, "Подтверждение удаления",
                f"Удалить собственника '{current_item.text()}'?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                self.owners_list.takeItem(self.owners_list.row(current_item))
                self.save_owners()

    def save_owners(self):
        owners = [self.owners_list.item(i).text() 
                 for i in range(self.owners_list.count())]
        self.settings.setValue("owners", owners)

class ManageDivisionsDialog(QDialog):
    def __init__(self, settings, parent=None):
        super().__init__(parent)
        self.settings = settings
        self.setWindowTitle("Управление списком подразделений")
        self.setup_ui()
        self.load_divisions()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        # Список подразделений
        self.divisions_list = QtWidgets.QListWidget()
        layout.addWidget(self.divisions_list)
        
        # Поле ввода нового подразделения
        input_layout = QHBoxLayout()
        self.new_division_edit = QLineEdit()
        self.new_division_edit.setPlaceholderText("Введите название подразделения")
        add_btn = QPushButton("Добавить")
        add_btn.clicked.connect(self.add_division)
        input_layout.addWidget(self.new_division_edit)
        input_layout.addWidget(add_btn)
        layout.addLayout(input_layout)
        
        # Кнопки управления
        buttons_layout = QHBoxLayout()
        delete_btn = QPushButton("Удалить")
        delete_btn.clicked.connect(self.delete_division)
        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(self.accept)
        buttons_layout.addWidget(delete_btn)
        buttons_layout.addWidget(close_btn)
        layout.addLayout(buttons_layout)

    def load_divisions(self):
        divisions = self.settings.value("divisions", ["ЛВЧ-1", "ЛВЧ-2", "ЛВЧД-1", "ЛВЧД-2"], type=list)
        self.divisions_list.clear()
        self.divisions_list.addItems(divisions)

    def add_division(self):
        new_division = self.new_division_edit.text().strip()
        if new_division:
            if self.divisions_list.findItems(new_division, Qt.MatchExactly):
                QMessageBox.warning(self, "Предупреждение", "Такое подразделение уже существует")
                return
            self.divisions_list.addItem(new_division)
            self.new_division_edit.clear()
            self.save_divisions()

    def delete_division(self):
        current_item = self.divisions_list.currentItem()
        if current_item:
            reply = QMessageBox.question(
                self, "Подтверждение удаления",
                f"Удалить подразделение '{current_item.text()}'?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                self.divisions_list.takeItem(self.divisions_list.row(current_item))
                self.save_divisions()

    def save_divisions(self):
        divisions = [self.divisions_list.item(i).text() 
                    for i in range(self.divisions_list.count())]
        self.settings.setValue("divisions", divisions)

class ManageRepairTypesDialog(QDialog):
    def __init__(self, settings, parent=None):
        super().__init__(parent)
        self.settings = settings
        self.setWindowTitle("Управление типами ремонта")
        self.setup_ui()
        self.load_repair_types()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        form_layout = QFormLayout()
        # Поля для названий типов ремонта
        self.repair_type1 = QLineEdit()
        self.repair_type2 = QLineEdit()
        self.repair_type3 = QLineEdit()
        self.repair_type4 = QLineEdit()
        form_layout.addRow("Тип ремонта 1:", self.repair_type1)
        form_layout.addRow("Тип ремонта 2:", self.repair_type2)
        form_layout.addRow("Тип ремонта 3:", self.repair_type3)
        form_layout.addRow("Тип ремонта 4:", self.repair_type4)
        layout.addLayout(form_layout)
        # Кнопки
        buttons_layout = QHBoxLayout()
        save_btn = QPushButton("Сохранить")
        save_btn.clicked.connect(self.save_repair_types)
        cancel_btn = QPushButton("Отмена")
        cancel_btn.clicked.connect(self.reject)
        buttons_layout.addWidget(save_btn)
        buttons_layout.addWidget(cancel_btn)
        layout.addLayout(buttons_layout)

    def load_repair_types(self):
        repair_types = self.settings.value("repair_types", ["КР", "КР1", "КВР", "КР1"], type=list)
        self.repair_type1.setText(repair_types[0] if len(repair_types) > 0 else "КР")
        self.repair_type2.setText(repair_types[1] if len(repair_types) > 1 else "КР1")
        self.repair_type3.setText(repair_types[2] if len(repair_types) > 2 else "КВР")
        self.repair_type4.setText(repair_types[3] if len(repair_types) > 3 else "КР1")

    def save_repair_types(self):
        repair_types = [
            self.repair_type1.text().strip() or "КР",
            self.repair_type2.text().strip() or "КР1",
            self.repair_type3.text().strip() or "КВР",
            self.repair_type4.text().strip() or "КР1"
        ]
        self.settings.setValue("repair_types", repair_types)
        self.accept()

class AddWagonDialog(QDialog):
    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self.settings = QSettings("MyCompany", "WagonApp")
        self.setWindowTitle("Добавление вагона")
        # Store references to the label widgets for repair dates for easier updating
        self._repair_date_labels = []
        self.setup_ui()

    def setup_ui(self):
        layout = QFormLayout(self)
        self.setLayout(layout) # Ensure the main layout is set for self.layout() to work later

        # Номер вагона
        self.number_edit = QLineEdit()
        self.number_edit.setPlaceholderText("024-06064")
        layout.addRow("Номер вагона:", self.number_edit)

        # Собственник (выпадающий список + кнопка управления)
        owner_layout = QHBoxLayout()
        self.owner_combo = QComboBox()
        self.load_owners()
        manage_owners_btn = QPushButton("...")
        manage_owners_btn.setToolTip("Управление списком собственников")
        manage_owners_btn.setFixedWidth(30)
        manage_owners_btn.clicked.connect(self.manage_owners)
        owner_layout.addWidget(self.owner_combo)
        owner_layout.addWidget(manage_owners_btn)
        layout.addRow("Собственник:", owner_layout)

        # Подразделение (выпадающий список + кнопка управления)
        division_layout = QHBoxLayout()
        self.division_combo = QComboBox()
        self.load_divisions()
        manage_divisions_btn = QPushButton("...")
        manage_divisions_btn.setToolTip("Управление списком подразделений")
        manage_divisions_btn.setFixedWidth(30)
        manage_divisions_btn.clicked.connect(self.manage_divisions)
        division_layout.addWidget(self.division_combo)
        division_layout.addWidget(manage_divisions_btn)
        layout.addRow("Подразделение:", division_layout)

        # Даты ремонта - теперь напрямую в основной layout
        self.repair_dates = []
        self._repair_date_labels = [] # Reset a_labels list

        for _ in range(4):
            date_edit = QDateEdit()
            date_edit.setCalendarPopup(True)
            date_edit.setDate(QDate.currentDate())
            date_edit.setDisplayFormat("dd.MM.yyyy")
            date_edit.setKeyboardTracking(True)
            date_edit.setReadOnly(False)
            
            clear_btn = QPushButton("×")
            clear_btn.setFixedWidth(20)
            clear_btn.setToolTip("Очистить дату")
            clear_btn.clicked.connect(lambda checked, d=date_edit: self.clear_date(d))
            
            date_widget_layout = QHBoxLayout()
            date_widget_layout.addWidget(date_edit)
            date_widget_layout.addWidget(clear_btn)
            
            # We will add this to the form layout in update_repair_date_labels
            self.repair_dates.append((date_edit, date_widget_layout)) 
        
        # Добавляем поля с метками из настроек напрямую в главный layout
        self.update_repair_date_labels() # No layout argument needed, uses self.layout()
        
        # Кнопка управления типами ремонта - отдельной строкой в основном layout
        manage_types_btn = QPushButton("Настроить типы ремонта...")
        manage_types_btn.clicked.connect(self.manage_repair_types)
        layout.addRow(manage_types_btn) # Add as a full row span or just the button? Let's make it take the field column
        # To make it cleaner, let's add it as a simple row, spanning if necessary, or just on the right
        # For QFormLayout, adding a widget without a label spans it.
        # Or add it with an empty label: layout.addRow("", manage_types_btn)
        layout.addRow(manage_types_btn)


        # Кнопки Сохранить/Отмена
        buttons_layout = QHBoxLayout()
        save_btn = QPushButton("Сохранить")
        save_btn.clicked.connect(self.save_wagon)
        cancel_btn = QPushButton("Отмена")
        cancel_btn.clicked.connect(self.reject)
        buttons_layout.addWidget(save_btn)
        buttons_layout.addWidget(cancel_btn)
        layout.addRow(buttons_layout)

    def clear_date(self, date_edit):
        date_edit.setDate(QDate(2000, 1, 1)) 
        
    def is_date_empty(self, date):
        return date == QDate(2000, 1, 1)

    def update_repair_date_labels(self):
        # Remove old date rows if they exist and were added by this method
        # This is tricky if other rows are intermingled.
        # A simpler way: store the label widgets and update their text.
        
        main_layout = self.layout() # Get the main QFormLayout
        repair_types = self.settings.value("repair_types", ["КР", "КР1", "КВР", "ДР"], type=list)

        # If labels are already created, update them
        if self._repair_date_labels and len(self._repair_date_labels) == len(self.repair_dates):
            for i, (date_edit_widget, date_widget_layout) in enumerate(self.repair_dates):
                label_widget = self._repair_date_labels[i]
                repair_type = repair_types[i] if i < len(repair_types) else f"Тип {i+1}"
                label_widget.setText(f"{repair_type}:")
        else: # First time creation, add rows
            self._repair_date_labels = [] # Clear and rebuild
            for i, (date_edit_widget, date_widget_layout) in enumerate(self.repair_dates):
                repair_type = repair_types[i] if i < len(repair_types) else f"Тип {i+1}"
                # QFormLayout.addRow returns the label it creates, or None
                # We need to create a QLabel ourselves to store it.
                label = QLabel(f"{repair_type}:")
                main_layout.addRow(label, date_widget_layout)
                self._repair_date_labels.append(label)


    def manage_repair_types(self):
        dialog = ManageRepairTypesDialog(self.settings, self)
        if dialog.exec_() == QDialog.Accepted:
            # Обновляем метки полей дат в основном layout
            self.update_repair_date_labels() # This will now update existing labels or re-create if needed

    def load_owners(self):
        owners = self.settings.value("owners", ["ДОСС", "ФПК", "Гранд Экспресс"], type=list)
        self.owner_combo.clear()
        self.owner_combo.addItems(owners)
        self.owner_combo.insertItem(0, "")
        self.owner_combo.setCurrentIndex(0)

    def load_divisions(self):
        divisions = self.settings.value("divisions", ["ЛВЧ-1", "ЛВЧ-2", "ЛВЧД-1", "ЛВЧД-2"], type=list)
        self.division_combo.clear()
        self.division_combo.addItems(divisions)
        self.division_combo.insertItem(0, "")
        self.division_combo.setCurrentIndex(0)

    def manage_owners(self):
        dialog = ManageOwnersDialog(self.settings, self)
        if dialog.exec_() == QDialog.Accepted:
            current_owner = self.owner_combo.currentText()
            self.load_owners()
            index = self.owner_combo.findText(current_owner)
            if index >= 0:
                self.owner_combo.setCurrentIndex(index)

    def manage_divisions(self):
        dialog = ManageDivisionsDialog(self.settings, self)
        if dialog.exec_() == QDialog.Accepted:
            current_division = self.division_combo.currentText()
            self.load_divisions()
            index = self.division_combo.findText(current_division)
            if index >= 0:
                self.division_combo.setCurrentIndex(index)

    def save_wagon(self):
        number = self.number_edit.text().strip()
        if not number:
            QMessageBox.warning(self, "Ошибка", "Номер вагона обязателен для заполнения")
            return

        query = QtSql.QSqlQuery(self.db)
        sql_statement = """
            INSERT INTO вагоны (номер, собственник, подразделение, дата_кр, дата_кр1, дата_квр, дата_др)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """
        
        if not query.prepare(sql_statement):
            QMessageBox.critical(self, "Ошибка подготовки SQL", 
                                 f"Не удалось подготовить запрос: {query.lastError().text()}\nSQL: {sql_statement}")
            return
        
        query.addBindValue(number)
        query.addBindValue(self.owner_combo.currentText())
        query.addBindValue(self.division_combo.currentText())
        
        print(f"Перед связыванием дат, количество элементов в self.repair_dates: {len(self.repair_dates)}") # DEBUG PRINT
        if len(self.repair_dates) != 4:
            QMessageBox.critical(self, "Ошибка данных", 
                                 f"Внутренняя ошибка: ожидалось 4 элемента для дат ремонта, найдено {len(self.repair_dates)}.")
            return

        print("Сохранение дат ремонта:") # DEBUG PRINT
        current_repair_types = self.settings.value("repair_types", ["КР", "КР1", "КВР", "ДР"], type=list)
        for i, (date_edit, _) in enumerate(self.repair_dates):
            date = date_edit.date()
            value_to_bind = None if self.is_date_empty(date) else date.toString("yyyy-MM-dd")
            query.addBindValue(value_to_bind)
            repair_type_name = current_repair_types[i] if i < len(current_repair_types) else f"Тип {i+1}"
            print(f"  Дата {i+1} ({repair_type_name}): {value_to_bind}") # DEBUG PRINT

        if query.exec_():
            QMessageBox.information(self, "Успех", "Вагон успешно добавлен")
            # Регистрируем undo
            parent = self.parent()
            if parent and hasattr(parent, 'register_undo_add'):
                # Получаем id только что добавленного вагона
                last_id = None
                q = QtSql.QSqlQuery(self.db)
                q.exec_("SELECT MAX(id) FROM вагоны")
                if q.next():
                    last_id = q.value(0)
                parent.register_undo_add("вагоны", last_id)
            self.accept()
        else:
            QMessageBox.critical(self, "Ошибка при добавлении вагона", 
                                 f"Ошибка: {query.lastError().text()}\nПроверьте консоль для вывода отладочной информации по датам.")

class SQLiteEditor(QWidget):
    TABLES_RUSSIAN_NAMES = {
        "вагоны": "Вагоны",
        "договоры": "Договоры",
        "услуги": "Услуги",
        "договорные_услуги": "Услуги по договорам",
        "исполнители": "Исполнители",
        "выполненные_работы": "Выполненные работы"
    }

    def __init__(self):
        super().__init__()
        self.db = None
        self.model = None
        self.settings = QSettings("MyCompany", "WagonApp")
        # Track last operation for undo
        self.last_operation = None
        self.last_operation_data = None
        self.init_ui()
        self.apply_styles()
        self.load_last_database()

    def apply_styles(self):
        app = QApplication.instance()
        app.setStyle("Fusion")

        primary = QColor(183, 28, 28)
        secondary = QColor(38, 50, 56)

        dark_palette = QPalette()
        dark_palette.setColor(QPalette.Window, QColor(45, 45, 45))
        dark_palette.setColor(QPalette.WindowText, Qt.white)
        dark_palette.setColor(QPalette.Base, QColor(30, 30, 30))
        dark_palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
        dark_palette.setColor(QPalette.ToolTipBase, Qt.white)
        dark_palette.setColor(QPalette.ToolTipText, Qt.black)
        dark_palette.setColor(QPalette.Text, Qt.white)
        dark_palette.setColor(QPalette.Button, QColor(60, 60, 60))
        dark_palette.setColor(QPalette.ButtonText, Qt.white)
        dark_palette.setColor(QPalette.BrightText, Qt.red)
        dark_palette.setColor(QPalette.Link, primary)
        dark_palette.setColor(QPalette.Highlight, primary)
        dark_palette.setColor(QPalette.HighlightedText, Qt.white)
        dark_palette.setColor(QPalette.Disabled, QPalette.Text, QColor(127, 127, 127))
        dark_palette.setColor(QPalette.Disabled, QPalette.ButtonText, QColor(127, 127, 127))

        app.setPalette(dark_palette)

        app.setStyleSheet("""
            QWidget {
                font-size: 10pt;
            }
            QPushButton {
                background-color: #424242;
                color: #E0E0E0;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                min-width: 100px;
                border-bottom: 3px solid #252525;
                border-right: 2px solid #252525;
            }
            QPushButton#OpenDbButton {
                 background-color: #B71C1C; /* Красный цвет */
            }
             QPushButton#OpenDbButton:hover {
                 background-color: #C62828; /* Более светлый красный для наведения */
            }
             QPushButton#OpenDbButton:pressed {
                 background-color: #A31515; /* Более темный красный для нажатия */
            }
            QPushButton#CreateDbButton {
                 background-color: #1976D2; /* Синий цвет */
            }
            QPushButton#CreateDbButton:hover {
                 background-color: #2196F3; /* Более светлый синий для наведения */
            }
            QPushButton#CreateDbButton:pressed {
                 background-color: #1565C0; /* Более темный синий для нажатия */
            }
            QPushButton#DeleteDbButton {
                 background-color: #484848;
            }
            QPushButton#DeleteDbButton:hover {
                 background-color: #585858;
            }
            QPushButton#DeleteDbButton:pressed {
                 background-color: #303030;
            }
            QPushButton:hover {
                background-color: #565656;
            }
            QPushButton:pressed {
                background-color: #363636;
                border-bottom: 1px solid #252525;
                border-right: 1px solid #252525;
                margin-top: 2px;
                margin-left: 2px;
            }
            QPushButton:disabled {
                background-color: #3A3A3A;
                color: #707070;
            }
            QTableView {
                gridline-color: #616161;
                selection-background-color: #B71C1C;
                alternate-background-color: #3A3A3A;
                background-color: #2E2E2E;
                color: white;
            }
            QTableView QHeaderView::section {
                background-color: #3A3A3A;
                color: white;
                padding: 4px;
                border: 1px solid #616161;
                font-weight: bold;
            }
            QComboBox {
                padding: 5px;
                border: 1px solid #616161;
                border-radius: 3px;
                min-width: 150px;
            }
            QComboBox::drop-down {
                 border: none;
             }
             QComboBox QAbstractItemView {
                 background-color: #3A3A3A;
                 color: white;
                 selection-background-color: #B71C1C;
             }
            QLabel {
                padding: 2px;
            }
            QTabWidget::pane {
                border: 1px solid #616161;
                background-color: #2E2E2E;
            }
            QTabBar::tab {
                background: #424242;
                color: white;
                padding: 8px 20px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                border: 1px solid #616161;
                border-bottom: none;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background: #2E2E2E;
                color: #B71C1C;
                font-weight: bold;
            }
            QTabBar::tab:!selected {
                margin-top: 2px;
            }
            QSplitter::handle {
                background-color: #616161;
                height: 3px;
            }
        """)
        self.style().unpolish(self)
        self.style().polish(self)

    def init_ui(self):
        self.setWindowTitle("Система управления ремонта вагонного оборудования")
        self.resize(1200, 800)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 10, 10, 10)

        # Реорганизуем кнопки в два ряда
        db_buttons_layout = QGridLayout()
        
        # Первый ряд кнопок базы данных
        open_btn = QPushButton("Открыть БД")
        open_btn.setObjectName("OpenDbButton")
        open_btn.setToolTip("Открыть существующий файл базы данных SQLite")
        open_btn.clicked.connect(self.open_database)
        db_buttons_layout.addWidget(open_btn, 0, 0)
        
        create_btn = QPushButton("Создать БД")
        create_btn.setObjectName("CreateDbButton")
        create_btn.setToolTip("Создать новый файл базы данных SQLite")
        create_btn.clicked.connect(self.create_database)
        db_buttons_layout.addWidget(create_btn, 0, 1)
        
        # Второй ряд кнопок базы данных
        fill_btn = QPushButton("Заполнить тестовыми данными")
        fill_btn.setToolTip("Заполнить текущую базу данных тестовыми записями (перезапишет существующие!)")
        fill_btn.clicked.connect(self.fill_test_data)
        db_buttons_layout.addWidget(fill_btn, 1, 0)

        delete_db_btn = QPushButton("Удалить БД")
        delete_db_btn.setObjectName("DeleteDbButton")
        delete_db_btn.setToolTip("Удалить текущий файл базы данных (НЕОБРАТИМО!)")
        delete_db_btn.clicked.connect(self.delete_database)
        db_buttons_layout.addWidget(delete_db_btn, 1, 1)

        # Добавляем кнопки DB в основной layout
        main_layout.addLayout(db_buttons_layout)
        
        # Добавляем промежуток между группами кнопок
        spacer = QSpacerItem(20, 20, QSizePolicy.Minimum, QSizePolicy.Fixed)
        main_layout.addSpacerItem(spacer)
        
        # Создаем сетку для кнопок отчетов
        report_buttons_layout = QGridLayout()
        
        # Первый ряд кнопок отчетов
        self.word_report_btn = QPushButton("Заполнить шаблон Word")
        self.word_report_btn.setToolTip("Заполнить шаблон документа Word данными из БД")
        self.word_report_btn.clicked.connect(self.show_fill_word_dialog)
        self.word_report_btn.setEnabled(False)
        report_buttons_layout.addWidget(self.word_report_btn, 0, 0)

        self.excel_report_btn = QPushButton("Акт работ (Excel)")
        self.excel_report_btn.setToolTip("Сформировать Акт выполненных работ в формате Excel")
        self.excel_report_btn.clicked.connect(self.show_excel_report_dialog)
        self.excel_report_btn.setEnabled(False)
        report_buttons_layout.addWidget(self.excel_report_btn, 0, 1)
        
        # Второй ряд кнопок отчетов
        self.contract_report_btn = QPushButton("Выписка по договору (Excel)")
        self.contract_report_btn.setToolTip("Сформировать выписку по выбранному договору в формате Excel")
        self.contract_report_btn.clicked.connect(self.show_contract_report_dialog)
        self.contract_report_btn.setEnabled(False)
        report_buttons_layout.addWidget(self.contract_report_btn, 1, 0)

        self.worker_payment_btn = QPushButton("Расчет оплаты работника")
        self.worker_payment_btn.setToolTip("Рассчитать сдельную оплату для выбранного работника за период")
        self.worker_payment_btn.clicked.connect(self.show_worker_payment_dialog)
        self.worker_payment_btn.setEnabled(False)
        report_buttons_layout.addWidget(self.worker_payment_btn, 1, 1)

        # Добавляем кнопки отчетов в основной layout
        main_layout.addLayout(report_buttons_layout)

        self.tab_widget = QTabWidget()
        main_layout.addWidget(self.tab_widget)

        self.data_tab = QWidget()
        self.tab_widget.addTab(self.data_tab, "Данные")
        data_tab_layout = QVBoxLayout(self.data_tab)
        data_tab_layout.setContentsMargins(5, 10, 5, 5)

        splitter = QSplitter(Qt.Horizontal)

        left_panel_widget = QWidget()
        left_panel_layout = QVBoxLayout(left_panel_widget)
        left_panel_layout.setContentsMargins(0, 0, 5, 0)

        table_select_label = QLabel("Выберите таблицу:")
        left_panel_layout.addWidget(table_select_label)

        self.table_combo = QComboBox()
        self.table_combo.setToolTip("Выберите таблицу для просмотра и редактирования")
        self.table_combo.currentIndexChanged.connect(self.load_table)
        left_panel_layout.addWidget(self.table_combo)

        left_panel_layout.addSpacerItem(QSpacerItem(20, 20, QSizePolicy.Minimum, QSizePolicy.Fixed))

        record_control_label = QLabel("Управление записями:")
        left_panel_layout.addWidget(record_control_label)

        self.undo_btn = QPushButton("Отменить последнее действие")
        self.undo_btn.setToolTip("Отменить последнее добавление/удаление/изменение записи")
        self.undo_btn.clicked.connect(self.undo_last_operation)
        self.undo_btn.setEnabled(False)
        left_panel_layout.addWidget(self.undo_btn)

        self.add_record_btn = QPushButton("Добавить запись")
        self.add_record_btn.setToolTip("Добавить новую пустую строку в выбранную таблицу")
        self.add_record_btn.clicked.connect(self.add_record)
        self.add_record_btn.setEnabled(False)
        left_panel_layout.addWidget(self.add_record_btn)
        
        # Добавляем кнопку изменения записи
        self.edit_record_btn = QPushButton("Изменить запись")
        self.edit_record_btn.setToolTip("Открыть диалог редактирования выбранной записи")
        self.edit_record_btn.clicked.connect(self.edit_record)
        self.edit_record_btn.setEnabled(False)
        left_panel_layout.addWidget(self.edit_record_btn)
        
        self.add_work_btn = QPushButton("Добавить вып. работу")
        self.add_work_btn.setToolTip("Открыть диалог для добавления новой выполненной работы")
        self.add_work_btn.clicked.connect(self.show_add_work_dialog)
        self.add_work_btn.setEnabled(False)
        self.add_work_btn.setVisible(False)
        left_panel_layout.addWidget(self.add_work_btn)

        # New button for managing contract services
        self.manage_contract_services_btn = QPushButton("Управлять услугами договора")
        self.manage_contract_services_btn.setToolTip("Открыть диалог управления услугами для выбранного типа таблицы")
        self.manage_contract_services_btn.clicked.connect(self.show_manage_contract_services_dialog)
        self.manage_contract_services_btn.setEnabled(False)
        self.manage_contract_services_btn.setVisible(False)
        left_panel_layout.addWidget(self.manage_contract_services_btn)

        self.delete_record_btn = QPushButton("Удалить запись(и)")
        self.delete_record_btn.setToolTip("Удалить выбранные строки из таблицы")
        self.delete_record_btn.clicked.connect(self.delete_record)
        self.delete_record_btn.setEnabled(False)
        left_panel_layout.addWidget(self.delete_record_btn)

        self.import_excel_btn = QPushButton("Загрузить из Excel")
        self.import_excel_btn.setToolTip("Импортировать данные из файла Excel в текущую таблицу")
        self.import_excel_btn.clicked.connect(self.import_from_excel)
        self.import_excel_btn.setEnabled(False)
        left_panel_layout.addWidget(self.import_excel_btn)

        left_panel_layout.addStretch(1)
        splitter.addWidget(left_panel_widget)

        right_panel_widget = QWidget()
        right_panel_layout = QVBoxLayout(right_panel_widget)
        right_panel_layout.setContentsMargins(5, 0, 0, 0)

        self.table_view = QTableView()
        self.table_view.setSortingEnabled(True)
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table_view.verticalHeader().setVisible(False)
        self.table_view.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table_view.setAlternatingRowColors(True)
        # Двойной клик для редактирования записи
        self.table_view.doubleClicked.connect(self.edit_record)
        right_panel_layout.addWidget(self.table_view)
        splitter.addWidget(right_panel_widget)

        splitter.setSizes([250, 950])
        splitter.setStretchFactor(1, 3)

        data_tab_layout.addWidget(splitter)

    def load_last_database(self):
        last_db_path = self.settings.value("database/lastOpened", "")
        if last_db_path and os.path.exists(last_db_path):
            print(f"Загрузка последней использованной БД: {last_db_path}")
            self.open_database_file(last_db_path)
        else:
            print("Последняя БД не найдена или путь некорректен.")

    def create_database(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Создать базу данных SQLite", "", "SQLite Files (*.db)"
        )
        if path:
            if not path.endswith(".db"):
                path += ".db"
            if self.db and self.db.isOpen():
                table_name = self.table_combo.currentText()
                if self.model:
                    self.model.clear()
                self.db.close()
                print(f"Закрыта база данных: {self.db.databaseName()}")
                QtSql.QSqlDatabase.removeDatabase('qt_sql_default_connection')
                self.db = None
                self.model = None
                self.table_combo.clear()
                self.table_view.setModel(None)
                self.update_button_states(db_open=False)

            try:
                create_db(path)
                QMessageBox.information(self, "Успех", f"База данных успешно создана: {path}")
                self.open_database_file(path)
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось создать базу данных: {e}")
                self.update_button_states(db_open=False)

    def open_database(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Открыть базу данных SQLite", "", "SQLite Files (*.db *.sqlite *.sqlite3);;All Files (*)"
        )
        if path:
            self.open_database_file(path)

    def open_database_file(self, path):
        if self.db and self.db.isOpen():
            table_name = self.table_combo.currentText()
            if self.model:
                self.model.clear()
            self.db.close()
            print(f"Закрыта база данных: {self.db.databaseName()}")
            QtSql.QSqlDatabase.removeDatabase('qt_sql_default_connection')
            self.db = None
            self.model = None
            self.table_combo.clear()
            self.table_view.setModel(None)

        self.db = QtSql.QSqlDatabase.addDatabase('QSQLITE', 'qt_sql_default_connection')
        self.db.setDatabaseName(path)
        if not self.db.open():
            QMessageBox.critical(self, "Ошибка", f"Не удалось открыть базу данных: {self.db.lastError().text()}")
            self.db = None
            self.update_button_states(db_open=False)
            return

        print(f"Открыта база данных: {path}")
        self.settings.setValue("database/lastOpened", path)
        self.load_tables()
        self.update_button_states(db_open=True)

    def update_button_states(self, db_open):
        is_table_selected = bool(self.table_combo.currentText()) and db_open
        
        self.word_report_btn.setEnabled(db_open)
        self.excel_report_btn.setEnabled(db_open)
        self.contract_report_btn.setEnabled(db_open)
        self.worker_payment_btn.setEnabled(db_open)
        self.table_combo.setEnabled(db_open)
        self.add_record_btn.setEnabled(is_table_selected)
        self.edit_record_btn.setEnabled(is_table_selected)
        self.delete_record_btn.setEnabled(is_table_selected)
        self.import_excel_btn.setEnabled(is_table_selected)
        # Update undo button state based on last operation
        self.undo_btn.setEnabled(bool(self.last_operation and self.last_operation_data))
        
        current_table = self.table_combo.currentText()
        current_table_data_name = self.table_combo.itemData(self.table_combo.currentIndex()) # This is the actual table name

        self.add_work_btn.setEnabled(is_table_selected and current_table_data_name == "выполненные_работы")
        self.add_work_btn.setVisible(is_table_selected and current_table_data_name == "выполненные_работы")
        
        # Show/hide manage_contract_services_btn based on table
        is_contract_services_table = is_table_selected and current_table_data_name == "договорные_услуги"
        self.manage_contract_services_btn.setEnabled(is_contract_services_table)
        self.manage_contract_services_btn.setVisible(is_contract_services_table)
        
        if not db_open:
            self.table_view.setModel(None)

    def load_tables(self):
        if not self.db or not self.db.isOpen():
            return
        tables = self.db.tables()
        self.table_combo.clear()
        for table_name in tables:
            display_name = self.TABLES_RUSSIAN_NAMES.get(table_name, table_name)
            self.table_combo.addItem(display_name, table_name)
        
        if tables:
             self.table_combo.setCurrentIndex(0)
             self.load_table(0)
        self.update_button_states(db_open=True)

    def load_table(self, index):
        table_name = self.table_combo.itemData(index) 
        if not table_name or not self.db or not self.db.isOpen():
            self.table_view.setModel(None)
            self.update_button_states(db_open=bool(self.db and self.db.isOpen()))
            return

        print(f"Загрузка таблицы: {table_name}")

        if self.model:
            self.model.clear()
            try:
                self.model.dataChanged.disconnect()
                self.model.rowsInserted.disconnect()
                self.model.rowsRemoved.disconnect()
            except TypeError:
                pass 
        
        # Для таблицы "выполненные_работы" используем QSqlRelationalTableModel для связи с другими таблицами
        if table_name == "выполненные_работы":
            # Используем ReadOnlyRelationalTableModel для выполненных_работ
            self.model = ReadOnlyRelationalTableModel(self, self.db, 
                                                  read_only_columns_by_name=["id_вагона", "id_договора", "id_услуги", "id_исполнителя"])
            self.model.setTable(table_name)
            
            # Устанавливаем связи с другими таблицами
            # id_вагона -> вагоны.номер
            self.model.setRelation(self.model.fieldIndex("id_вагона"), QtSql.QSqlRelation("вагоны", "id", "номер"))
            # id_договора -> договоры.номер
            self.model.setRelation(self.model.fieldIndex("id_договора"), QtSql.QSqlRelation("договоры", "id", "номер"))
            # id_услуги -> услуги.наименование
            self.model.setRelation(self.model.fieldIndex("id_услуги"), QtSql.QSqlRelation("услуги", "id", "наименование"))
            # id_исполнителя -> исполнители.фио
            self.model.setRelation(self.model.fieldIndex("id_исполнителя"), QtSql.QSqlRelation("исполнители", "id", "фио"))
            
            # Устанавливаем заголовки столбцов
            self.model.setHeaderData(self.model.fieldIndex("id"), Qt.Horizontal, "ID")
            self.model.setHeaderData(self.model.fieldIndex("id_вагона"), Qt.Horizontal, "Вагон")
            self.model.setHeaderData(self.model.fieldIndex("id_договора"), Qt.Horizontal, "Договор")
            self.model.setHeaderData(self.model.fieldIndex("id_услуги"), Qt.Horizontal, "Услуга")
            self.model.setHeaderData(self.model.fieldIndex("id_исполнителя"), Qt.Horizontal, "Исполнитель")
            self.model.setHeaderData(self.model.fieldIndex("дата_начала"), Qt.Horizontal, "Дата начала")
            self.model.setHeaderData(self.model.fieldIndex("дата_окончания"), Qt.Horizontal, "Дата окончания")
            self.model.setHeaderData(self.model.fieldIndex("подписант"), Qt.Horizontal, "Подписант")
        elif table_name == "договорные_услуги":
            self.model = QtSql.QSqlRelationalTableModel(self, self.db)
            self.model.setTable(table_name)
            
            # Связь для id_договора (поле с индексом 1 в таблице договорные_услуги) с таблицей договоры
            self.model.setRelation(self.model.fieldIndex("id_договора"), QtSql.QSqlRelation("договоры", "id", "номер"))
            # Связь для id_услуги (поле с индексом 2) с таблицей услуги
            self.model.setRelation(self.model.fieldIndex("id_услуги"), QtSql.QSqlRelation("услуги", "id", "наименование"))
            
            # Устанавливаем заголовки столбцов
            self.model.setHeaderData(self.model.fieldIndex("id"), Qt.Horizontal, "ID")
            self.model.setHeaderData(self.model.fieldIndex("id_договора"), Qt.Horizontal, "Договор") # Будет отображать номер договора
            self.model.setHeaderData(self.model.fieldIndex("id_услуги"), Qt.Horizontal, "Услуга") # Будет отображать наименование услуги
        else:
            # Для остальных таблиц используем стандартную модель
            self.model = QtSql.QSqlTableModel(self, self.db)
            self.model.setTable(table_name)
        
        self.model.setEditStrategy(QtSql.QSqlTableModel.OnFieldChange)
        
        # Устанавливаем делегат для полей дат, если это таблица вагонов
        if table_name == "вагоны":
            date_delegate = DateDelegate(self.table_view)
            # Находим индексы колонок по их именам
            header = self.model.headerData
            num_cols = self.model.columnCount()
            for col in range(num_cols):
                col_name = header(col, Qt.Horizontal, Qt.DisplayRole)
                if col_name in ["дата_кр", "дата_кр1", "дата_квр", "дата_др"]:
                    self.table_view.setItemDelegateForColumn(col, date_delegate)
                    print(f"Установлен DateDelegate для колонки: {col_name} (индекс {col})")
                # Сбрасываем делегат для других колонок, если он был установлен ранее
                elif self.table_view.itemDelegateForColumn(col) == date_delegate:
                     self.table_view.setItemDelegateForColumn(col, QStyledItemDelegate(self.table_view))

        # Ensure editing is enabled in the view
        self.table_view.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed)
        
        self.model.select()

        if self.model.lastError().isValid():
            QMessageBox.critical(self, "Ошибка загрузки таблицы", 
                                 f"Не удалось загрузить таблицу '{table_name}': {self.model.lastError().text()}")
            self.table_view.setModel(None)
        else:
            self.table_view.setModel(self.model)
            
            # Для таблицы "выполненные_работы" используем QSqlRelationalDelegate для правильного отображения связанных таблиц
            if table_name == "выполненные_работы":
                self.table_view.setItemDelegate(QtSql.QSqlRelationalDelegate(self.table_view))
            # Generalize delegate setting for any QSqlRelationalTableModel
            if isinstance(self.model, QtSql.QSqlRelationalTableModel):
                self.table_view.setItemDelegate(QtSql.QSqlRelationalDelegate(self.table_view))
            
        # После установки модели и данных, применяем делегаты (если select() сбрасывает их)
        if table_name == "вагоны":
            date_delegate = DateDelegate(self.table_view)
            header = self.model.headerData
            num_cols = self.model.columnCount()
            for col in range(num_cols):
                col_name = header(col, Qt.Horizontal, Qt.DisplayRole)
                if col_name in ["дата_кр", "дата_кр1", "дата_квр", "дата_др"]:
                    self.table_view.setItemDelegateForColumn(col, date_delegate)
                # Эта часть для сброса делегата может быть избыточной, если модель пересоздается
                # elif self.table_view.itemDelegateForColumn(col) == date_delegate:
                #      self.table_view.setItemDelegateForColumn(col, QStyledItemDelegate(self.table_view))

        self.table_view.resizeColumnsToContents()
        
        self.update_button_states(db_open=True)

    def show_add_work_dialog(self):
        if not self.db or not self.db.isOpen():
            QMessageBox.warning(self, "Нет базы данных", "Пожалуйста, сначала откройте или создайте базу данных.")
            return
        dialog = AddWorkDialog(self.db, self)
        if dialog.exec_() == QDialog.Accepted:
            current_table_name = self.table_combo.itemData(self.table_combo.currentIndex())
            if current_table_name == "выполненные_работы" and self.model:
                self.model.select()
    
    def show_manage_contract_services_dialog(self):
        if not self.db or not self.db.isOpen():
            QMessageBox.warning(self, "Нет базы данных", "Пожалуйста, сначала откройте или создайте базу данных.")
            return
        
        # contract_id is set to None, so the dialog will use its internal QComboBox for contract selection.
        dialog = ManageContractServicesDialog(self.db, contract_id=None, parent=self)
        if dialog.exec_() == QDialog.Accepted:
            # If changes were made, refresh the current table if it's 'договорные_услуги'
            current_table_data_name = self.table_combo.itemData(self.table_combo.currentIndex())
            if self.model and current_table_data_name == "договорные_услуги":
                self.model.select() 

    def add_record(self):
        if not self.model:
            QMessageBox.warning(self, "Нет таблицы", "Пожалуйста, сначала выберите таблицу.")
            return

        current_table_name = self.table_combo.itemData(self.table_combo.currentIndex())
        print(f"DEBUG: Adding record to table: {current_table_name}")  # Debug log

        # Используем специальные диалоги для определенных таблиц
        if current_table_name == "выполненные_работы":
            self.show_add_work_dialog()
            return
        elif current_table_name == "вагоны":
            dialog = AddWagonDialog(self.db, self)
            if dialog.exec_() == QDialog.Accepted:
                self.model.select()
            return
        elif current_table_name == "договоры":
            dialog = AddContractDialog(self.db, self)
            if dialog.exec_() == QDialog.Accepted:
                self.model.select()
            return
        elif current_table_name == "услуги":
            # Для услуг показываем диалог для ввода данных до вставки
            dialog = AddServiceDialog(self.db, self)
            if dialog.exec_() == QDialog.Accepted:
                self.model.select()
            return
        elif current_table_name == "договорные_услуги":
            # Для договорных услуг требуется выбор существующего договора и услуги
            contract_id = None
            
            # Выбор договора
            contract_query = QtSql.QSqlQuery(self.db)
            contract_query.exec_("SELECT id, номер FROM договоры ORDER BY номер")
            contracts = []
            while contract_query.next():
                contracts.append((contract_query.value(0), contract_query.value(1)))
                
            if not contracts:
                QMessageBox.warning(self, "Ошибка", "Нет доступных договоров. Сначала добавьте договор.")
                return
                
            contract_items = [f"{c[1]} (ID: {c[0]})" for c in contracts]
            contract_item, ok = QtWidgets.QInputDialog.getItem(
                self, "Выбор договора", "Выберите договор:", contract_items, 0, False)
            
            if not ok or not contract_item:
                return
                
            contract_id = contracts[contract_items.index(contract_item)][0]
            
            # Выбор услуги
            service_query = QtSql.QSqlQuery(self.db)
            service_query.exec_("SELECT id, наименование FROM услуги ORDER BY наименование")
            services = []
            while service_query.next():
                services.append((service_query.value(0), service_query.value(1)))
                
            if not services:
                QMessageBox.warning(self, "Ошибка", "Нет доступных услуг. Сначала добавьте услугу.")
                return
                
            service_items = [f"{s[1]} (ID: {s[0]})" for s in services]
            service_item, ok = QtWidgets.QInputDialog.getItem(
                self, "Выбор услуги", "Выберите услугу:", service_items, 0, False)
            
            if not ok or not service_item:
                return
                
            service_id = services[service_items.index(service_item)][0]
            
            # Проверка на дубликат
            check_query = QtSql.QSqlQuery(self.db)
            check_query.prepare(
                "SELECT COUNT(*) FROM договорные_услуги WHERE id_договора = ? AND id_услуги = ?")
            check_query.addBindValue(contract_id)
            check_query.addBindValue(service_id)
            check_query.exec_()
            if check_query.next() and check_query.value(0) > 0:
                QMessageBox.warning(self, "Дубликат", 
                                    "Эта услуга уже добавлена к этому договору.")
                return
            
            # Добавление записи
            insert_query = QtSql.QSqlQuery(self.db)
            insert_query.prepare(
                "INSERT INTO договорные_услуги (id_договора, id_услуги) VALUES (?, ?)")
            insert_query.addBindValue(contract_id)
            insert_query.addBindValue(service_id)
            
            if insert_query.exec_():
                self.model.select()
                QMessageBox.information(self, "Успех", "Услуга успешно добавлена к договору")
            else:
                QMessageBox.critical(self, "Ошибка", 
                                     f"Не удалось добавить запись: {insert_query.lastError().text()}")
            return

        # Для остальных таблиц (без жестких ограничений) используем стандартную вставку
        row = self.model.rowCount()
        print(f"DEBUG: Attempting to insert row at index: {row}")  # Debug log
        
        if self.model.insertRow(row):
            # Store operation for undo
            self.last_operation = "add"
            self.last_operation_data = {"row": row, "table": current_table_name}
            print(f"DEBUG: Stored add operation data: {self.last_operation_data}")  # Debug log
            self.undo_btn.setEnabled(True)
            
            # Select new row
            index = self.model.index(row, 0)
            self.table_view.setCurrentIndex(index)
            self.table_view.scrollTo(index)
            
            # Submit changes immediately
            if not self.model.submitAll():
                print(f"DEBUG: Error submitting new row: {self.model.lastError().text()}")  # Debug log
                QMessageBox.critical(self, "Ошибка добавления", 
                                   f"Не удалось сохранить новую запись: {self.model.lastError().text()}")
                self.model.revertAll()
                self.last_operation = None
                self.last_operation_data = None
                self.undo_btn.setEnabled(False)
                return
                
            print("DEBUG: Successfully added and submitted new row")  # Debug log
        else:
            print(f"DEBUG: Failed to insert row: {self.model.lastError().text()}")  # Debug log
            QMessageBox.critical(self, "Ошибка добавления",
                               f"Не удалось добавить запись: {self.model.lastError().text()}")

    def delete_record(self):
        if not self.model:
            QMessageBox.warning(self, "Нет таблицы", "Пожалуйста, сначала выберите таблицу.")
            return
            
        selection = self.table_view.selectionModel()
        if not selection.hasSelection():
            QMessageBox.warning(self, "Удаление записи", "Пожалуйста, выберите строки для удаления.")
            return

        selected_rows = selection.selectedRows()
        if not selected_rows:
            rows = sorted({idx.row() for idx in selection.selectedIndexes()}, reverse=True)
        else:
            rows = sorted({idx.row() for idx in selected_rows}, reverse=True)

        print(f"DEBUG: Attempting to delete rows: {rows}")  # Debug log

        reply = QMessageBox.question(self, "Подтверждение удаления",
                                   f"Вы уверены, что хотите удалить {len(rows)} запись(ей)?",
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            # Store data for undo before deletion
            current_table_name = self.table_combo.itemData(self.table_combo.currentIndex())
            deleted_data = []
            for row in rows:
                row_data = {}
                for col in range(self.model.columnCount()):
                    index = self.model.index(row, col)
                    header = self.model.headerData(col, Qt.Horizontal)
                    value = self.model.data(index)
                    row_data[header] = value
                    print(f"DEBUG: Storing data for undo - Row {row}, Col {header}: {value}")  # Debug log
                deleted_data.append(row_data)
            
            print(f"DEBUG: Stored delete operation data: {deleted_data}")  # Debug log
            
            self.model.database().transaction()
            try:
                success = True
                for row in rows:
                    if not self.model.removeRow(row):
                        print(f"DEBUG: Error removing row {row}: {self.model.lastError().text()}")  # Debug log
                        success = False
                        break
                
                if success:
                    if self.model.submitAll():
                        self.model.database().commit()
                        print(f"DEBUG: Successfully deleted {len(rows)} rows")  # Debug log
                        # Store operation for undo
                        self.last_operation = "delete"
                        self.last_operation_data = {
                            "table": current_table_name,
                            "data": deleted_data
                        }
                        print(f"DEBUG: Set last operation to delete with data: {self.last_operation_data}")  # Debug log
                        self.undo_btn.setEnabled(True)
                        self.model.select()
                    else:
                        print(f"DEBUG: Error submitting delete changes: {self.model.lastError().text()}")  # Debug log
                        self.model.database().rollback()
                        QMessageBox.critical(self, "Ошибка удаления", 
                                           f"Не удалось сохранить изменения после удаления: {self.model.lastError().text()}")
                else:
                    self.model.database().rollback()
                    QMessageBox.critical(self, "Ошибка удаления", 
                                       f"Не удалось удалить одну из строк: {self.model.lastError().text()}")

            except Exception as e:
                print(f"DEBUG: Exception during delete: {str(e)}")  # Debug log
                self.model.database().rollback()
                QMessageBox.critical(self, "Ошибка транзакции", f"Произошла ошибка во время удаления: {e}")

    def fill_test_data(self):
        if not self.db or not self.db.isOpen():
            QMessageBox.warning(self, "Нет базы данных", "Пожалуйста, сначала откройте или создайте базу данных.")
            return
        
        reply = QMessageBox.question(self, "Заполнение тестовыми данными",
                                     "Это действие перезапишет существующие данные в базе.\nВы уверены, что хотите продолжить?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            db_path = self.db.databaseName()
            try:
                fill_test_data(db_path)
                QMessageBox.information(self, "Успех", "База данных заполнена тестовыми данными.")
                if self.model:
                    self.model.select()
            except Exception as e:
                 QMessageBox.critical(self, "Ошибка", f"Не удалось заполнить базу данных: {e}")

    def delete_database(self):
        if not self.db or not self.db.isOpen():
            QMessageBox.warning(self, "Нет базы данных", "Нет открытой базы данных для удаления.")
            return

        db_path = self.db.databaseName()
        reply = QMessageBox.critical(self, "Подтверждение удаления БД",
                                     f"Вы уверены, что хотите НАВСЕГДА удалить файл базы данных?\n{db_path}\n\nЭто действие НЕОБРАТИМО!",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            table_name = self.table_combo.currentText()
            if self.model:
                self.model.clear()
            self.db.close()
            print(f"Закрыта база данных перед удалением: {db_path}")
            connection_name = self.db.connectionName()
            QtSql.QSqlDatabase.removeDatabase(connection_name)
            self.db = None
            self.model = None
            self.table_combo.clear()
            self.table_view.setModel(None)
            self.settings.remove("database/lastOpened")
            self.update_button_states(db_open=False)

            try:
                os.remove(db_path)
                QMessageBox.information(self, "Успех", f"Файл базы данных удален: {db_path}")
                self.setWindowTitle("Система управления ремонта вагонного оборудования")

            except OSError as e:
                QMessageBox.critical(self, "Ошибка удаления файла", f"Не удалось удалить файл {db_path}: {e}")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Произошла ошибка при удалении файла: {e}")

    def show_worker_payment_dialog(self):
        if not self.db or not self.db.isOpen():
            QMessageBox.warning(self, "Нет базы данных", "Пожалуйста, сначала откройте или создайте базу данных.")
            return
        dialog = WorkerPaymentDialog(self.db, self)
        dialog.exec_()

    def show_excel_report_dialog(self):
        if not self.db or not self.db.isOpen():
            QMessageBox.warning(self, "Нет базы данных", "Пожалуйста, сначала откройте или создайте базу данных.")
            return
        dialog = ExcelReportDialog(self.db, self)
        dialog.exec_()

    def show_contract_report_dialog(self):
        if not self.db or not self.db.isOpen():
            QMessageBox.warning(self, "Нет базы данных", "Пожалуйста, сначала откройте или создайте базу данных.")
            return
        dialog = ContractReportDialog(self.db, self)
        dialog.exec_()

    def show_fill_word_dialog(self):
        print("DEBUG: Метод show_fill_word_dialog вызван")
        
        if not self.db or not self.db.isOpen():
            QMessageBox.warning(self, "Нет базы данных", "Пожалуйста, сначала откройте или создайте базу данных.")
            return
            
        # Блокируем кнопку, чтобы предотвратить двойной клик
        # Отключаем сигнал, чтобы предотвратить повторный вызов
        self.word_report_btn.setEnabled(False)
        self.word_report_btn.clicked.disconnect(self.show_fill_word_dialog)
        
        try:
            template_path, _ = QFileDialog.getOpenFileName(
                self, "Выберите шаблон Word (.docx)", "", "Word Documents (*.docx)"
            )
            print(f"DEBUG: Выбран шаблон: {template_path}")
            if not template_path:
                print("DEBUG: Пользователь отменил выбор шаблона")
                return

            try:
                placeholders = extract_placeholders(template_path)
                print(f"DEBUG: Найдены маркеры: {placeholders}")
                if not placeholders:
                    QMessageBox.information(self, "Нет маркеров", "В выбранном шаблоне не найдено маркеров вида [имя_маркера].")
                    return
            except Exception as e:
                print(f"DEBUG: Ошибка при извлечении маркеров: {e}")
                QMessageBox.critical(self, "Ошибка чтения шаблона", f"Не удалось прочитать маркеры из шаблона: {e}")
                return

            dialog = QDialog(self)
            dialog.setWindowTitle("Заполнение данных для шаблона Word")
            dialog.setMinimumWidth(500)
            form_layout = QFormLayout(dialog)

            standard_fields = {
                "номер_акта": ("Номер акта:", QLineEdit()),
                "дата_акта": ("Дата акта:", QDateEdit(calendarPopup=True, date=QDate.currentDate())),
                "объем_ТО": ("Объем ТО:", QLineEdit()),
                "город": ("Город:", QLineEdit("Санкт-Петербург")),
            }
            field_widgets = {}
            for key, (label_text, widget) in standard_fields.items():
                form_layout.addRow(label_text, widget)
                field_widgets[key] = widget

            # Отфильтровываем функциональные маркеры из обработки в диалоге
            contract_placeholders = [p for p in placeholders if p.startswith("договор")]
            wagon_placeholders = [p for p in placeholders if p.startswith("вагон")]
            service_placeholders = [p for p in placeholders if p.startswith("услуг") or p.startswith("работ")]
            
            # Отфильтровываем функциональные маркеры (содержащие открывающую и закрывающую скобки)
            regular_placeholders = [p for p in placeholders if not (
                p.startswith("договор") or 
                p.startswith("вагон") or 
                p.startswith("услуг") or 
                p.startswith("работ") or
                p in standard_fields or
                func_pattern.match(p) or  # Проверяем соответствие паттерну функции
                '(' in p or             # Альтернативная проверка скобок
                p.startswith("сумма") or  
                p.startswith("список_")
            )]

            if contract_placeholders or service_placeholders:
                contract_combo = QComboBox()
                query = QtSql.QSqlQuery(self.db)
                query.exec_("SELECT id, номер FROM договоры ORDER BY номер")
                while query.next():
                    contract_combo.addItem(query.value(1), query.value(0))
                form_layout.addRow("Выберите договор:", contract_combo)
                field_widgets['договор_combo'] = contract_combo

            if wagon_placeholders:
                wagon_combo = QComboBox()
                query = QtSql.QSqlQuery(self.db)
                query.exec_("SELECT id, номер FROM вагоны ORDER BY номер")
                while query.next():
                    wagon_combo.addItem(query.value(1), query.value(0))
                form_layout.addRow("Выберите вагон:", wagon_combo)
                field_widgets['вагон_combo'] = wagon_combo

            if any(p.startswith("исполнитель") for p in placeholders):
                worker_combo = QComboBox()
                query = QtSql.QSqlQuery(self.db)
                query.exec_("SELECT id, фио FROM исполнители ORDER BY фио")
                while query.next():
                    worker_combo.addItem(query.value(1), query.value(0))
                form_layout.addRow("Выберите исполнителя:", worker_combo)
                field_widgets['исполнитель_combo'] = worker_combo
                
            # Добавляем форму для обычных переменных, которых нет в БД
            if regular_placeholders:
                form_layout.addRow(QLabel("--- Прочие данные ---"))
                for placeholder in regular_placeholders:
                    # Проверяем, является ли плейсхолдер функциональным (содержит скобки)
                    if '(' in placeholder or ')' in placeholder:
                        continue
                    
                    # Специальная обработка полей подписанта
                    if placeholder.startswith("подписант"):
                        if "должность" in placeholder.lower():
                            continue  # Пропускаем поле должности
                        widget = QLineEdit()
                        form_layout.addRow(f"[{placeholder}]:", widget)
                        field_widgets[placeholder] = widget
                    else:
                        widget = QLineEdit()
                        form_layout.addRow(f"[{placeholder}]:", widget)
                        field_widgets[placeholder] = widget

            button_box = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
            button_box.accepted.connect(dialog.accept)
            button_box.rejected.connect(dialog.reject)
            form_layout.addRow(button_box)

            if dialog.exec_() == QDialog.Accepted:
                mapping = {}
                
                for key, widget in field_widgets.items():
                    if key in standard_fields:
                        if isinstance(standard_fields[key][1], QLineEdit):
                            mapping[key] = widget.text()
                        elif isinstance(standard_fields[key][1], QDateEdit):
                            mapping[key] = widget.date().toString("dd.MM.yyyy")
                    elif key not in ['договор_combo', 'вагон_combo', 'исполнитель_combo']:
                        mapping[key] = widget.text()

                if 'договор_combo' in field_widgets:
                    contract_id = field_widgets['договор_combo'].currentData()
                    contract_number = field_widgets['договор_combo'].currentText()
                    mapping['договоры.номер'] = contract_number
                    query = QtSql.QSqlQuery(self.db)
                    query.prepare("SELECT дата FROM договоры WHERE id = ?")
                    query.addBindValue(contract_id)
                    if query.exec_() and query.next():
                        mapping['договоры.дата'] = QDate.fromString(query.value(0), "yyyy-MM-dd").toString("dd MMMM yyyy г.")
                    
                    if any(p.startswith("список_работ") or p.startswith("список_услуг") for p in placeholders):
                        services_query = QtSql.QSqlQuery(self.db)
                        services_query.prepare("""
                            SELECT у.наименование 
                            FROM услуги у 
                            JOIN договорные_услуги ду ON у.id = ду.id_услуги
                            WHERE ду.id_договора = ?
                        """)
                        services_query.addBindValue(contract_id)
                        services_list = []
                        if services_query.exec_():
                            while services_query.next():
                                services_list.append(services_query.value(0))
                        service_list_str = "LIST:" + "|".join(services_list) if services_list else "Нет услуг по договору"
                        
                        # Добавляем значения для каждого маркера списка работ/услуг с его оригинальным именем
                        for placeholder in placeholders:
                            if placeholder.startswith("список_работ") or placeholder.startswith("список_услуг"):
                                mapping[placeholder] = service_list_str
                        
                        # Также добавляем базовые ключи для обратной совместимости
                        mapping['список_работ'] = service_list_str
                        mapping['список_услуг'] = service_list_str
                        
                        # Рассчитываем сумму стоимости услуг для маркеров вида сумма(...)
                        sum_query = QtSql.QSqlQuery(self.db)
                        sum_query.prepare("""
                            SELECT SUM(у.стоимость_с_ндс)
                            FROM услуги у 
                            JOIN договорные_услуги ду ON у.id = ду.id_услуги
                            WHERE ду.id_договора = ?
                        """)
                        sum_query.addBindValue(contract_id)
                        total_sum = 0
                        if sum_query.exec_() and sum_query.next():
                            total_sum = sum_query.value(0) or 0
                        
                        # Добавляем переменные для суммы по договору
                        for placeholder in placeholders:
                            if placeholder.startswith("сумма"):
                                mapping[placeholder] = f"{total_sum:.2f} руб."

                if 'вагон_combo' in field_widgets:
                    wagon_id = field_widgets['вагон_combo'].currentData()
                    wagon_number = field_widgets['вагон_combo'].currentText()
                    mapping['вагоны.номер'] = wagon_number
                    query = QtSql.QSqlQuery(self.db)
                    query.prepare("SELECT собственник, подразделение, дата_ремонта FROM вагоны WHERE id = ?")
                    query.addBindValue(wagon_id)
                    if query.exec_() and query.next():
                        mapping['вагоны.собственник'] = query.value(0) or ""
                        mapping['вагоны.подразделение'] = query.value(1) or ""
                        date_str = query.value(2)
                        if date_str:
                            mapping['вагоны.дата_ремонта'] = QDate.fromString(date_str, "yyyy-MM-dd").toString("dd.MM.yyyy")
                        else:
                            mapping['вагоны.дата_ремонта'] = ""

                if 'исполнитель_combo' in field_widgets:
                    worker_id = field_widgets['исполнитель_combo'].currentData()
                    worker_fio = field_widgets['исполнитель_combo'].currentText()
                    mapping['исполнители.фио'] = worker_fio

                    if 'исполнители.фио' in mapping and mapping['исполнители.фио']:
                        fio_parts = mapping['исполнители.фио'].split()
                        if len(fio_parts) >= 1:
                            mapping['исполнитель_буква'] = fio_parts[0][0].upper() if fio_parts[0] else ""
                        else:
                            mapping['исполнитель_буква'] = ""

                print("--- Данные для замены ---")
                for k, v in mapping.items():
                    print(f"[{k}] -> {v}")
                print("------------------------")
                
                # Теперь, после сбора всех данных, запрашиваем путь для сохранения
                output_path, _ = QFileDialog.getSaveFileName(
                    self, "Сохранить заполненный документ Word", "", "Word Documents (*.docx)"
                )
                print(f"DEBUG: Выбран выходной файл: {output_path}")
                if not output_path:
                    print("DEBUG: Пользователь отменил выбор выходного файла")
                    return
                if not output_path.endswith(".docx"):
                    output_path += ".docx"
                
                try:
                    replace_placeholders(template_path, output_path, mapping)
                    QMessageBox.information(self, "Успех", f"Документ успешно создан:\n{output_path}")
                except FileNotFoundError as e:
                    QMessageBox.critical(self, "Ошибка", f"Файл не найден: {e}")
                except PermissionError as e:
                    QMessageBox.critical(self, "Ошибка прав доступа", f"Нет прав на запись файла: {e}")
                except ValueError as e:
                    QMessageBox.critical(self, "Ошибка", f"{e}")
                except Exception as e:
                    QMessageBox.critical(self, "Ошибка заполнения", f"Не удалось заполнить шаблон: {e}\n\nПроверьте маркеры в шаблоне и введенные данные.")
        
        finally:
            # Восстанавливаем нормальное состояние кнопки
            print("DEBUG: Восстанавливаем состояние кнопки")
            self.word_report_btn.setEnabled(True)
            self.word_report_btn.clicked.connect(self.show_fill_word_dialog)

    def closeEvent(self, event):
        self.settings.setValue("geometry", self.saveGeometry())
        if self.db and self.db.isOpen():
            self.db.close()
            print(f"Закрыта база данных при выходе: {self.db.databaseName()}")
        event.accept()

    def import_from_excel(self):
        if not self.db or not self.db.isOpen() or not self.model:
            QMessageBox.warning(self, "Импорт из Excel", "Пожалуйста, сначала откройте базу данных и выберите таблицу.")
            return

        print("DEBUG: Entered import_from_excel") # 1

        current_table_name = self.table_combo.itemData(self.table_combo.currentIndex())
        if not current_table_name:
            QMessageBox.warning(self, "Импорт из Excel", "Не выбрана таблица для импорта.")
            print("DEBUG: No table selected for import. Exiting import_from_excel.") # Added
            return
        print(f"DEBUG: Current table for import: {current_table_name}") # 2

        excel_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл Excel", "", "Excel Files (*.xlsx *.xls)"
        )
        if not excel_path:
            print("DEBUG: No Excel file selected by user. Exiting import_from_excel.") # 3
            return
        print(f"DEBUG: Excel file path selected: {excel_path}") # 4

        try:
            print("DEBUG: Top-level try block in import_from_excel entered.") # 5
            xls = pd.ExcelFile(excel_path)
            sheet_names = xls.sheet_names
            print(f"DEBUG: pd.ExcelFile successful. Sheet names: {sheet_names}") # 6

            if not sheet_names:
                QMessageBox.warning(self, "Импорт из Excel", "В Excel файле не найдено листов.")
                print("DEBUG: No sheets found in Excel file. Exiting.") # Added
                return

            sheet_name_to_import = sheet_names[0]
            if len(sheet_names) > 1:
                print("DEBUG: Multiple sheets found. Prompting user for selection.") # Added
                sheet_name_to_import, ok = QtWidgets.QInputDialog.getItem(
                    self, "Выбор листа", "Выберите лист для импорта:", sheet_names, 0, False
                )
                if not ok or not sheet_name_to_import:
                    print("DEBUG: User cancelled sheet selection or no sheet selected. Exiting.") # Added
                    return
            print(f"DEBUG: Sheet selected for import: {sheet_name_to_import}") # 7
            
            df = pd.read_excel(xls, sheet_name=sheet_name_to_import)
            print(f"DEBUG: pd.read_excel successful. DataFrame shape: {df.shape}") # 8
            
            if df.empty:
                QMessageBox.information(self, "Импорт из Excel", f"Лист '{sheet_name_to_import}' пуст.")
                print(f"DEBUG: DataFrame for sheet '{sheet_name_to_import}' is empty. Exiting.") # Added
                return

            table_record = self.model.record()
            table_columns = [table_record.fieldName(i) for i in range(table_record.count())]
            print(f"DEBUG: Database table columns: {table_columns}") # Added
            
            excel_columns_lower = {col.lower(): col for col in df.columns}
            mapped_columns = {} 
            insert_columns_ordered = [] 
            
            for tc in table_columns:
                if tc.lower() in excel_columns_lower:
                    mapped_columns[tc] = excel_columns_lower[tc.lower()]
                    insert_columns_ordered.append(tc)

            print(f"DEBUG: Columns from Excel mapped to table: {mapped_columns}") # Added
            print(f"DEBUG: Ordered columns for SQL INSERT: {insert_columns_ordered}") # 9
            
            if not insert_columns_ordered:
                QMessageBox.warning(self, "Импорт из Excel", 
                                    "Не удалось сопоставить ни одного столбца из Excel с таблицей.\n"
                                    "Убедитесь, что названия столбцов в Excel совпадают с названиями в таблице (регистр не важен).")
                print("DEBUG: No columns were successfully mapped. Exiting.") # 10
                return

            print("DEBUG: Starting database transaction.") # 11
            self.db.transaction()
            query = QtSql.QSqlQuery(self.db)
            
            placeholders = ", ".join(["?"] * len(insert_columns_ordered))
            sql_insert = f"INSERT INTO {current_table_name} ({', '.join(insert_columns_ordered)}) VALUES ({placeholders})"
            print(f"DEBUG: Constructed SQL Insert Statement: {sql_insert}") # 12
            
            inserted_rows = 0
            failed_rows = 0

            for index, row_data in df.iterrows():
                print(f"DEBUG: Processing Excel row index: {index}") # 13
                values_to_bind = []
                problematic_row_data = {} 
                try:
                    for table_col_name in insert_columns_ordered:
                        excel_col_name = mapped_columns[table_col_name]
                        original_value = row_data[excel_col_name]
                        processed_value = None

                        if pd.isna(original_value): 
                            processed_value = None
                        elif isinstance(original_value, (datetime, pd.Timestamp)):
                            ts_value = pd.to_datetime(original_value)
                            if ts_value.hour == 0 and ts_value.minute == 0 and ts_value.second == 0 and \
                               ts_value.microsecond == 0 and getattr(ts_value, 'nanosecond', 0) == 0:
                                processed_value = ts_value.strftime("%Y-%m-%d")
                            else:
                                processed_value = ts_value.strftime("%Y-%m-%d %H:%M:%S")
                        elif type(original_value) is bool: # Check for bool first, as bool is subclass of int
                            processed_value = int(original_value)
                        elif type(original_value) is int: # Check for Python int (already know it's not bool)
                            processed_value = int(original_value)
                        elif type(original_value) is float: # Check for Python float
                            processed_value = float(original_value)
                        elif isinstance(original_value, str):
                            processed_value = original_value
                        else: # Fallback for other types (e.g., numpy numerics, or unexpected types)
                            # Try to convert to a Python numeric type if possible, otherwise string
                            try:
                                num_val = float(original_value) # Try float conversion first
                                if num_val.is_integer():
                                    processed_value = int(num_val)
                                else:
                                    processed_value = num_val
                                print(f"DEBUG:   Converted non-standard numeric/unknown type '{original_value}' ({type(original_value)}) to {type(processed_value)}: {processed_value}")
                            except (ValueError, TypeError):
                                print(f"DEBUG:   Warning: Value '{original_value}' of type {type(original_value)} is being converted to string as a final fallback.")
                                processed_value = str(original_value)
                        
                        values_to_bind.append(processed_value)
                        problematic_row_data[table_col_name] = processed_value 
                        # print(f"DEBUG:   Table col: {table_col_name}, Excel col: {excel_col_name}, Original: {original_value}, Processed: {processed_value}") # 14 - Too verbose for now

                except Exception as e_prepare_bind_values:
                    print(f"DEBUG: ERROR during value preparation for row {index}: {type(e_prepare_bind_values).__name__} - {str(e_prepare_bind_values)}") # Added
                    if self.db.inTransaction():
                        print("DEBUG: Rolling back transaction due to value preparation error.") # Added
                        self.db.rollback()
                    QMessageBox.critical(self, "Ошибка обработки строки Excel",
                                         f"Ошибка при обработке данных из строки Excel (индекс {index}): {e_prepare_bind_values}\n"
                                         f"Данные строки (обработанные): {problematic_row_data}\n"
                                         "Импорт отменен.")
                    print("DEBUG: Exiting import_from_excel due to value preparation error.") # Added
                    return 

                print(f"DEBUG:   Values to bind for row {index}: {values_to_bind}") # 15

                if not query.prepare(sql_insert):
                    error_text = query.lastError().text()
                    print(f"DEBUG: SQL prepare FAILED: {error_text}") # 16
                    QMessageBox.critical(self, "Ошибка SQL", f"Ошибка подготовки запроса: {error_text}\nSQL: {sql_insert}")
                    if self.db.inTransaction():
                        print("DEBUG: Rolling back transaction due to SQL prepare failure.") # Added
                        self.db.rollback()
                    print("DEBUG: Exiting import_from_excel due to SQL prepare failure.") # Added
                    return
                print(f"DEBUG: SQL prepare SUCCEEDED for row {index}.") # 17

                for val_idx, val in enumerate(values_to_bind):
                    query.bindValue(val_idx, val)
                print(f"DEBUG: Values bound for row {index}.") # 18
                
                if query.exec_():
                    inserted_rows += 1
                    print(f"DEBUG: SQL exec SUCCEEDED for row {index}. inserted_rows: {inserted_rows}") # 19
                else:
                    failed_rows += 1
                    error_text = query.lastError().text()
                    print(f"DEBUG: SQL exec FAILED for row {index}: {error_text}. failed_rows: {failed_rows}") # 20
                    print(f"DEBUG:   Failed SQL: {sql_insert}") # Added
                    print(f"DEBUG:   Failed Values: {values_to_bind}") # Added
                    
            print(f"DEBUG: Finished processing all rows. Inserted: {inserted_rows}, Failed: {failed_rows}") # 21
            
            if failed_rows > 0:
                print("DEBUG: Rolling back transaction due to one or more failed row insertions.") # 22
                self.db.rollback()
                QMessageBox.warning(self, "Импорт из Excel", 
                                    f"Импорт завершен с ошибками.\n"
                                    f"Успешно вставлено: {inserted_rows} строк.\n"
                                    f"Не удалось вставить: {failed_rows} строк.\n"
                                    "Изменения отменены. Проверьте консоль для деталей.")
            else:
                print("DEBUG: Attempting to commit transaction as all rows were processed (or no rows to process).") # 23
                if self.db.commit():
                    print("DEBUG: Transaction committed successfully.") # 24
                    QMessageBox.information(self, "Успех", f"Успешно импортировано {inserted_rows} строк в таблицу '{current_table_name}'.")
                    if self.model:
                        self.model.select()
                else:
                    error_text = self.db.lastError().text()
                    print(f"DEBUG: Transaction commit FAILED: {error_text}") # 25
                    QMessageBox.critical(self, "Ошибка фиксации", f"Не удалось зафиксировать транзакцию: {error_text}")
                    # Attempt to rollback again if commit failed, though it might already be in an invalid state
                    if self.db.inTransaction(): # Check if still in transaction
                         print("DEBUG: Rolling back transaction due to commit failure.") # Added
                         self.db.rollback()

        except Exception as e_outer:
            print(f"DEBUG: CRITICAL ERROR in import_from_excel (outer try-except): {type(e_outer).__name__} - {str(e_outer)}") # 26
            import traceback
            print("DEBUG: Full traceback of the critical error:") # 27
            traceback.print_exc()
            if self.db.isOpen() and self.db.inTransaction():
                print("DEBUG: Rolling back transaction due to critical error in outer try-except.") # Added
                self.db.rollback()
            QMessageBox.critical(self, "Критическая ошибка импорта", f"Произошла критическая ошибка при импорте: {e_outer}\nПроверьте консоль.")
        finally:
            print("DEBUG: Exiting import_from_excel function (finally block).") # 28

    # Добавляем метод для редактирования записи
    def edit_record(self):
        if not self.model:
            QMessageBox.warning(self, "Нет таблицы", "Пожалуйста, сначала выберите таблицу.")
            return
            
        selection = self.table_view.selectionModel()
        if not selection.hasSelection():
            QMessageBox.warning(self, "Редактирование записи", "Пожалуйста, выберите строку для редактирования.")
            return
            
        # Получаем индекс выбранной строки
        selected_rows = selection.selectedRows()
        if not selected_rows:
            rows = sorted({idx.row() for idx in selection.selectedIndexes()})
            if not rows:
                return
            row = rows[0]
        else:
            row = selected_rows[0].row()
            
        # Создаем и отображаем диалог редактирования
        dialog = EditRecordDialog(self.model, row, self)
        if dialog.exec_() == QDialog.Accepted:
            # После успешного редактирования обновляем отображение
            self.model.select()

    def undo_last_operation(self):
        print(f"DEBUG: Attempting to undo operation: {self.last_operation}")  # Debug log
        print(f"DEBUG: Operation data: {self.last_operation_data}")  # Debug log
        
        if not self.last_operation or not self.last_operation_data:
            print("DEBUG: No operation to undo")  # Debug log
            return

        if self.last_operation == "add":
            print("DEBUG: Undoing add operation")  # Debug log
            # Undo add operation by deleting the last row
            row = self.last_operation_data["row"]
            print(f"DEBUG: Attempting to remove row {row}")  # Debug log
            
            self.model.database().transaction()
            try:
                if self.model.removeRow(row):
                    if self.model.submitAll():
                        self.model.database().commit()
                        print("DEBUG: Successfully undid add operation")  # Debug log
                        self.model.select()
                        self.last_operation = None
                        self.last_operation_data = None
                        self.undo_btn.setEnabled(False)
                    else:
                        print(f"DEBUG: Error submitting undo changes: {self.model.lastError().text()}")  # Debug log
                        self.model.database().rollback()
                        QMessageBox.critical(self, "Ошибка отмены", 
                                           f"Не удалось отменить добавление записи: {self.model.lastError().text()}")
                else:
                    print(f"DEBUG: Error removing row: {self.model.lastError().text()}")  # Debug log
                    QMessageBox.critical(self, "Ошибка отмены", 
                                       f"Не удалось удалить добавленную запись: {self.model.lastError().text()}")
            except Exception as e:
                print(f"DEBUG: Exception during undo add: {str(e)}")  # Debug log
                self.model.database().rollback()
                QMessageBox.critical(self, "Ошибка отмены", f"Произошла ошибка при отмене добавления: {e}")

        elif self.last_operation == "delete":
            print("DEBUG: Undoing delete operation")  # Debug log
            # Undo delete operation by restoring deleted rows
            table_name = self.last_operation_data["table"]
            deleted_data = self.last_operation_data["data"]
            print(f"DEBUG: Restoring {len(deleted_data)} rows to table {table_name}")  # Debug log
            
            self.model.database().transaction()
            try:
                success = True
                for row_data in deleted_data:
                    row = self.model.rowCount()
                    print(f"DEBUG: Attempting to insert row at index {row}")  # Debug log
                    if not self.model.insertRow(row):
                        print(f"DEBUG: Error inserting row: {self.model.lastError().text()}")  # Debug log
                        success = False
                        break
                    
                    # Restore data for each column
                    for col in range(self.model.columnCount()):
                        header = self.model.headerData(col, Qt.Horizontal)
                        if header in row_data:
                            value = row_data[header]
                            index = self.model.index(row, col)
                            print(f"DEBUG: Setting column {header} to value {value}")  # Debug log
                            if not self.model.setData(index, value):
                                print(f"DEBUG: Error setting data for column {header}: {self.model.lastError().text()}")  # Debug log
                                success = False
                                break
                
                if success:
                    if self.model.submitAll():
                        self.model.database().commit()
                        print("DEBUG: Successfully restored deleted rows")  # Debug log
                        self.model.select()
                        self.last_operation = None
                        self.last_operation_data = None
                        self.undo_btn.setEnabled(False)
                    else:
                        print(f"DEBUG: Error submitting restored rows: {self.model.lastError().text()}")  # Debug log
                        self.model.database().rollback()
                        QMessageBox.critical(self, "Ошибка отмены", 
                                           f"Не удалось сохранить восстановленные записи: {self.model.lastError().text()}")
                else:
                    print("DEBUG: Failed to restore all data")  # Debug log
                    self.model.database().rollback()
                    QMessageBox.critical(self, "Ошибка отмены", "Не удалось восстановить удаленные записи")
            
            except Exception as e:
                print(f"DEBUG: Exception during undo: {str(e)}")  # Debug log
                self.model.database().rollback()
                QMessageBox.critical(self, "Ошибка отмены", f"Произошла ошибка при отмене удаления: {e}")

    def update_button_states(self, db_open):
        is_table_selected = bool(self.table_combo.currentText()) and db_open
        
        self.word_report_btn.setEnabled(db_open)
        self.excel_report_btn.setEnabled(db_open)
        self.contract_report_btn.setEnabled(db_open)
        self.worker_payment_btn.setEnabled(db_open)
        self.table_combo.setEnabled(db_open)
        self.add_record_btn.setEnabled(is_table_selected)
        self.edit_record_btn.setEnabled(is_table_selected)
        self.delete_record_btn.setEnabled(is_table_selected)
        self.import_excel_btn.setEnabled(is_table_selected)
        
        # Update undo button state based on last operation
        has_undo_operation = bool(self.last_operation and self.last_operation_data)
        print(f"DEBUG: Updating undo button state - has operation: {has_undo_operation}")  # Debug log
        self.undo_btn.setEnabled(has_undo_operation)

    def register_undo_add(self, table_name, row_id):
        print(f"DEBUG: register_undo_add: table={table_name}, row_id={row_id}")
        self.last_operation = "add"
        self.last_operation_data = {"table": table_name, "row_id": row_id}
        self.undo_btn.setEnabled(True)

    def register_undo_delete(self, table_name, deleted_rows_data):
        print(f"DEBUG: register_undo_delete: table={table_name}, data={deleted_rows_data}")
        self.last_operation = "delete"
        self.last_operation_data = {"table": table_name, "data": deleted_rows_data}
        self.undo_btn.setEnabled(True)

    def undo_last_operation(self):
        print(f"DEBUG: Attempting to undo operation: {self.last_operation}")
        print(f"DEBUG: Operation data: {self.last_operation_data}")
        if not self.last_operation or not self.last_operation_data:
            print("DEBUG: No operation to undo")
            return

        if self.last_operation == "add":
            table_name = self.last_operation_data["table"]
            row_id = self.last_operation_data["row_id"]
            print(f"DEBUG: Undoing add in table {table_name} for id {row_id}")
            # Удаляем запись по id
            query = QtSql.QSqlQuery(self.db)
            query.prepare(f"DELETE FROM {table_name} WHERE id = ?")
            query.addBindValue(row_id)
            if query.exec_():
                print("DEBUG: Successfully undid add operation (deleted row by id)")
                self.model.select()
                self.last_operation = None
                self.last_operation_data = None
                self.undo_btn.setEnabled(False)
            else:
                print(f"DEBUG: Error deleting row by id: {query.lastError().text()}")
                QMessageBox.critical(self, "Ошибка отмены", f"Не удалось отменить добавление записи: {query.lastError().text()}")

        elif self.last_operation == "delete":
            table_name = self.last_operation_data["table"]
            deleted_data = self.last_operation_data["data"]
            print(f"DEBUG: Restoring {len(deleted_data)} rows to table {table_name}")
            success = True
            for row_data in deleted_data:
                fields = list(row_data.keys())
                placeholders = ','.join(['?'] * len(fields))
                sql = f"INSERT INTO {table_name} ({','.join(fields)}) VALUES ({placeholders})"
                query = QtSql.QSqlQuery(self.db)
                query.prepare(sql)
                for field in fields:
                    query.addBindValue(row_data[field])
                if not query.exec_():
                    print(f"DEBUG: Error restoring row: {query.lastError().text()}")
                    success = False
            if success:
                print("DEBUG: Successfully restored deleted rows")
                self.model.select()
                self.last_operation = None
                self.last_operation_data = None
                self.undo_btn.setEnabled(False)
            else:
                QMessageBox.critical(self, "Ошибка отмены", "Не удалось восстановить удаленные записи")

class ManageContractServicesDialog(QDialog):
    def __init__(self, db, contract_id=None, parent=None):
        super().__init__(parent)
        self.db = db
        self.passed_initial_contract_id = contract_id # Store the ID that was passed in
        self.contract_id = None # This will be updated by on_contract_changed
        self.setWindowTitle("Управление услугами по договору")
        self.setup_ui()
        self.resize(800, 600)

    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        # Contract selection
        contract_layout = QHBoxLayout()
        contract_label = QLabel("Договор:")
        self.contract_combo = QComboBox()
        self.contract_combo.currentIndexChanged.connect(self.on_contract_changed)
        # Populate contracts and attempt to select the initial one
        self.load_contracts(self.passed_initial_contract_id) 

        contract_layout.addWidget(contract_label)
        contract_layout.addWidget(self.contract_combo)
        layout.addLayout(contract_layout)
        
        # --- Service Selection Group --- 
        self.service_selection_group = QGroupBox("Выбор услуг")
        service_selection_layout = QVBoxLayout(self.service_selection_group)
        # layout.addWidget(self.service_selection_group) # Added later

        # Mode selection (Individual or Range)
        mode_layout = QHBoxLayout()
        self.individual_mode_radio = QRadioButton("Выбрать отдельные услуги")
        self.individual_mode_radio.setChecked(True)
        self.individual_mode_radio.toggled.connect(self.toggle_service_selection_mode)
        mode_layout.addWidget(self.individual_mode_radio)

        self.range_mode_radio = QRadioButton("Выбрать диапазон услуг (по ID)")
        self.range_mode_radio.toggled.connect(self.toggle_service_selection_mode)
        mode_layout.addWidget(self.range_mode_radio)
        service_selection_layout.addLayout(mode_layout)

        # Range input fields (initially hidden)
        self.range_input_widget = QWidget()
        range_input_layout = QHBoxLayout(self.range_input_widget)
        range_input_layout.setContentsMargins(0,0,0,0)
        range_input_layout.addWidget(QLabel("ID с:"))
        self.id_from_edit = QLineEdit()
        self.id_from_edit.setPlaceholderText("Начальный ID")
        range_input_layout.addWidget(self.id_from_edit)
        range_input_layout.addWidget(QLabel("ID до:"))
        self.id_to_edit = QLineEdit()
        self.id_to_edit.setPlaceholderText("Конечный ID")
        range_input_layout.addWidget(self.id_to_edit)
        self.range_input_widget.setVisible(False)
        service_selection_layout.addWidget(self.range_input_widget)
        
        # Services selection (Available vs Contract)
        self.available_services_label = QLabel("Доступные услуги (для выбора отдельных):")
        service_selection_layout.addWidget(self.available_services_label)
        
        services_layout = QHBoxLayout()
        
        # Left panel with available services
        self.left_panel_widget = QWidget() # Create a parent widget for easier show/hide
        left_panel = QVBoxLayout(self.left_panel_widget)
        left_panel.setContentsMargins(0,0,0,0)

        self.available_services = QTableView()
        self.available_services.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.available_services.horizontalHeader().setStretchLastSection(True)
        self.available_services.doubleClicked.connect(self.edit_available_service)
        left_panel.addWidget(self.available_services)
        
        self.edit_available_btn = QPushButton("Изменить доступную услугу")
        self.edit_available_btn.clicked.connect(self.edit_available_service)
        left_panel.addWidget(self.edit_available_btn)
        
        
        services_layout.addWidget(self.left_panel_widget) # Add the parent widget
        
        # Buttons
        buttons_panel_layout = QVBoxLayout()
        self.add_btn = QPushButton(">>")
        self.add_btn.setToolTip("Добавить выбранные услуги к договору")
        self.add_btn.clicked.connect(self.add_services)
        self.remove_btn = QPushButton("<<")
        self.remove_btn.setToolTip("Удалить выбранные услуги из договора")
        self.remove_btn.clicked.connect(self.remove_services)
        buttons_panel_layout.addWidget(self.add_btn)
        buttons_panel_layout.addWidget(self.remove_btn)
        buttons_panel_layout.addStretch()
        services_layout.addLayout(buttons_panel_layout)
        
        # Right panel with contract services
        right_panel = QVBoxLayout()
        right_panel.addWidget(QLabel("Услуги по договору:")) # Added label for clarity
        self.contract_services = QTableView()
        self.contract_services.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.contract_services.horizontalHeader().setStretchLastSection(True)
        self.contract_services.doubleClicked.connect(self.edit_contract_service)
        right_panel.addWidget(self.contract_services)
        
        edit_contract_btn = QPushButton("Изменить услугу договора")
        edit_contract_btn.clicked.connect(self.edit_contract_service)
        right_panel.addWidget(edit_contract_btn)
        
        services_layout.addLayout(right_panel)
        
        service_selection_layout.addLayout(services_layout) # Add the main services hbox to the groupbox
        
        # Add the group box to the main dialog layout
        layout.addWidget(self.service_selection_group)
        
        # Close button
        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)

        # Manually call on_contract_changed AFTER all UI elements for the group are created
        # to set the initial state of service_selection_group and its children
        # based on the initially selected contract (if any).
        if self.contract_combo.count() > 0:
            self.on_contract_changed(self.contract_combo.currentIndex()) # Pass current index
        else: # No contracts to select, ensure group is disabled and sub-elements are toggled
            self.service_selection_group.setEnabled(False)
            self.toggle_service_selection_mode() 

    def toggle_service_selection_mode(self):
        is_range_mode = self.range_mode_radio.isChecked()
        
        self.range_input_widget.setVisible(is_range_mode)
        self.left_panel_widget.setVisible(not is_range_mode)
        
        if is_range_mode:
            self.available_services_label.setText("Добавление услуг по диапазону ID:")
            self.add_btn.setToolTip("Добавить услуги из указанного диапазона ID")
            # Clear selection in the hidden table to avoid confusion if it was visible before
            self.available_services.clearSelection()
            # Clear the model for available_services as it's not used in range mode
            if self.available_services.model() and isinstance(self.available_model, QtSql.QSqlTableModel):
                self.available_model.setFilter("1=0") # Effectively clear by setting an impossible filter
                self.available_model.select()
        else: # Individual mode
            self.available_services_label.setText("Доступные услуги (для выбора отдельных):")
            self.add_btn.setToolTip("Добавить выбранные услуги к договору")
            # Reload available services when switching to individual mode
            self.load_available_services()
        
        # It might be good to reload available services if individual mode is activated
        # if not is_range_mode:
        #     self.load_available_services()
        # else:
        #     # Optionally clear or hide the available_services model/view if not already hidden by left_panel_widget
        #     if self.available_services.model():
        #          # Cast to QSqlTableModel to call clear if it's the right type
        #         model_to_clear = self.available_services.model()
        #         if isinstance(model_to_clear, QtSql.QSqlTableModel):
        #             # model_to_clear.clear() # This might not be what we want; better to set an empty filter
        #             model_to_clear.setFilter("1=0") # Set an impossible filter
        #             model_to_clear.select()
        #         elif hasattr(model_to_clear, 'setQuery'): # For QSqlQueryModel
        #             model_to_clear.setQuery(QtSql.QSqlQuery(self.db)) # Clear by setting an empty query
            
    def edit_available_service(self):
        selection = self.available_services.selectionModel()
        if not selection.hasSelection():
            QMessageBox.warning(self, "Редактирование услуги", "Пожалуйста, выберите услугу для редактирования.")
            return
            
        # Получаем индекс выбранной строки
        selected_rows = selection.selectedRows()
        if not selected_rows:
            rows = sorted({idx.row() for idx in selection.selectedIndexes()})
            if not rows:
                return
            row = rows[0]
        else:
            row = selected_rows[0].row()
        
        # Получаем модель таблицы
        model = self.available_services.model()
        if not model:
            return
            
        # Создаем и отображаем диалог редактирования
        dialog = EditRecordDialog(model, row, self)
        if dialog.exec_() == QDialog.Accepted:
            # После успешного редактирования обновляем отображение
            self.load_available_services()
            
    def edit_contract_service(self):
        selection = self.contract_services.selectionModel()
        if not selection.hasSelection():
            QMessageBox.warning(self, "Редактирование услуги договора", "Пожалуйста, выберите услугу для редактирования.")
            return
            
        # Получаем индекс выбранной строки
        selected_rows = selection.selectedRows()
        if not selected_rows:
            rows = sorted({idx.row() for idx in selection.selectedIndexes()})
            if not rows:
                return
            row = rows[0]
        else:
            row = selected_rows[0].row()
        
        # Получаем модель таблицы
        model = self.contract_services.model()
        if not model:
            return
            
        # Создаем и отображаем диалог редактирования
        dialog = EditRecordDialog(model, row, self)
        if dialog.exec_() == QDialog.Accepted:
            # После успешного редактирования обновляем отображение
            self.load_contract_services()

    def load_available_services(self):
        contract_id = self.contract_combo.currentData()
        # Ensure contract_id is an int for the SQL query if it's not None
        # This handles QVariant from itemData directly if it contains an int
        if isinstance(contract_id, QVariant):
            contract_id_val = contract_id.value()
        else:
            contract_id_val = contract_id

        if not contract_id_val: # if contract_id_val is None or 0 (invalid ID)
             # Clear the model if no valid contract is selected or if in range mode (handled by toggle_service_selection_mode)
            if hasattr(self, 'available_model') and self.available_model:
                self.available_model.setFilter("1=0")
                self.available_model.select()
            return
            
        # Create the model for available services - use table model instead of query model
        self.available_model = QtSql.QSqlTableModel(self, self.db)
        self.available_model.setTable("услуги")
        # Filter to only show services not already in the contract
        self.available_model.setFilter(f"id NOT IN (SELECT id_услуги FROM договорные_услуги WHERE id_договора = {contract_id_val})")
        self.available_model.setEditStrategy(QtSql.QSqlTableModel.OnFieldChange)
        self.available_model.select()
        
        # Set headers
        self.available_model.setHeaderData(0, Qt.Horizontal, "ID")
        self.available_model.setHeaderData(1, Qt.Horizontal, "Наименование")
        self.available_model.setHeaderData(2, Qt.Horizontal, "Цена (без НДС)")
        self.available_model.setHeaderData(3, Qt.Horizontal, "Цена (с НДС)")
        
        self.available_services.setModel(self.available_model)
        self.available_services.hideColumn(0)  # Hide ID column
        self.available_services.resizeColumnsToContents()

    def load_contract_services(self):
        contract_id = self.contract_combo.currentData()
        if isinstance(contract_id, QVariant):
            contract_id_val = contract_id.value()
        else:
            contract_id_val = contract_id
        
        if not contract_id_val:
            if hasattr(self, 'contract_services_rel_model') and self.contract_services_rel_model:
                self.contract_services_rel_model.setFilter("1=0")
                self.contract_services_rel_model.select()
            return
            
        # Use a dedicated model for the relation table that can be edited
        self.contract_services_rel_model = QtSql.QSqlRelationalTableModel(self, self.db)
        self.contract_services_rel_model.setTable("договорные_услуги")
        self.contract_services_rel_model.setFilter(f"id_договора = {contract_id_val}")
        self.contract_services_rel_model.setRelation(1, QtSql.QSqlRelation("договоры", "id", "номер"))
        self.contract_services_rel_model.setRelation(2, QtSql.QSqlRelation("услуги", "id", "наименование"))
        self.contract_services_rel_model.setEditStrategy(QtSql.QSqlTableModel.OnFieldChange)
        self.contract_services_rel_model.select()
        
        # Set headers
        self.contract_services_rel_model.setHeaderData(0, Qt.Horizontal, "ID")
        self.contract_services_rel_model.setHeaderData(1, Qt.Horizontal, "Договор")
        self.contract_services_rel_model.setHeaderData(2, Qt.Horizontal, "Услуга")
        
        self.contract_services.setModel(self.contract_services_rel_model)
        self.contract_services.resizeColumnsToContents()

    def on_contract_changed(self, index):
        # Get current data from combo, this is the definitive source now for self.contract_id
        data = self.contract_combo.itemData(index) # Use index to be sure
        
        # Extract the Python value from QVariant if necessary
        if isinstance(data, QVariant):
            current_contract_id = data.value()
        else:
            current_contract_id = data 
        
        # Check if it's a valid ID (not None, not 0 if 0 is used as invalid placeholder data)
        # For QVariant containing None, .value() should yield None.
        # For QVariant containing int 0, .value() yields int 0.
        # We use None for the placeholder data with QVariant().
        is_contract_selected = current_contract_id is not None

        self.contract_id = current_contract_id # Update the instance variable used by other methods

        self.service_selection_group.setEnabled(is_contract_selected)
        self.load_data() # Load data related to the new contract_id (or clear if None)

        if not is_contract_selected:
            # Reset to individual mode if contract is deselected (placeholder selected)
            if not self.individual_mode_radio.isChecked(): # Avoid redundant signal if already checked
                self.individual_mode_radio.setChecked(True) 
            else: # if already individual mode, still need to ensure toggle runs for UI consistency
                self.toggle_service_selection_mode()
        else:
            self.toggle_service_selection_mode() # Ensure UI consistency for the selected mode

    def add_services(self):
        contract_id = self.contract_combo.currentData()
        if not contract_id:
            QMessageBox.warning(self, "Ошибка", "Выберите договор")
            return

        if self.individual_mode_radio.isChecked():
            # --- Standard individual selection --- 
            selected_rows = self.available_services.selectionModel().selectedRows()
            if not selected_rows:
                QMessageBox.warning(self, "Ошибка", "Выберите услуги для добавления")
                return
            
            services_to_add = []
            for index in selected_rows:
                service_id = self.available_model.data(self.available_model.index(index.row(), 0))
                services_to_add.append(service_id)
        
        elif self.range_mode_radio.isChecked():
            # --- Range selection by ID ---
            try:
                id_from = int(self.id_from_edit.text())
                id_to = int(self.id_to_edit.text())
                if id_from <= 0 or id_to <= 0:
                    QMessageBox.warning(self, "Ошибка ввода", "ID должны быть положительными числами.")
                    return
                if id_from > id_to:
                    QMessageBox.warning(self, "Ошибка ввода", "Начальный ID не может быть больше конечного ID.")
                    return
            except ValueError:
                QMessageBox.warning(self, "Ошибка ввода", "ID должны быть числовыми значениями.")
                return

            range_query = QtSql.QSqlQuery(self.db)
            # Select services within range that are not already linked to the contract
            sql = """
                SELECT id FROM услуги 
                WHERE id >= ? AND id <= ? 
                AND id NOT IN (SELECT id_услуги FROM договорные_услуги WHERE id_договора = ?)
            """
            range_query.prepare(sql)
            range_query.addBindValue(id_from)
            range_query.addBindValue(id_to)
            range_query.addBindValue(contract_id)
            
            services_to_add = []
            if range_query.exec_():
                while range_query.next():
                    services_to_add.append(range_query.value(0))
            else:
                QMessageBox.critical(self, "Ошибка SQL", f"Ошибка при поиске услуг по диапазону ID: {range_query.lastError().text()}")
                return
            
            if not services_to_add:
                QMessageBox.information(self, "Информация", "Не найдено новых услуг в указанном диапазоне ID для добавления.")
                return
        else:
            # Should not happen if radio buttons are set up correctly
            return

        # --- Common logic for adding services_to_add list ---
        self.db.transaction()
        try:
            insert_query = QtSql.QSqlQuery(self.db)
            insert_query.prepare("""
                INSERT INTO договорные_услуги (id_договора, id_услуги)
                VALUES (?, ?)
            """)
            
            success_count = 0
            for service_id in services_to_add:
                insert_query.bindValue(0, contract_id)
                insert_query.bindValue(1, service_id)
                
                if insert_query.exec_():
                    success_count += 1
                else:
                    # Log specific error for this service_id if needed
                    print(f"DEBUG: SQL Error adding service {service_id}: {insert_query.lastError().text()}")
            
            if success_count > 0:
                self.db.commit()
                QMessageBox.information(self, "Успех", f"Добавлено {success_count} услуг к договору")
                self.load_data() # Reload both available and contract services lists
            elif services_to_add: # If there were services to add but none succeeded
                self.db.rollback()
                QMessageBox.warning(self, "Ошибка", "Не удалось добавить выбранные услуги. Проверьте консоль.")
            # If services_to_add was empty (e.g. range had no new services), no message needed here as it was handled above

        except Exception as e:
            self.db.rollback()
            QMessageBox.critical(self, "Ошибка транзакции", f"Ошибка при добавлении услуг: {e}")

    def remove_services(self):
        contract_id = self.contract_combo.currentData()
        if not contract_id:
            QMessageBox.warning(self, "Ошибка", "Выберите договор")
            return
            
        selected_rows = self.contract_services.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, "Ошибка", "Выберите услуги для удаления")
            return
            
        # Ask for confirmation
        reply = QMessageBox.question(
            self, "Подтверждение удаления",
            f"Вы действительно хотите удалить {len(selected_rows)} услуг из договора?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
            
        # Start transaction
        self.db.transaction()
        
        try:
            delete_query = QtSql.QSqlQuery(self.db)
            delete_query.prepare("""
                DELETE FROM договорные_услуги 
                WHERE id_договора = ? AND id_услуги = ?
            """)
            
            success_count = 0
            
            for index in selected_rows:
                # Fix: Use the new contract_services_rel_model variable which is a QSqlRelationalTableModel
                # We need to get the id_услуги which is in column 2 for the relation model
                service_id = self.contract_services_rel_model.data(self.contract_services_rel_model.index(index.row(), 2))
                
                delete_query.bindValue(0, contract_id)
                delete_query.bindValue(1, service_id)
                
                if delete_query.exec_():
                    success_count += 1
                else:
                    print(f"DEBUG: SQL Error: {delete_query.lastError().text()}")
            
            if success_count > 0:
                self.db.commit()
                QMessageBox.information(self, "Успех", f"Удалено {success_count} услуг из договора")
                self.load_data()
            else:
                self.db.rollback()
                QMessageBox.warning(self, "Ошибка", "Не удалось удалить услуги")
                
        except Exception as e:
            self.db.rollback()
            QMessageBox.critical(self, "Ошибка", f"Ошибка при удалении услуг: {e}")

    def manage_contract_services(self):
        print("DEBUG: ManageContractServicesDialog.manage_contract_services called")
        contract_id = self.contract_combo.currentData()
        if not contract_id:
            QMessageBox.warning(self, "Ошибка", "Выберите договор")
            return
            
        dialog = ManageContractServicesDialog(self.db, contract_id, self)
        if dialog.exec_() == QDialog.Accepted:
            # Reload services after management
            self.load_services()

    # New method to save changes in available services
    def save_available_changes(self):
        if hasattr(self, 'available_model'):
            if self.available_model.submitAll():
                QMessageBox.information(self, "Успех", "Изменения сохранены")
            else:
                QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить изменения: {self.available_model.lastError().text()}")

    # New method to save changes in contract services
    def save_contract_changes(self):
        if hasattr(self, 'contract_services_rel_model'):
            if self.contract_services_rel_model.submitAll():
                QMessageBox.information(self, "Успех", "Изменения сохранены")
            else:
                QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить изменения: {self.contract_services_rel_model.lastError().text()}")

    # Добавляем недостающие методы
    def load_contracts(self, initial_contract_to_select_id):
        self.contract_combo.blockSignals(True) # Block signals during population/setting index
        self.contract_combo.clear()
        
        current_selection_idx = -1 # Default to no specific selection
        placeholder_idx = 0 # Index of the placeholder if added

        # Always add a placeholder item first. Its data is QVariant() which .value() becomes None.
        self.contract_combo.addItem("--- Выберите договор ---", QVariant())
        
        query = QtSql.QSqlQuery(self.db)
        query.exec_("SELECT id, номер FROM договоры ORDER BY номер")
        
        while query.next():
            contract_id_val = query.value(0)
            contract_number = query.value(1)
            self.contract_combo.addItem(contract_number, contract_id_val)
            if contract_id_val == initial_contract_to_select_id:
                current_selection_idx = self.contract_combo.count() - 1
            
        if current_selection_idx != -1:
            self.contract_combo.setCurrentIndex(current_selection_idx)
        else:
            self.contract_combo.setCurrentIndex(placeholder_idx) # Default to placeholder
        
        self.contract_combo.blockSignals(False)

    def load_data(self):
        self.load_available_services()
        self.load_contract_services()
        
    def manage_contract_services(self):
        print("DEBUG: ManageContractServicesDialog.manage_contract_services called")
        contract_id = self.contract_combo.currentData()
        if not contract_id:
            QMessageBox.warning(self, "Ошибка", "Выберите договор")
            return
            
        dialog = ManageContractServicesDialog(self.db, contract_id, self)
        if dialog.exec_() == QDialog.Accepted:
            # Reload services after management
            self.load_services()
            
    def add_services(self):
        contract_id = self.contract_combo.currentData()
        if not contract_id:
            QMessageBox.warning(self, "Ошибка", "Выберите договор")
            return

        if self.individual_mode_radio.isChecked():
            # --- Standard individual selection --- 
            selected_rows = self.available_services.selectionModel().selectedRows()
            if not selected_rows:
                QMessageBox.warning(self, "Ошибка", "Выберите услуги для добавления")
                return
            
            services_to_add = []
            for index in selected_rows:
                service_id = self.available_model.data(self.available_model.index(index.row(), 0))
                services_to_add.append(service_id)
        
        elif self.range_mode_radio.isChecked():
            # --- Range selection by ID ---
            try:
                id_from = int(self.id_from_edit.text())
                id_to = int(self.id_to_edit.text())
                if id_from <= 0 or id_to <= 0:
                    QMessageBox.warning(self, "Ошибка ввода", "ID должны быть положительными числами.")
                    return
                if id_from > id_to:
                    QMessageBox.warning(self, "Ошибка ввода", "Начальный ID не может быть больше конечного ID.")
                    return
            except ValueError:
                QMessageBox.warning(self, "Ошибка ввода", "ID должны быть числовыми значениями.")
                return

            range_query = QtSql.QSqlQuery(self.db)
            # Select services within range that are not already linked to the contract
            sql = """
                SELECT id FROM услуги 
                WHERE id >= ? AND id <= ? 
                AND id NOT IN (SELECT id_услуги FROM договорные_услуги WHERE id_договора = ?)
            """
            range_query.prepare(sql)
            range_query.addBindValue(id_from)
            range_query.addBindValue(id_to)
            range_query.addBindValue(contract_id)
            
            services_to_add = []
            if range_query.exec_():
                while range_query.next():
                    services_to_add.append(range_query.value(0))
            else:
                QMessageBox.critical(self, "Ошибка SQL", f"Ошибка при поиске услуг по диапазону ID: {range_query.lastError().text()}")
                return
            
            if not services_to_add:
                QMessageBox.information(self, "Информация", "Не найдено новых услуг в указанном диапазоне ID для добавления.")
                return
        else:
            # Should not happen if radio buttons are set up correctly
            return

        # --- Common logic for adding services_to_add list ---
        self.db.transaction()
        try:
            insert_query = QtSql.QSqlQuery(self.db)
            insert_query.prepare("""
                INSERT INTO договорные_услуги (id_договора, id_услуги)
                VALUES (?, ?)
            """)
            
            success_count = 0
            for service_id in services_to_add:
                insert_query.bindValue(0, contract_id)
                insert_query.bindValue(1, service_id)
                
                if insert_query.exec_():
                    success_count += 1
                else:
                    # Log specific error for this service_id if needed
                    print(f"DEBUG: SQL Error adding service {service_id}: {insert_query.lastError().text()}")
            
            if success_count > 0:
                self.db.commit()
                QMessageBox.information(self, "Успех", f"Добавлено {success_count} услуг к договору")
                self.load_data() # Reload both available and contract services lists
            elif services_to_add: # If there were services to add but none succeeded
                self.db.rollback()
                QMessageBox.warning(self, "Ошибка", "Не удалось добавить выбранные услуги. Проверьте консоль.")
            # If services_to_add was empty (e.g. range had no new services), no message needed here as it was handled above

        except Exception as e:
            self.db.rollback()
            QMessageBox.critical(self, "Ошибка транзакции", f"Ошибка при добавлении услуг: {e}")

    def remove_services(self):
        contract_id = self.contract_combo.currentData()
        if not contract_id:
            QMessageBox.warning(self, "Ошибка", "Выберите договор")
            return
            
        selected_rows = self.contract_services.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, "Ошибка", "Выберите услуги для удаления")
            return
            
        # Ask for confirmation
        reply = QMessageBox.question(
            self, "Подтверждение удаления",
            f"Вы действительно хотите удалить {len(selected_rows)} услуг из договора?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
            
        # Start transaction
        self.db.transaction()
        
        try:
            delete_query = QtSql.QSqlQuery(self.db)
            delete_query.prepare("""
                DELETE FROM договорные_услуги 
                WHERE id_договора = ? AND id_услуги = ?
            """)
            
            success_count = 0
            
            for index in selected_rows:
                # Fix: Use the new contract_services_rel_model variable which is a QSqlRelationalTableModel
                # We need to get the id_услуги which is in column 2 for the relation model
                service_id = self.contract_services_rel_model.data(self.contract_services_rel_model.index(index.row(), 2))
                
                delete_query.bindValue(0, contract_id)
                delete_query.bindValue(1, service_id)
                
                if delete_query.exec_():
                    success_count += 1
                else:
                    print(f"DEBUG: SQL Error: {delete_query.lastError().text()}")
            
            if success_count > 0:
                self.db.commit()
                QMessageBox.information(self, "Успех", f"Удалено {success_count} услуг из договора")
                self.load_data()
            else:
                self.db.rollback()
                QMessageBox.warning(self, "Ошибка", "Не удалось удалить услуги")
                
        except Exception as e:
            self.db.rollback()
            QMessageBox.critical(self, "Ошибка", f"Ошибка при удалении услуг: {e}")

class AddContractDialog(QDialog):
    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("Добавление договора")
        self.setup_ui()

    def setup_ui(self):
        layout = QFormLayout(self)
        
        # Номер договора
        self.number_edit = QLineEdit()
        self.number_edit.setPlaceholderText("2024.123456")
        layout.addRow("Номер договора:", self.number_edit)
        
        # Дата договора
        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(QDate.currentDate())
        self.date_edit.setDisplayFormat("dd.MM.yyyy")
        layout.addRow("Дата договора:", self.date_edit)
        
        # Кнопки
        buttons_layout = QHBoxLayout()
        save_btn = QPushButton("Сохранить")
        save_btn.clicked.connect(self.save_contract)
        cancel_btn = QPushButton("Отмена")
        cancel_btn.clicked.connect(self.reject)
        buttons_layout.addWidget(save_btn)
        buttons_layout.addWidget(cancel_btn)
        layout.addRow(buttons_layout)
        
    def save_contract(self):
        number = self.number_edit.text().strip()
        if not number:
            QMessageBox.warning(self, "Ошибка", "Номер договора обязателен для заполнения")
            return
            
        date = self.date_edit.date().toString("yyyy-MM-dd")
        
        query = QtSql.QSqlQuery(self.db)
        query.prepare("INSERT INTO договоры (номер, дата) VALUES (?, ?)")
        query.addBindValue(number)
        query.addBindValue(date)
        
        if query.exec_():
            QMessageBox.information(self, "Успех", "Договор успешно добавлен")
            self.accept()
        else:
            error = query.lastError().text()
            QMessageBox.critical(self, "Ошибка", f"Ошибка при добавлении договора: {error}")

class EditRecordDialog(QDialog):
    def __init__(self, model, row, parent=None):
        super().__init__(parent)
        self.model = model
        self.row = row
        self.setWindowTitle("Редактирование записи")
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)

        form_layout = QFormLayout()
        self.field_widgets = {}

        record = self.model.record(self.row) # Получаем структуру записи для текущей строки
                                             # или self.model.record() если структура одинакова для всех строк

        for col in range(self.model.columnCount()):
            header = self.model.headerData(col, Qt.Horizontal)
            field_name = self.model.record().fieldName(col) # Получаем фактическое имя поля из модели
            value = self.model.data(self.model.index(self.row, col))
            current_cell_flags = self.model.flags(self.model.index(self.row, col))
            is_editable = bool(current_cell_flags & Qt.ItemIsEditable)

            if isinstance(value, QDate) or (isinstance(value, str) and self.is_date_string(value)):
                widget = QDateEdit()
                widget.setCalendarPopup(True)
                widget.setDisplayFormat("dd.MM.yyyy")
                widget.setKeyboardTracking(True)
                widget.setReadOnly(False)
                if isinstance(value, QDate):
                    widget.setDate(value)
                elif value:
                    parsed_date = QDate.fromString(value, "yyyy-MM-dd")
                    if parsed_date.isValid():
                        widget.setDate(parsed_date)
                    else:
                        widget.setDate(QDate.currentDate())
                else:
                    widget.setDate(QDate.currentDate())
            else:
                widget = QLineEdit()
                if value is not None:
                    widget.setText(str(value))

            # Отключаем редактирование для поля 'id' (без учета регистра) или если флаг ItemIsEditable отсутствует
            if field_name.lower() == "id" or not is_editable:
                widget.setEnabled(False)

            form_layout.addRow(f"{header}:", widget)
            self.field_widgets[col] = widget

        layout.addLayout(form_layout)

        buttons_layout = QHBoxLayout()
        save_btn = QPushButton("Сохранить")
        save_btn.clicked.connect(self.save_changes)
        cancel_btn = QPushButton("Отмена")
        cancel_btn.clicked.connect(self.reject)
        buttons_layout.addWidget(save_btn)
        buttons_layout.addWidget(cancel_btn)
        layout.addLayout(buttons_layout)

    def is_date_string(self, text):
        if not isinstance(text, str):
            return False
        try:
            # Проверяем несколько распространенных форматов дат
            formats_to_check = ["yyyy-MM-dd", "dd.MM.yyyy", "yyyy-MM-ddTHH:mm:ss.zzzZ", "yyyy-MM-dd HH:mm:ss"]
            for fmt in formats_to_check:
                if QDate.fromString(text, fmt).isValid():
                    return True
                # Для datetime строк также проверяем QDateTime
                if QDateTime.fromString(text, fmt).isValid():
                    return True
            return False
        except Exception:
            return False

    def save_changes(self):
        for col, widget in self.field_widgets.items():
            field_name = self.model.record().fieldName(col) # Получаем фактическое имя поля
            current_cell_flags = self.model.flags(self.model.index(self.row, col))

            # Пропускаем сохранение для поля 'id' или если оно нередактируемое
            if field_name.lower() == "id" or not (current_cell_flags & Qt.ItemIsEditable):
                continue

            if isinstance(widget, QDateEdit):
                value = widget.date().toString("yyyy-MM-dd")
            else:
                value = widget.text()

            self.model.setData(self.model.index(self.row, col), value)

        if hasattr(self.model, 'submitAll'):
            if self.model.submitAll():
                QMessageBox.information(self, "Успех", "Изменения сохранены")
                self.accept()
            else:
                QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить изменения: {self.model.lastError().text()}")
        else:
            # Если у модели нет submitAll (например, это не QSqlTableModel, а кастомная),
            # просто принимаем диалог. Предполагается, что setData уже применило изменения.
            self.accept()

# Добавляем новый класс диалога для добавления услуги
class AddServiceDialog(QDialog):
    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("Добавление услуги")
        self.setup_ui()

    def setup_ui(self):
        layout = QFormLayout(self)
        
        # Наименование услуги (обязательное поле)
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("Введите наименование услуги")
        layout.addRow("Наименование услуги *:", self.name_edit)
        
        # Стоимость без НДС (обязательное поле)
        self.price_no_vat_edit = QLineEdit()
        self.price_no_vat_edit.setPlaceholderText("0.00")
        layout.addRow("Стоимость без НДС *:", self.price_no_vat_edit)
        
        # Стоимость с НДС
        self.price_with_vat_edit = QLineEdit()
        self.price_with_vat_edit.setPlaceholderText("0.00")
        layout.addRow("Стоимость с НДС *:", self.price_with_vat_edit)
        
        # Стоимость работнику
        self.worker_price_edit = QLineEdit()
        self.worker_price_edit.setPlaceholderText("0.00")
        layout.addRow("Оплата работнику *:", self.worker_price_edit)
        
        # Описание услуги
        self.description_edit = QLineEdit()
        self.description_edit.setPlaceholderText("Описание услуги (необязательно)")
        layout.addRow("Описание:", self.description_edit)
        
        # Автоматический расчет НДС при изменении стоимости без НДС
        self.price_no_vat_edit.textChanged.connect(self.calculate_vat)
        
        # Обязательные поля помечены звездочкой
        note_label = QLabel("* Обязательные поля")
        layout.addRow(note_label)
        
        # Кнопки
        buttons_layout = QHBoxLayout()
        self.save_btn = QPushButton("Сохранить")
        self.save_btn.clicked.connect(self.save_service)
        self.save_btn.setEnabled(False)  # Изначально кнопка неактивна
        cancel_btn = QPushButton("Отмена")
        cancel_btn.clicked.connect(self.reject)
        buttons_layout.addWidget(self.save_btn)
        buttons_layout.addWidget(cancel_btn)
        layout.addRow(buttons_layout)
        
        # Проверяем обязательные поля при изменении
        self.name_edit.textChanged.connect(self.check_required_fields)
        self.price_no_vat_edit.textChanged.connect(self.check_required_fields)
        self.price_with_vat_edit.textChanged.connect(self.check_required_fields)
        self.worker_price_edit.textChanged.connect(self.check_required_fields)

    def calculate_vat(self):
        try:
            price_no_vat = float(self.price_no_vat_edit.text().replace(',', '.'))
            price_with_vat = price_no_vat * 1.2  # НДС 20%
            self.price_with_vat_edit.setText(f"{price_with_vat:.2f}")
        except ValueError:
            # Если введено не число, не делаем ничего
            pass

    def check_required_fields(self):
        # Проверяем, что все обязательные поля заполнены
        has_name = bool(self.name_edit.text().strip())
        try:
            price_no_vat = float(self.price_no_vat_edit.text().replace(',', '.')) if self.price_no_vat_edit.text() else 0
            price_with_vat = float(self.price_with_vat_edit.text().replace(',', '.')) if self.price_with_vat_edit.text() else 0
            worker_price = float(self.worker_price_edit.text().replace(',', '.')) if self.worker_price_edit.text() else 0
            has_prices = price_no_vat > 0 and price_with_vat > 0 and worker_price > 0
        except ValueError:
            has_prices = False
        
        self.save_btn.setEnabled(has_name and has_prices)

    def save_service(self):
        name = self.name_edit.text().strip()
        if not name:
            QMessageBox.warning(self, "Ошибка", "Наименование услуги обязательно для заполнения")
            return
        
        try:
            price_no_vat = float(self.price_no_vat_edit.text().replace(',', '.'))
            price_with_vat = float(self.price_with_vat_edit.text().replace(',', '.'))
            worker_price = float(self.worker_price_edit.text().replace(',', '.'))
        except ValueError:
            QMessageBox.warning(self, "Ошибка", "Некорректно указана стоимость услуги")
            return
        
        description = self.description_edit.text().strip()
        
        query = QtSql.QSqlQuery(self.db)
        query.prepare("""
            INSERT INTO услуги (наименование, стоимость_без_ндс, стоимость_с_ндс, стоимость_работнику, описание)
            VALUES (?, ?, ?, ?, ?)
        """)
        query.addBindValue(name)
        query.addBindValue(price_no_vat)
        query.addBindValue(price_with_vat)
        query.addBindValue(worker_price)
        query.addBindValue(description)
        
        if query.exec_():
            QMessageBox.information(self, "Успех", "Услуга успешно добавлена")
            self.accept()
        else:
            error = query.lastError().text()
            QMessageBox.critical(self, "Ошибка", f"Ошибка при добавлении услуги: {error}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    
    editor = SQLiteEditor()
    
    geometry = editor.settings.value("geometry")
    if geometry:
        editor.restoreGeometry(geometry)
    else:
        screen_geometry = QApplication.desktop().screenGeometry()
        x = (screen_geometry.width() - editor.width()) / 2
        y = (screen_geometry.height() - editor.height()) / 2
        editor.move(int(x), int(y))

    editor.show()
    sys.exit(app.exec_())
