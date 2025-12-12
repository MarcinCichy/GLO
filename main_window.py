import os
import time
import calendar
import logging

import holidays
# Import konieczny dla pliku EXE
import holidays.countries.poland

import win32api
import win32print

from PyQt5.QtCore import Qt, QDate
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, QPushButton,
    QGroupBox, QCalendarWidget, QMessageBox, QFileDialog, QListWidget, QListWidgetItem,
    QDialog, QAbstractItemView, QProgressBar
)

from config_store import ConfigStore
from worker_thread import WorkerThread
from ui_dialogs import PrinterDialog, ExcelPreviewDialog, EdytorPracownikow
from ui_widgets import KlikalnyKalendarz


class GeneratorListObecnosci(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Generator List Obecności v20.0 (Final Complete)")
        self.setGeometry(100, 100, 1100, 800)
        self.generated_files_map = {}

        self.config_store = ConfigStore()
        self.last_save_folder = self.config_store.last_save_folder

        self.initUI()
        self.load_employees_on_startup()

    def save_config(self):
        self.config_store.last_save_folder = self.last_save_folder
        self.config_store.save()

    def initUI(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout()

        # LEWY
        left_panel = QVBoxLayout()

        date_group = QGroupBox("1. Wybierz okres")
        date_layout = QVBoxLayout()
        self.month_cb = QComboBox()
        self.month_names = ["Styczeń", "Luty", "Marzec", "Kwiecień", "Maj", "Czerwiec",
                            "Lipiec", "Sierpień", "Wrzesień", "Październik", "Listopad", "Grudzień"]
        self.month_cb.addItems(self.month_names)
        self.year_cb = QComboBox()
        current_year = QDate.currentDate().year()
        for y in range(current_year - 1, current_year + 5):
            self.year_cb.addItem(str(y))
        self.year_cb.setCurrentText(str(current_year))
        self.month_cb.currentIndexChanged.connect(self.sync_calendar_from_combo)
        self.year_cb.currentIndexChanged.connect(self.sync_calendar_from_combo)
        date_layout.addWidget(QLabel("Miesiąc:"))
        date_layout.addWidget(self.month_cb)
        date_layout.addWidget(QLabel("Rok:"))
        date_layout.addWidget(self.year_cb)
        date_group.setLayout(date_layout)
        left_panel.addWidget(date_group)

        emp_group = QGroupBox("2. Pracownicy")
        emp_layout = QVBoxLayout()
        self.emp_list_widget = QListWidget()
        self.emp_list_widget.setSelectionMode(QAbstractItemView.NoSelection)
        self.emp_list_widget.itemChanged.connect(self.update_print_button_state)
        # PODWÓJNE KLIKNIĘCIE OTWIERA PODGLĄD WBUDOWANY
        self.emp_list_widget.itemDoubleClicked.connect(self.open_preview_window)
        emp_layout.addWidget(self.emp_list_widget)
        edit_btns_layout = QHBoxLayout()
        self.btn_edit_text = QPushButton("Edytuj listę")
        self.btn_edit_text.clicked.connect(self.open_text_editor)
        self.btn_toggle_select = QPushButton("Odznacz wszystkich")
        self.btn_toggle_select.clicked.connect(self.toggle_selection)
        edit_btns_layout.addWidget(self.btn_edit_text)
        edit_btns_layout.addWidget(self.btn_toggle_select)
        emp_layout.addLayout(edit_btns_layout)
        emp_group.setLayout(emp_layout)
        left_panel.addWidget(emp_group)

        action_group = QGroupBox("4. Akcje")
        action_layout = QVBoxLayout()
        self.generate_btn = QPushButton("Generuj wybrane")
        self.generate_btn.setMinimumHeight(40)
        self.generate_btn.setStyleSheet("font-weight: bold; font-size: 14px; background-color: #4CAF50; color: white;")
        self.generate_btn.clicked.connect(self.start_generation_thread)
        action_layout.addWidget(self.generate_btn)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        action_layout.addWidget(self.progress_bar)

        self.print_btn = QPushButton("Drukuj karty...")
        self.print_btn.setMinimumHeight(40)
        self.print_btn.setStyleSheet("font-weight: bold; font-size: 12px; background-color: #2196F3; color: white;")
        self.print_btn.clicked.connect(self.print_generated_files)
        self.print_btn.setEnabled(False)
        action_layout.addWidget(self.print_btn)

        # PRZYCISKI FOLDERU
        folder_btns_layout = QHBoxLayout()

        self.open_folder_btn = QPushButton("Otwórz folder")
        self.open_folder_btn.clicked.connect(self.open_last_folder)
        self.open_folder_btn.setEnabled(False)

        self.change_folder_btn = QPushButton("Zmień folder zapisu")
        self.change_folder_btn.clicked.connect(self.change_save_folder)

        folder_btns_layout.addWidget(self.open_folder_btn)
        folder_btns_layout.addWidget(self.change_folder_btn)
        action_layout.addLayout(folder_btns_layout)

        # Jeśli folder jest w configu, aktywujemy przycisk
        if self.last_save_folder and os.path.exists(self.last_save_folder):
            self.open_folder_btn.setEnabled(True)

        action_group.setLayout(action_layout)
        left_panel.addWidget(action_group)
        main_layout.addLayout(left_panel, stretch=1)

        # PRAWY
        right_panel = QVBoxLayout()
        cal_group = QGroupBox("3. Kalendarz (Święta i dni wolne)")
        cal_layout = QVBoxLayout()
        cal_layout.addWidget(QLabel(
            "SZARY = Wolne ustawowo | CZERWONY = Wolne dodatkowe\nUżyj rolki myszy na kalendarzu, aby zmienić miesiąc."))
        self.calendar = KlikalnyKalendarz()
        self.calendar.setGridVisible(True)
        self.calendar.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
        self.calendar.setNavigationBarVisible(False)
        self.calendar.currentPageChanged.connect(self.sync_combo_from_calendar)
        self.calendar.clicked.connect(self.toggle_holiday)
        cal_layout.addWidget(self.calendar)
        cal_group.setLayout(cal_layout)
        right_panel.addWidget(cal_group, stretch=2)
        main_layout.addLayout(right_panel)
        main_widget.setLayout(main_layout)

        current_month_idx = QDate.currentDate().month() - 1
        self.month_cb.setCurrentIndex(current_month_idx)
        self.sync_calendar_from_combo()

    def sync_calendar_from_combo(self):
        self.calendar.blockSignals(True)
        month = self.month_cb.currentIndex() + 1
        year = int(self.year_cb.currentText())
        self.calendar.setCurrentPage(year, month)
        self.calendar.cached_holidays = holidays.PL(years=year)
        self.calendar.blockSignals(False)

    def sync_combo_from_calendar(self, year, month):
        self.month_cb.blockSignals(True)
        self.year_cb.blockSignals(True)
        self.year_cb.setCurrentText(str(year))
        self.month_cb.setCurrentIndex(month - 1)
        self.calendar.cached_holidays = holidays.PL(years=year)
        self.month_cb.blockSignals(False)
        self.year_cb.blockSignals(False)

    def load_employees_on_startup(self):
        filename = "pracownicy.txt"
        if not os.path.exists(filename):
            try:
                with open(filename, "w", encoding="utf-8") as f:
                    f.write("Jan Kowalski, Opiekun\nAnna Nowak, Kucharka")
            except Exception:
                pass
        try:
            with open(filename, "r", encoding="utf-8") as f:
                content = f.read()
                self.update_list_widget(content)
        except Exception as e:
            QMessageBox.warning(self, "Błąd", f"Problem z plikiem pracownicy.txt: {e}")

    def update_list_widget(self, text):
        self.emp_list_widget.blockSignals(True)
        self.emp_list_widget.clear()
        lines = text.split('\n')
        for line in lines:
            line = line.strip()
            if line:
                item = QListWidgetItem(line)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Checked)
                self.emp_list_widget.addItem(item)
        self.emp_list_widget.blockSignals(False)
        self.btn_toggle_select.setText("Odznacz wszystkich")

    def open_text_editor(self):
        current_text = ""
        for i in range(self.emp_list_widget.count()):
            current_text += self.emp_list_widget.item(i).text() + "\n"
        dialog = EdytorPracownikow(current_text, self)
        if dialog.exec_():
            new_text = dialog.get_text()
            self.update_list_widget(new_text)
            try:
                with open("pracownicy.txt", "w", encoding="utf-8") as f:
                    f.write(new_text)
            except Exception as e:
                QMessageBox.warning(self, "Błąd zapisu", str(e))

    def toggle_selection(self):
        self.emp_list_widget.blockSignals(True)
        current_text = self.btn_toggle_select.text()
        if current_text == "Odznacz wszystkich":
            for i in range(self.emp_list_widget.count()):
                self.emp_list_widget.item(i).setCheckState(Qt.Unchecked)
            self.btn_toggle_select.setText("Zaznacz wszystkich")
        else:
            for i in range(self.emp_list_widget.count()):
                self.emp_list_widget.item(i).setCheckState(Qt.Checked)
            self.btn_toggle_select.setText("Odznacz wszystkich")
        self.emp_list_widget.blockSignals(False)
        self.update_print_button_state()

    def toggle_holiday(self, date):
        holidays_set = self.calendar.custom_holidays
        if date in holidays_set:
            holidays_set.remove(date)
        else:
            holidays_set.add(date)
        self.calendar.updateCell(date)

    def calculate_holidays(self, year, month):
        final_holidays = set()
        pl_holidays = holidays.PL(years=year)
        days_in_month = calendar.monthrange(year, month)[1]
        for day in range(1, days_in_month + 1):
            current_date = QDate(year, month, day)
            py_date = current_date.toPyDate()
            if current_date.dayOfWeek() >= 6:
                final_holidays.add(current_date)
            elif py_date in pl_holidays:
                final_holidays.add(current_date)
            elif current_date in self.calendar.custom_holidays:
                final_holidays.add(current_date)
        return final_holidays

    def update_print_button_state(self):
        count = 0
        for i in range(self.emp_list_widget.count()):
            item = self.emp_list_widget.item(i)
            if item.checkState() == Qt.Checked and item.text() in self.generated_files_map:
                count += 1

        if count > 0:
            self.print_btn.setEnabled(True)
            self.print_btn.setText(f"Drukuj {count} kart...")
        else:
            if not self.generated_files_map:
                self.print_btn.setEnabled(False)
                self.print_btn.setText("Drukuj karty...")
            else:
                self.print_btn.setEnabled(False)
                self.print_btn.setText("Zaznacz pracowników do druku")

    # --- ZMIANA: ZMIANA FOLDERU ZAPISU ---
    def change_save_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Wybierz nowy folder domyślny")
        if folder:
            self.last_save_folder = folder
            self.save_config()
            self.open_folder_btn.setEnabled(True)
            QMessageBox.information(self, "Sukces", f"Domyślny folder zmieniony na:\n{folder}")

    def start_generation_thread(self):
        employees_to_generate = []
        for i in range(self.emp_list_widget.count()):
            item = self.emp_list_widget.item(i)
            if item.checkState() == Qt.Checked:
                line = item.text()
                parts = line.split(',')
                name = parts[0].strip()
                job = parts[1].strip() if len(parts) > 1 else "Pracownik"
                employees_to_generate.append({'name': name, 'job': job, 'key': line})

        if not employees_to_generate:
            QMessageBox.warning(self, "Błąd", "Nie zaznaczono żadnego pracownika!")
            return

        # Ustalanie folderu docelowego
        target_folder = ""

        # 1. Sprawdzamy, czy jest zapisany w configu
        if self.last_save_folder and os.path.exists(self.last_save_folder):
            target_folder = self.last_save_folder
        else:
            # 2. Jeśli nie ma, tworzymy folder domyślny obok programu
            default_path = os.path.join(os.getcwd(), "Wygenerowane Karty")
            if not os.path.exists(default_path):
                try:
                    os.makedirs(default_path)
                except Exception:
                    pass
            target_folder = default_path

            # Zapisujemy ten domyślny folder jako ostatni używany
            self.last_save_folder = target_folder
            self.save_config()
            self.open_folder_btn.setEnabled(True)

        month_idx = self.month_cb.currentIndex() + 1
        year = int(self.year_cb.currentText())
        month_name = self.month_names[month_idx - 1].upper()
        holidays_map = self.calculate_holidays(year, month_idx)

        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.generate_btn.setEnabled(False)
        self.print_btn.setEnabled(False)

        self.thread = WorkerThread(employees_to_generate, target_folder, month_idx, year, month_name, holidays_map)
        self.thread.progress_updated.connect(self.progress_bar.setValue)
        self.thread.finished.connect(self.on_generation_finished)
        self.thread.error_occurred.connect(self.on_generation_error)
        self.thread.start()

    def on_generation_finished(self, generated_map, folder):
        self.generated_files_map = generated_map
        self.progress_bar.setVisible(False)
        self.generate_btn.setEnabled(True)
        self.update_print_button_state()
        QMessageBox.information(self, "Sukces", f"Wygenerowano plików: {len(generated_map)}\nLokalizacja: {folder}")

    def on_generation_error(self, error_msg):
        self.progress_bar.setVisible(False)
        self.generate_btn.setEnabled(True)
        QMessageBox.critical(self, "Błąd generowania", error_msg)

    def open_last_folder(self):
        if self.last_save_folder and os.path.exists(self.last_save_folder):
            os.startfile(self.last_save_folder)
        else:
            QMessageBox.warning(self, "Info", "Folder nie został jeszcze wybrany lub nie istnieje.")

    def open_preview_window(self, item):
        employee_key = item.text()
        if employee_key in self.generated_files_map:
            filepath = self.generated_files_map[employee_key]
            if os.path.exists(filepath):
                # TO OTWIERA WBUDOWANY PODGLĄD (NOWY DLA WERSJI 20.0)
                preview = ExcelPreviewDialog(filepath, self)
                preview.exec_()
            else:
                QMessageBox.warning(self, "Błąd", "Plik nie istnieje.")
        else:
            QMessageBox.information(self, "Info", "Najpierw wygeneruj karty.")

    def print_generated_files(self):
        files_to_print = []
        for i in range(self.emp_list_widget.count()):
            item = self.emp_list_widget.item(i)
            if item.checkState() == Qt.Checked and item.text() in self.generated_files_map:
                files_to_print.append(self.generated_files_map[item.text()])

        if not files_to_print:
            QMessageBox.warning(self, "Info", "Brak zaznaczonych kart do wydruku.")
            return

        printer_dialog = PrinterDialog(self)
        if printer_dialog.exec_() == QDialog.Accepted:
            selected_printer = printer_dialog.selected_printer
            original_printer = win32print.GetDefaultPrinter()
            try:
                win32print.SetDefaultPrinter(selected_printer)
                for filepath in files_to_print:
                    win32api.ShellExecute(0, "print", filepath, None, ".", 0)
                    time.sleep(2)
                QMessageBox.information(self, "Sukces", f"Wysłano {len(files_to_print)} plików do: {selected_printer}")
            except Exception as e:
                logging.error(f"Błąd druku: {e}")
                QMessageBox.critical(self, "Błąd druku", f"Wystąpił błąd: {e}")
            finally:
                win32print.SetDefaultPrinter(original_printer)
