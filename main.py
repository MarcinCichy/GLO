import sys
import os
import holidays
import calendar  # Do sprawdzania liczby dni w miesiącu
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QLabel, QComboBox, QTextEdit,
                             QPushButton, QGroupBox, QCalendarWidget, QTableView, QMessageBox, QFileDialog)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QTextCharFormat, QColor, QPainter

# Importy do Excela
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter


# === KLASA KALENDARZA ===
class KlikalnyKalendarz(QCalendarWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.custom_holidays = set()
        self.cached_holidays = holidays.PL()

    def paintCell(self, painter, rect, date):
        py_date = date.toPyDate()
        if date in self.custom_holidays:
            painter.save()
            painter.fillRect(rect, QColor("salmon"))
            painter.setPen(Qt.black)
            painter.drawText(rect, Qt.AlignCenter, str(date.day()))
            painter.restore()
        elif date.dayOfWeek() >= 6 or py_date in self.cached_holidays:
            painter.save()
            painter.fillRect(rect, QColor("#D3D3D3"))
            painter.setPen(Qt.black)
            painter.drawText(rect, Qt.AlignCenter, str(date.day()))
            painter.restore()
        else:
            super().paintCell(painter, rect, date)


# === GŁÓWNA APLIKACJA ===
class GeneratorListObecnosci(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("GLO - Etap 3: Generator Strony 1 (Poprawiony)")
        self.setGeometry(100, 100, 950, 650)
        self.initUI()

    def initUI(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout()

        # LEWY PANEL
        left_panel = QVBoxLayout()

        # Data
        date_group = QGroupBox("1. Wybierz okres")
        date_layout = QVBoxLayout()
        self.month_cb = QComboBox()
        self.month_names = ["Styczeń", "Luty", "Marzec", "Kwiecień", "Maj", "Czerwiec",
                            "Lipiec", "Sierpień", "Wrzesień", "Październik", "Listopad", "Grudzień"]
        self.month_cb.addItems(self.month_names)
        self.month_cb.currentIndexChanged.connect(self.sync_calendar_view)

        self.year_cb = QComboBox()
        current_year = QDate.currentDate().year()
        for y in range(current_year - 1, current_year + 5):
            self.year_cb.addItem(str(y))
        self.year_cb.setCurrentText(str(current_year))
        self.year_cb.currentIndexChanged.connect(self.sync_calendar_view)

        date_layout.addWidget(QLabel("Miesiąc:"))
        date_layout.addWidget(self.month_cb)
        date_layout.addWidget(QLabel("Rok:"))
        date_layout.addWidget(self.year_cb)
        date_group.setLayout(date_layout)
        left_panel.addWidget(date_group)

        # Pracownicy
        emp_group = QGroupBox("2. Pracownicy")
        emp_layout = QVBoxLayout()
        self.employees_text = QTextEdit()
        self.employees_text.setPlaceholderText("Jan Kowalski, Opiekun\nAnna Nowak, Kucharka")
        emp_layout.addWidget(self.employees_text)

        btns_layout = QHBoxLayout()
        self.save_list_btn = QPushButton("Zapisz listę")
        self.save_list_btn.clicked.connect(self.save_employees_to_file)
        self.load_list_btn = QPushButton("Wczytaj listę")
        self.load_list_btn.clicked.connect(self.load_employees_from_file)
        btns_layout.addWidget(self.save_list_btn)
        btns_layout.addWidget(self.load_list_btn)
        emp_layout.addLayout(btns_layout)
        emp_group.setLayout(emp_layout)
        left_panel.addWidget(emp_group)

        # GENERUJ
        self.generate_btn = QPushButton("Generuj Pliki (Strona 1)")
        self.generate_btn.setMinimumHeight(50)
        self.generate_btn.setStyleSheet("font-weight: bold; font-size: 14px; background-color: #4CAF50; color: white;")
        self.generate_btn.clicked.connect(self.start_generation)
        left_panel.addWidget(self.generate_btn)

        main_layout.addLayout(left_panel, stretch=1)

        # PRAWY PANEL
        right_panel = QVBoxLayout()
        cal_group = QGroupBox("3. Wybierz dodatkowe dni wolne")
        cal_layout = QVBoxLayout()
        cal_layout.addWidget(QLabel("SZARY = Wolne ustawowo\nCZERWONY = Wolne dodatkowe (kliknij)"))

        self.calendar = KlikalnyKalendarz()
        self.calendar.setGridVisible(True)
        self.calendar.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
        self.calendar.setNavigationBarVisible(False)
        calendar_view = self.calendar.findChild(QTableView)
        if calendar_view:
            calendar_view.setSelectionMode(QTableView.NoSelection)
        self.calendar.clicked.connect(self.toggle_holiday)

        cal_layout.addWidget(self.calendar)
        cal_group.setLayout(cal_layout)
        right_panel.addWidget(cal_group, stretch=2)

        main_layout.addLayout(right_panel)
        main_widget.setLayout(main_layout)

        # Start
        current_month_idx = QDate.currentDate().month() - 1
        self.month_cb.setCurrentIndex(current_month_idx)
        self.sync_calendar_view()

    # --- METODY GUI ---
    def sync_calendar_view(self):
        month = self.month_cb.currentIndex() + 1
        year = int(self.year_cb.currentText())
        self.calendar.setCurrentPage(year, month)
        self.calendar.cached_holidays = holidays.PL(years=year)

    def toggle_holiday(self, date):
        holidays_set = self.calendar.custom_holidays
        if date in holidays_set:
            holidays_set.remove(date)
        else:
            holidays_set.add(date)
        self.calendar.updateCell(date)

    def save_employees_to_file(self):
        try:
            with open("pracownicy.txt", "w", encoding="utf-8") as f:
                f.write(self.employees_text.toPlainText())
            QMessageBox.information(self, "Sukces", "Zapisano!")
        except Exception as e:
            QMessageBox.critical(self, "Błąd", str(e))

    def load_employees_from_file(self):
        if os.path.exists("pracownicy.txt"):
            with open("pracownicy.txt", "r", encoding="utf-8") as f:
                self.employees_text.setPlainText(f.read())

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

    # --- ETAP 3: GENEROWANIE EXCELA ---

    def start_generation(self):
        raw_employees = self.employees_text.toPlainText().strip()
        if not raw_employees:
            QMessageBox.warning(self, "Błąd", "Lista pracowników jest pusta!")
            return

        folder = QFileDialog.getExistingDirectory(self, "Wybierz folder zapisu")
        if not folder:
            return

        month_idx = self.month_cb.currentIndex() + 1
        year = int(self.year_cb.currentText())
        month_name = self.month_names[month_idx - 1].upper()

        holidays_map = self.calculate_holidays(year, month_idx)

        employees = []
        for line in raw_employees.split('\n'):
            parts = line.split(',')
            name = parts[0].strip()
            job = parts[1].strip() if len(parts) > 1 else "Pracownik"
            if name:
                employees.append({'name': name, 'job': job})

        try:
            for emp in employees:
                self.create_excel_file(emp, folder, month_idx, year, month_name, holidays_map)

            QMessageBox.information(self, "Gotowe", f"Wygenerowano plików: {len(employees)}")
        except Exception as e:
            QMessageBox.critical(self, "Błąd krytyczny", str(e))

    def create_excel_file(self, emp, folder, month, year, month_str, holidays_map):
        wb = Workbook()
        ws = wb.active
        ws.title = "Lista Obecności"

        # --- POPRAWKA TUTAJ ---
        # Ustawienia strony A4 (używamy kodu 9 dla A4 i stringa 'portrait')
        ws.page_setup.paperSize = 9
        ws.page_setup.orientation = 'portrait'
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1

        # Marginesy
        ws.page_margins.left = 0.4
        ws.page_margins.right = 0.4
        ws.page_margins.top = 0.4
        ws.page_margins.bottom = 0.4

        # Style
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'),
                             bottom=Side(style='thin'))
        grey_fill = PatternFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid")
        font_small = Font(size=8)
        align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
        align_left_top = Alignment(horizontal='left', vertical='top', wrap_text=True)

        # Szerokości kolumn
        ws.column_dimensions['A'].width = 5
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 18
        ws.column_dimensions['D'].width = 18
        ws.column_dimensions['E'].width = 25
        ws.column_dimensions['F'].width = 15

        # --- NAGŁÓWEK ---
        ws.merge_cells('A1:F1')
        ws['A1'] = f"LISTA OBECNOŚCI {month_str} {year} ROK"
        ws['A1'].font = Font(bold=True, size=14)
        ws['A1'].alignment = Alignment(horizontal='center')

        ws.merge_cells('A2:F4')
        ws['A2'] = f"{emp['name']}\n{emp['job']}"
        ws['A2'].alignment = align_left_top
        ws['A2'].font = Font(bold=True, size=12)

        for r in range(2, 5):
            for c in range(1, 7):
                ws.cell(row=r, column=c).border = thin_border

        # --- TABELA ---
        headers = ["Lp.", "GODZINA\nROZPOCZĘCIA\nPRACY", "GODZINA\nZAKOŃCZENIA\nPRACY", "PODPIS", "UWAGI",
                   "PODPIS OSOBY\nUPOWAŻNIONEJ\nDO KONTROLI"]
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=5, column=col_num, value=header)
            cell.alignment = align_center
            cell.font = Font(bold=True, size=8)
            cell.border = thin_border

        # --- DNI ---
        days_in_month = calendar.monthrange(year, month)[1]

        for day in range(1, 32):
            row = 5 + day
            cell_lp = ws.cell(row=row, column=1)
            cell_lp.alignment = align_center
            cell_lp.border = thin_border

            for col in range(2, 7):
                c = ws.cell(row=row, column=col)
                c.border = thin_border

            if day <= days_in_month:
                cell_lp.value = day
                current_qdate = QDate(year, month, day)
                if current_qdate in holidays_map:
                    for col in range(1, 7):
                        ws.cell(row=row, column=col).fill = grey_fill
            else:
                cell_lp.value = "X"
                for col in range(2, 7):
                    c = ws.cell(row=row, column=col)
                    c.value = "X"
                    c.alignment = align_center

        # --- STOPKA ---
        last_row = 38
        ws.merge_cells(f'A{last_row}:F{last_row}')
        ws[f'A{last_row}'] = "Oznaczenia: N - nieobecność; "
        ws[f'A{last_row}'].font = font_small
        ws[f'A{last_row}'].alignment = Alignment(vertical='bottom')

        # Zapis
        filename = f"{emp['name'].replace(' ', '_')}_{month_str}_{year}.xlsx"
        path = os.path.join(folder, filename)
        wb.save(path)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = GeneratorListObecnosci()
    ex.show()
    sys.exit(app.exec_())