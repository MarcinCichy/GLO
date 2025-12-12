import logging

from PyQt5.QtCore import QThread, pyqtSignal

from excel_generator import ExcelGenerator


class WorkerThread(QThread):
    progress_updated = pyqtSignal(int)
    finished = pyqtSignal(dict, str)
    error_occurred = pyqtSignal(str)

    def __init__(self, employees, folder, month, year, month_str, holidays_map):
        super().__init__()
        self.employees = employees
        self.folder = folder
        self.month = month
        self.year = year
        self.month_str = month_str
        self.holidays_map = holidays_map
        self.generator = ExcelGenerator()

    def run(self):
        generated_map = {}
        total = len(self.employees)
        try:
            for i, emp in enumerate(self.employees):
                # Jeśli folder nie jest zdefiniowany (np. użytkownik anulował), wątek się nie uda
                if not self.folder:
                    raise Exception("Nie wybrano folderu zapisu.")

                filepath = self.generator.create_file(
                    emp, self.folder, self.month, self.year, self.month_str, self.holidays_map
                )
                generated_map[emp['key']] = filepath
                progress_percent = int(((i + 1) / total) * 100)
                self.progress_updated.emit(progress_percent)

            self.finished.emit(generated_map, self.folder)

        except PermissionError as e:
            clean_msg = (f"Nie można zapisać pliku dla pracownika: {emp['name']}.\n\n"
                         f"PRAWDOPODOBNA PRZYCZYNA:\n"
                         f"Plik jest otwarty w Excelu/OpenOffice.\n\n"
                         f"Zamknij plik i spróbuj ponownie.")
            logging.error(f"PermissionError: {e}")
            self.error_occurred.emit(clean_msg)

        except Exception as e:
            logging.error(f"Critical Error: {e}")
            self.error_occurred.emit(f"Wystąpił nieoczekiwany błąd:\n{str(e)}")
