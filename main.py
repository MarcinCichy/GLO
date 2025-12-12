import sys
import logging

# Logowanie jak w oryginale (debug.log)
import logging_setup  # noqa: F401

import holidays
# Import konieczny dla pliku EXE
import holidays.countries.poland  # noqa: F401

from PyQt5.QtWidgets import QApplication

from main_window import GeneratorListObecnosci


if __name__ == '__main__':
    try:
        app = QApplication(sys.argv)
        ex = GeneratorListObecnosci()
        ex.show()
        sys.exit(app.exec_())
    except Exception as e:
        logging.critical(f"Krytyczny błąd aplikacji: {e}")
