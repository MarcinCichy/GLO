import holidays
# Import konieczny dla pliku EXE
import holidays.countries.poland

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import QCalendarWidget


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
