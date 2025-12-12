from PyQt5.QtCore import Qt, QRectF
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import QStyledItemDelegate


class RotatedHeaderDelegate(QStyledItemDelegate):
    """
    Rysuje pionowy tekst w tabeli podglądu.
    Naprawiono błąd: Nie obraca tekstu w stopce (wiersze > 33).
    """

    def paint(self, painter, option, index):
        # Jeśli to wiersz stopki (indeks > 32, czyli wiersz 34+), rysuj normalnie
        if index.row() > 32:
            QStyledItemDelegate.paint(self, painter, option, index)
            return

        text = index.data(Qt.DisplayRole)
        bg_brush = index.data(Qt.BackgroundRole)

        painter.save()

        # 1. Tło
        if bg_brush:
            painter.fillRect(option.rect, bg_brush)

        # 2. Ramka
        painter.setPen(QColor("#dcdcdc"))
        painter.drawRect(option.rect)

        # 3. Tekst
        if text:
            painter.setFont(option.font)
            painter.setPen(Qt.black)

            # Translacja środka i rotacja -90
            rect = option.rect
            painter.translate(rect.center())
            painter.rotate(-90)

            # Po obrocie zamieniamy wymiary (W, H -> H, W)
            text_rect = QRectF(-rect.height() / 2, -rect.width() / 2, rect.height(), rect.width())
            painter.drawText(text_rect, Qt.AlignCenter | Qt.TextWordWrap, text)

        painter.restore()
