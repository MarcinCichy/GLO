# PLAN PRACY (TASK LIST)

## 1. Setup projektu
- [x] Stworzenie struktury plików projektu.
- [x] Instalacja bibliotek: `PyQt5`, `openpyxl`, `holidays`, `pyinstaller`.

## 2. GUI i Logika Kalendarza
- [x] Utworzenie Layoutu: Panel górny (Data), Lewy (Pracownicy), Prawy (Kalendarz), Dolny (Akcje).
- [x] Oprogramowanie `QComboBox` (Miesiąc/Rok).
- [x] Dodanie `QCalendarWidget` z własnym malowaniem (`paintCell`).
    - [x] Zablokowanie systemowego zaznaczania (niebieska ramka).
    - [x] Synchronizacja widoku z ComboBoxami.
    - [x] Obsługa `clicked`: Dodawanie/usuwanie dni z listy customowej.

## 3. Zarządzanie Danymi
- [x] Pole edycji tekstu dla pracowników.
- [x] Funkcje Zapisz/Wczytaj listę do pliku `pracownicy.txt`.
- [x] Funkcja `calculate_holidays(year, month)`:
    - [x] Pobieranie świąt z `holidays.PL`.
    - [x] Pobieranie weekendów.
    - [x] Uwzględnianie dni zaznaczonych ręcznie w kalendarzu.

## 4. Generator Excel - Strona 1 (Lista Obecności)
- [ ] Konfiguracja `openpyxl` (style, czcionki, obramowania).
- [ ] Ustawienie strony A4 (marginesy, skalowanie).
- [ ] Generowanie nagłówka (Miesiąc/Rok) i danych pracownika.
- [ ] Generowanie tabeli dni (1-31).
- [ ] **Logika:** Jeśli dzień jest na liście `holidays` -> Szare tło (PatternFill).
- [ ] Obsługa dni nieistniejących (np. 30 luty) -> Wstawienie "X".

## 5. Generator Excel - Strona 2 (Ewidencja)
- [ ] Ustawienie orientacji poziomej (Landscape).
- [ ] Obracanie nagłówków o 90 stopni.
- [ ] Stopki z podpisami ("PODPIS DYREKTORA ŻŁOBKA").
- [ ] Zacienianie dni wolnych.

## 6. Finalizacja
- [ ] Integracja paska postępu.
- [ ] Wybór folderu zapisu.
- [ ] Kompilacja do `.exe`.