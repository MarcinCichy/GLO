# PLAN PRACY (TASK LIST) - STATUS: UKOŃCZONY

## 1. Setup projektu
- [x] Stworzenie struktury plików projektu.
- [x] Instalacja bibliotek: `PyQt5`, `openpyxl`, `holidays`, `pyinstaller`, `pywin32`.

## 2. GUI i Logika Kalendarza
- [x] Utworzenie Layoutu: Panel górny (Data), Lewy (Pracownicy), Prawy (Kalendarz), Dolny (Akcje).
- [x] Oprogramowanie `QComboBox` (Miesiąc/Rok) z synchronizacją.
- [x] Dodanie `QCalendarWidget`:
    - [x] Własne malowanie (święta, weekendy, dni customowe).
    - [x] Usunięcie paska nawigacji (zmiana tylko rolką/combo).
    - [x] Synchronizacja dwukierunkowa (Combo <-> Kalendarz).

## 3. Zarządzanie Danymi
- [x] Lista pracowników z Checkboxami (`QListWidget`).
- [x] Funkcja "Zaznacz/Odznacz wszystkich" (Toggle).
- [x] Edytor tekstowy dla szybkiego wklejania listy.
- [x] Automatyczne wczytywanie `pracownicy.txt` przy starcie.
- [x] Obsługa błędów (brak pliku, tworzenie domyślnego).

## 4. Generator Excel - Strona 1 (Lista Obecności)
- [x] Konfiguracja `openpyxl` (style, czcionki Cambria/Times New Roman).
- [x] Ustawienie strony A4 (marginesy, skalowanie `fitToPage`).
- [x] Kalibracja wymiarów:
    - [x] Wysokości wierszy (nagłówek 1.80 cm).
    - [x] Szerokość kolumn (Kolumna F poszerzona do 14.5, aby literka "J" nie spadała).
- [x] Formatowanie:
    - [x] Obrót tekstu 90 stopni bez zawijania.
    - [x] Wyśrodkowanie danych pracownika.
    - [x] Usunięcie zbędnych ramek przy podpisach.

## 5. Generator Excel - Strona 2 (Ewidencja)
- [x] Utworzenie arkusza "Ewidencja" w tym samym pliku.
- [x] Skalowanie wydruku (wymuszenie 1 strony szerokości/wysokości).
- [x] Tabela z 18 kolumnami (A-R):
    - [x] Wiersz 1: Scalone nagłówki grupowe.
    - [x] Wiersz 2: Pionowe nagłówki (czcionki 8 i 7 pkt).
- [x] Stopka i Legenda:
    - [x] Pełna treść legendy w scalonej komórce (poprawiono błąd z brakującą zmienną).
    - [x] Podpisy Kadr i Dyrektora (dopasowane do strony).
- [x] Logika drukowania: Zaznaczenie obu arkuszy (`tabSelected = True`).

## 6. Moduł Drukowania i UX
- [x] Integracja z systemem Windows (`win32api`, `win32print`).
- [x] Okno wyboru drukarki (`PrinterDialog`).
- [x] Przycisk otwierający ustawienia systemowe drukarki (sprawdzenie Duplexu).
- [x] Dynamiczny przycisk druku:
    - [x] Aktualizacja licznika plików po zmianie zaznaczenia na liście.
    - [x] Drukowanie tylko zaznaczonych osób (mapowanie nazwisk na pliki).

## 7. Finalizacja (Deployment)
- [x] Przygotowanie instrukcji kompilacji.
- [x] Rozwiązanie problemu z `ModuleNotFoundError: holidays.countries` (dodanie jawnego importu).
- [x] Kompilacja do jednego pliku `.exe` (PyInstaller `--onefile --noconsole`).
- [x] Testy działania pliku wykonywalnego na "czysto".