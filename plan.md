# PLAN PRACY (TASK LIST) - STATUS: UKOŃCZONY (Wersja 10.0 Final)

## 1. Setup projektu
- [x] Stworzenie struktury plików projektu.
- [x] Instalacja bibliotek: `PyQt5`, `openpyxl`, `holidays`, `pyinstaller`, `pywin32`.

## 2. GUI i Logika Kalendarza
- [x] Utworzenie Layoutu: Panel górny (Data), Lewy (Pracownicy), Prawy (Kalendarz), Dolny (Akcje).
- [x] Oprogramowanie `QComboBox` (Miesiąc/Rok) z synchronizacją.
- [x] Dodanie `QCalendarWidget` z własnym malowaniem (święta, weekendy).
- [x] Synchronizacja dwukierunkowa (Combo <-> Kalendarz + Rolka myszy).

## 3. Zarządzanie Danymi
- [x] Lista pracowników z Checkboxami (`QListWidget`).
- [x] Funkcja "Zaznacz/Odznacz wszystkich" (Toggle).
- [x] Edytor tekstowy dla szybkiego wklejania listy.
- [x] Automatyczne wczytywanie `pracownicy.txt` przy starcie.
- [x] Obsługa błędów (brak pliku, tworzenie domyślnego).

## 4. Generator Excel - Strona 1 (Lista Obecności)
- [x] Konfiguracja `openpyxl` (style, czcionki Cambria/Times New Roman).
- [x] Ustawienie strony A4 (marginesy, skalowanie `fitToPage`, centrowanie poziome).
- [x] Kalibracja wymiarów:
    - [x] Wysokości wierszy (nagłówek 1.80 cm).
    - [x] Szerokość kolumn (Kolumna F poszerzona do 14.5).
- [x] Formatowanie Nagłówka F5 ("PODPIS OSOBY..."):
    - [x] Orientacja pozioma (kąt 0).
    - [x] Wycentrowanie w pionie i poziomie.
    - [x] Czcionka zmniejszona do 7 pkt (bez pogrubienia).
- [x] Wyśrodkowanie danych pracownika (Imię/Stanowisko).
- [x] Usunięcie zbędnych ramek przy podpisach.

## 5. Generator Excel - Strona 2 (Ewidencja)
- [x] Utworzenie arkusza "Ewidencja" w tym samym pliku.
- [x] Skalowanie wydruku (wymuszenie 1 strony szerokości/wysokości).
- [x] Tabela z 18 kolumnami (A-R):
    - [x] Wiersz 1: Scalone nagłówki grupowe.
    - [x] Wiersz 2: Pionowe nagłówki (czcionki 8 i 7 pkt).
- [x] **NOWOŚĆ:** Automatyczne obliczanie normy czasu pracy (komórka R2).
    - [x] Algorytm uwzględniający: dni robocze, święta ustawowe ORAZ dni zaznaczone przez użytkownika.
    - [x] Wynik w formacie "w dniach: X / w godzinach: Y".
- [x] Stopka i Legenda:
    - [x] Pełna treść legendy w scalonej komórce.
    - [x] Podpisy Kadr i Dyrektora (poprawiono literówkę "ODPIS" -> "PODPIS").
- [x] Logika drukowania: Zaznaczenie obu arkuszy (`tabSelected = True`).

## 6. Moduł Drukowania i UX
- [x] Integracja z systemem Windows (`win32api`, `win32print`).
- [x] Okno wyboru drukarki (`PrinterDialog`) z przyciskiem ustawień Duplex.
- [x] Dynamiczna kolejka wydruku (drukowanie tylko zaznaczonych osób).
- [x] **NOWOŚĆ:** Przycisk "Otwórz folder" (aktywny po wybraniu ścieżki).
- [x] **NOWOŚĆ:** "Ludzkie" komunikaty błędów (np. informacja o otwartym pliku Excel).

## 7. Architektura Premium (Back-end)
- [x] **Wielowątkowość:** Przeniesienie generowania do `WorkerThread` (interfejs nie zamarza).
- [x] **Progress Bar:** Pasek postępu informujący o stanie pracy.
- [x] **Logowanie:** Zapis błędów do pliku `debug.log`.
- [x] **Konfiguracja:** Zapamiętywanie ostatniego folderu w `config.ini`.

## 8. Finalizacja (Deployment)
- [x] Rozwiązanie problemu z `ModuleNotFoundError` (jawny import `holidays.countries.poland`).
- [x] Kompilacja do jednego pliku `.exe` (PyInstaller `--onefile --noconsole`).