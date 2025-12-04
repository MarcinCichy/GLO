# BRIEF APLIKACJI: Generator List Obecności (GLO)

## 1. Cel biznesowy
Stworzenie desktopowej aplikacji okienkowej automatyzującej proces tworzenia miesięcznych list obecności i kart ewidencji czasu pracy. Aplikacja ma drastycznie skrócić czas potrzebny na ręczne formatowanie tabel w Excelu oraz eliminować błędy w oznaczaniu dni wolnych.

## 2. Użytkownik docelowy
Pracownicy administracji/kadr (np. w żłobku), odpowiedzialni za przygotowywanie dokumentacji pracowniczej.

## 3. Wymagania Techniczne
* **Język programowania:** Python.
* **Interfejs (GUI):** PyQt5 (zalecane) lub TKinter.
* **Format wyjściowy:** Pliki `.xlsx` (Excel), gotowe do druku (A4).
* **Dystrybucja:** Plik wykonywalny `.exe` (Windows), działający bez instalacji środowiska Python.

## 4. Kluczowe Funkcjonalności

### A. Konfiguracja Czasu i Dni Wolnych
* **Wybór okresu:** Użytkownik wybiera Miesiąc i Rok.
* **Automatyczne dni wolne:** Program automatycznie rozpoznaje weekendy (soboty, niedziele) oraz święta ustawowe w Polsce (np. za pomocą biblioteki `holidays`).
* **Niestandardowe dni wolne (NOWE):** Użytkownik ma możliwość kliknięcia w interaktywny kalendarz, aby oznaczyć dodatkowe dni wolne w danym miesiącu (np. dni dyrektorskie, dodatkowe święta zakładowe). Te dni będą traktowane przez generator tak samo jak niedziele (zacienione).

### B. Baza Pracowników
* **Wprowadzanie danych:** Pole tekstowe lub lista do wprowadzania pracowników w formacie "Imię Nazwisko, Stanowisko".
* **Import:** Możliwość wklejenia listy z innego źródła.

### C. Generowanie Dokumentów (Silnik Excel)
* **Zgodność wizualna:** Wygenerowany plik musi wizualnie odpowiadać dostarczonym wzorom (`strona_1.png` i `strona_2.png`).
* **Struktura pliku:** Jeden plik Excel na pracownika, zawierający dwa arkusze (lub dwie strony do druku):
    1.  **Lista Obecności:** Tabela wejść/wyjść, dane nagłówkowe.
    2.  **Karta Ewidencji:** Szczegółowa tabela godzinowa, obrócone nagłówki, stopka ze stałym podpisem "PODPIS DYREKTORA ŻŁOBKA".
* **Logika wizualna:** Wiersze odpowiadające dniom wolnym (weekendy + święta + dni wybrane ręcznie) muszą mieć szare tło (zacienienie).

## 5. Wygląd Aplikacji (GUI)
Jedno okno główne podzielone na sekcje:
1.  **Sekcja Daty:** Wybór Roku/Miesiąca + Kalendarz do "odklikiwania" dodatkowych dni wolnych.
2.  **Sekcja Pracowników:** Pole do wprowadzania/edycji listy osób.
3.  **Sekcja Akcji:** Przycisk "Generuj", pasek postępu, wybór folderu zapisu.