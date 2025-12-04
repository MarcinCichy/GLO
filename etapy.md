# ETAPY TWORZENIA APLIKACJI GLO

## Etap 1: Prototyp Interfejsu (GUI) i Synchronizacja
Celem jest stworzenie interfejsu, który reaguje na działania użytkownika.
* Stworzenie głównego okna aplikacji.
* Implementacja selektorów Daty (Miesiąc/Rok).
* Implementacja interaktywnego kalendarza (`QCalendarWidget`).
* **Kluczowa logika:** Zmiana daty w selektorze automatycznie zmienia widok kalendarza na wybrany miesiąc.
* Implementacja mechanizmu "klikania" w dni na kalendarzu (dodawanie/usuwanie dat z listy "własne wolne" i zmiana koloru tła dnia).

## Etap 2: Zarządzanie Danymi (Pracownicy i Czas)
Celem jest obsługa danych wejściowych przed generowaniem raportu.
* Obsługa listy pracowników: wpisywanie ręczne + **funkcja zapisu/odczytu listy z pliku** (by dane nie ginęły po wyłączeniu programu).
* Integracja z biblioteką `holidays` (pobieranie świąt państwowych).
* Logika scalania dat: stworzenie jednej listy dni wolnych, która łączy: [Weekendy] + [Święta PL] + [Dni zaznaczone ręcznie w kalendarzu].

## Etap 3: Silnik Raportujący (Excel - Strona 1)
Odwzorowanie "Listy Obecności".
* Konfiguracja `openpyxl`.
* Ustawienie parametrów strony A4 (marginesy, skalowanie).
* Generowanie nagłówka i tabeli wejść/wyjść.
* Implementacja warunkowego formatowania: jeśli dzień jest na "liście scalonej" z Etapu 2 -> zaciemnij wiersz (szary kolor).

## Etap 4: Silnik Raportujący (Excel - Strona 2)
Odwzorowanie "Karty Ewidencji".
* Budowa skomplikowanej tabeli z obróconymi nagłówkami.
* Precyzyjne ustawienie szerokości kolumn (tak, by tabela nie była zbyt wąska ani nie wychodziła poza stronę).
* Dodanie stałych stopek z podpisami (zgodnie ze wzorem).
* Zastosowanie tego samego mechanizmu zaciemniania dni wolnych co na Stronie 1.

## Etap 5: Integracja, Budowanie i Testy
* Połączenie przycisku "Generuj" z logiką Excela.
* Dodanie paska postępu (Progress Bar) i wyboru folderu zapisu.
* Kompilacja do pliku `.exe`.
* **Weryfikacja:** Wydrukowanie wygenerowanych stron i przyłożenie ich do wzorca (pliki png) w celu sprawdzenia marginesów.