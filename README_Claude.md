# GLO

Generator Listy Obecności - aplikacja do automatycznego generowania arkuszy Excel z listami obecności dla pracowników.

## Opis

GLO (Generator Listy Obecności) to narzędzie w języku Python służące do automatycznego generowania list obecności w formacie Excel dla pracowników firmy. Aplikacja odczytuje listę pracowników z pliku tekstowego i tworzy dla każdego z nich osobny arkusz Excel zawierający listę obecności na dany miesiąc.

Projekt został zaprojektowany z myślą o uproszczeniu procesu przygotowywania dokumentacji obecności w miejscu pracy. Generowane pliki Excel zawierają predefiniowaną strukturę z datami, dniami tygodnia oraz miejscem na oznaczanie obecności pracowników.

**Główne funkcjonalności:**
- Odczyt listy pracowników z pliku tekstowego
- Automatyczne generowanie arkuszy Excel dla każdego pracownika
- Tworzenie list obecności na wybrany miesiąc i rok
- Wypełnianie dat i dni tygodnia automatycznie
- Obsługa formatowania i stylizacji arkuszy

## Zrzuty ekranu

![Screenshot](glo/screenshot/main.png)

## Struktura projektu

```
GLO/
├── .idea/
│   ├── .gitignore
│   ├── GLO.iml
│   ├── inspectionProfiles/
│   │   ├── Project_Default.xml
│   │   └── profiles_settings.xml
│   ├── misc.xml
│   ├── modules.xml
│   ├── vcs.xml
│   └── workspace.xml
├── glo/
│   ├── .gitignore
│   ├── Anna_Nowak_GRUDZIEŃ_2025.xlsx
│   ├── Jan_Kowalski_GRUDZIEŃ_2025.xlsx
│   ├── Ktos_Tam_GRUDZIEŃ_2025.xlsx
│   ├── brief.md
│   ├── etapy.md
│   ├── main.py
│   ├── plan.md
│   ├── pracownicy.txt
│   └── screenshot/
│       └── main.png
└── README.md
```

## Instalacja

1. Sklonuj repozytorium:
```bash
git clone <repository-url>
cd GLO
```

2. Zainstaluj wymagane zależności:
```bash
pip install openpyxl
```

## Konfiguracja

Aby skonfigurować aplikację:

1. Utwórz plik `pracownicy.txt` w katalogu `glo/` zawierający listę pracowników (jeden pracownik w każdej linii):
```
Jan Kowalski
Anna Nowak
Ktos Tam
```

2. Dostosuj parametry w pliku `main.py` (miesiąc, rok) według potrzeb.

## Uruchomienie

Aby uruchomić aplikację:

```bash
cd glo
python main.py
```

Po uruchomieniu aplikacji, w katalogu `glo/` zostaną wygenerowane pliki Excel z listami obecności dla każdego pracownika z pliku `pracownicy.txt`.

## Zależności

Projekt wymaga następujących bibliotek Python:

- **openpyxl** - do tworzenia i edycji plików Excel
- **Python 3.x** - podstawowe środowisko uruchomieniowe

---

# GLO

Attendance List Generator - an application for automatically generating Excel spreadsheets with attendance lists for employees.

## Description

GLO (Generator Listy Obecności / Attendance List Generator) is a Python tool designed to automatically generate attendance lists in Excel format for company employees. The application reads a list of employees from a text file and creates a separate Excel spreadsheet for each employee containing an attendance list for a given month.

The project was designed to simplify the process of preparing attendance documentation in the workplace. Generated Excel files contain a predefined structure with dates, days of the week, and space for marking employee attendance.

**Main features:**
- Reading employee list from a text file
- Automatic generation of Excel spreadsheets for each employee
- Creating attendance lists for a selected month and year
- Automatic filling of dates and days of the week
- Support for sheet formatting and styling

## Screenshots

![Screenshot](glo/screenshot/main.png)

## Project structure

```
GLO/
├── .idea/
│   ├── .gitignore
│   ├── GLO.iml
│   ├── inspectionProfiles/
│   │   ├── Project_Default.xml
│   │   └── profiles_settings.xml
│   ├── misc.xml
│   ├── modules.xml
│   ├── vcs.xml
│   └── workspace.xml
├── glo/
│   ├── .gitignore
│   ├── Anna_Nowak_GRUDZIEŃ_2025.xlsx
│   ├── Jan_Kowalski_GRUDZIEŃ_2025.xlsx
│   ├── Ktos_Tam_GRUDZIEŃ_2025.xlsx
│   ├── brief.md
│   ├── etapy.md
│   ├── main.py
│   ├── plan.md
│   ├── pracownicy.txt
│   └── screenshot/
│       └── main.png
└── README.md
```

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd GLO
```

2. Install required dependencies:
```bash
pip install openpyxl
```

## Configuration

To configure the application:

1. Create a `pracownicy.txt` file in the `glo/` directory containing the list of employees (one employee per line):
```
Jan Kowalski
Anna Nowak
Ktos Tam
```

2. Adjust parameters in the `main.py` file (month, year) as needed.

## Usage

To run the application:

```bash
cd glo
python main.py
```

After running the application, Excel files with attendance lists will be generated in the `glo/` directory for each employee from the `pracownicy.txt` file.

## Dependencies

The project requires the following Python libraries:

- **openpyxl** - for creating and editing Excel files
- **Python 3.x** - basic runtime environment