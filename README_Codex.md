# GLO

## Opis

GLO to system zarządzania danymi pracowników i generowania raportów w formacie Excel. Aplikacja umożliwia automatyczne tworzenie plików `.xlsx` zawierających dane poszczególnych pracowników wraz z podsumowaniami miesięcznymi.

Główne funkcjonalności:
- Zarządzanie listą pracowników
- Generowanie miesięcznych raportów dla każdego pracownika
- Eksport danych do formatu Excel (.xlsx)
- Automatyczne formatowanie i strukturyzacja danych

Aplikacja została zaprojektowana jako narzędzie konsolowe w języku Python, które pozwala na szybkie przetwarzanie danych pracowniczych i tworzenie uporządkowanych raportów.

## Zrzuty ekranu

![Screenshot](glo/screenshot/main.png)

## Struktura projektu

```
GLO/
├── .idea/                                      # Konfiguracja PyCharm IDE
│   ├── .gitignore
│   ├── GLO.iml
│   ├── inspectionProfiles/
│   │   ├── Project_Default.xml
│   │   └── profiles_settings.xml
│   ├── misc.xml
│   ├── modules.xml
│   ├── vcs.xml
│   └── workspace.xml
├── glo/                                        # Główny katalog aplikacji
│   ├── .gitignore
│   ├── main.py                                # Główny plik aplikacji
│   ├── pracownicy.txt                         # Lista pracowników
│   ├── brief.md                               # Dokumentacja projektowa
│   ├── etapy.md                               # Plan etapów rozwoju
│   ├── plan.md                                # Plan projektu
│   ├── screenshot/                            # Zrzuty ekranu
│   │   └── main.png
│   ├── Anna_Nowak_GRUDZIEŃ_2025.xlsx         # Przykładowy raport
│   ├── Jan_Kowalski_GRUDZIEŃ_2025.xlsx       # Przykładowy raport
│   └── Ktos_Tam_GRUDZIEŃ_2025.xlsx           # Przykładowy raport
├── README.md                                   # Dokumentacja projektu
├── README.old.md                              # Archiwalne wersje README
├── README.old.20251208_231220.md
├── README.old.20251209_185259.md
└── README_Claude.md
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

**Uwaga:** Projekt wymaga Python 3.6 lub nowszego.

## Uruchomienie

Uruchom aplikację za pomocą polecenia:

```bash
cd glo
python main.py
```

Aplikacja wczyta listę pracowników z pliku `pracownicy.txt` i wygeneruje dla każdego z nich plik Excel z danymi za bieżący miesiąc.

## Zależności

Projekt wykorzystuje następujące biblioteki Python:

- **openpyxl** - do obsługi plików Excel (.xlsx)
- **datetime** - do obsługi dat (biblioteka standardowa)
- **pathlib** - do zarządzania ścieżkami plików (biblioteka standardowa)

---

# GLO

## Description

GLO is an employee data management and Excel report generation system. The application enables automatic creation of `.xlsx` files containing individual employee data with monthly summaries.

Main features:
- Employee list management
- Monthly report generation for each employee
- Data export to Excel format (.xlsx)
- Automatic data formatting and structuring

The application is designed as a Python console tool that allows quick processing of employee data and creation of organized reports.

## Screenshots

![Screenshot](glo/screenshot/main.png)

## Project structure

```
GLO/
├── .idea/                                      # PyCharm IDE configuration
│   ├── .gitignore
│   ├── GLO.iml
│   ├── inspectionProfiles/
│   │   ├── Project_Default.xml
│   │   └── profiles_settings.xml
│   ├── misc.xml
│   ├── modules.xml
│   ├── vcs.xml
│   └── workspace.xml
├── glo/                                        # Main application directory
│   ├── .gitignore
│   ├── main.py                                # Main application file
│   ├── pracownicy.txt                         # Employee list
│   ├── brief.md                               # Project documentation
│   ├── etapy.md                               # Development stages plan
│   ├── plan.md                                # Project plan
│   ├── screenshot/                            # Screenshots
│   │   └── main.png
│   ├── Anna_Nowak_GRUDZIEŃ_2025.xlsx         # Sample report
│   ├── Jan_Kowalski_GRUDZIEŃ_2025.xlsx       # Sample report
│   └── Ktos_Tam_GRUDZIEŃ_2025.xlsx           # Sample report
├── README.md                                   # Project documentation
├── README.old.md                              # Archived README versions
├── README.old.20251208_231220.md
├── README.old.20251209_185259.md
└── README_Claude.md
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

**Note:** The project requires Python 3.6 or newer.

## Usage

Run the application using the command:

```bash
cd glo
python main.py
```

The application will read the employee list from the `pracownicy.txt` file and generate an Excel file with data for the current month for each employee.

## Dependencies

The project uses the following Python libraries:

- **openpyxl** - for Excel file (.xlsx) handling
- **datetime** - for date handling (standard library)
- **pathlib** - for file path management (standard library)