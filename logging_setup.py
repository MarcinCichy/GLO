import logging

# === KONFIGURACJA LOGOWANIA BŁĘDÓW ===
logging.basicConfig(
    filename='debug.log',
    level=logging.ERROR,
    format='%(asctime)s:%(levelname)s:%(message)s'
)
