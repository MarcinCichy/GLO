import os
import configparser


class ConfigStore:
    def __init__(self, path: str = "config.ini"):
        self.path = path
        self.config = configparser.ConfigParser()
        self.last_save_folder = ""
        self.load()

    def load(self) -> None:
        if os.path.exists(self.path):
            try:
                self.config.read(self.path)
                self.last_save_folder = self.config.get('SETTINGS', 'LastFolder', fallback="")
            except Exception:
                pass

    def save(self) -> None:
        if 'SETTINGS' not in self.config:
            self.config['SETTINGS'] = {}
        self.config['SETTINGS']['LastFolder'] = self.last_save_folder
        with open(self.path, 'w') as configfile:
            self.config.write(configfile)
