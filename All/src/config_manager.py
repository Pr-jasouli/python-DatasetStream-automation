import json
import os
from tkinter import messagebox


class ConfigManager:
    def __init__(self, config_file='config.json'):
        self.config_file = config_file
        self.config_data = {}

    def load_config(self):
        if os.path.exists(self.config_file):
            with open(self.config_file, 'r') as file:
                self.config_data = json.load(file)
        else:
            self.config_data = self.default_config()
            self.save_config()
            messagebox.showinfo("Configuration", "No config.json found. A new empty configuration file has been created.")
        return self.config_data

    def save_config(self):
        with open(self.config_file, 'w') as file:
            json.dump(self.config_data, file, indent=4)

    def default_config(self):
        return {
            'audience_src': '',
            #OTHERS HERE
        }

    def update_config(self, key, value):
        self.config_data[key] = value
        self.save_config()

    def get_config(self):
        return self.config_data
