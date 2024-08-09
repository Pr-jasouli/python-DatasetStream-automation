import json
import os
import sys
from tkinter import ttk, Toplevel

import pandas as pd

from utilities.utils import center_window, show_message


class ConfigManager:
    def __init__(self, config_file=None):
        if config_file is None:
            if getattr(sys, 'frozen', False):
                base_dir = os.path.dirname(sys.executable)
            else:
                base_dir = os.path.dirname(os.path.abspath(__file__))
            self.config_file = os.path.join(base_dir, '..', '..', '.config', 'config.json')
        else:
            self.config_file = config_file
        self.config_data = {}

    def load_config(self, file_path=None):
        if file_path is None:
            file_path = self.config_file

        try:
            with open(file_path, 'r') as file:
                content = file.read()
                if content.strip():
                    self.config_data = json.loads(content)
                else:
                    raise ValueError("Empty configuration file")
        except (json.JSONDecodeError, ValueError, FileNotFoundError) as e:
            self.config_data = self.default_config()
            self.save_config()
        return self.config_data

    def save_config(self):
        """Write the configuration data to disk."""
        os.makedirs(os.path.dirname(self.config_file), exist_ok=True)
        with open(self.config_file, 'w') as file:
            json.dump(self.config_data, file, indent=4)

    def default_config(self):
        return {
            'audience_src': '',
            'audience_dest': '',
            'cost_src': '',
            # OTHERS HERE
        }

    def update_config(self, key, value):
        if isinstance(value, str) and '\\' in value:
            value = value.replace('\\', '\\\\')
        self.config_data[key] = value
        self.save_config()

    def get_config(self):
        return self.config_data


class ConfigLoaderPopup(Toplevel):
    def __init__(self, master, config_manager, callback):
        super().__init__(master)
        self.config_manager = config_manager
        self.callback = callback
        self.config_data = self.config_manager.get_config()
        self.loaded_files = set()
        self.selected_files = set()
        self.init_ui()
        center_window(self, master, 800, 400)
        self.transient(master)
        self.grab_set()
        master.wait_window(self)

    def init_ui(self):
        self.title("Load Files")
        self.configure(bg='#f0f0f0')

        files_to_load = {k: v for k, v in self.config_data.items() if os.path.isfile(v)}

        # unique_files = {}
        # for key, path in files_to_load.items():
        #     if path not in unique_files:
        #         unique_files[path] = key

        data = []
        for key, path in files_to_load.items():
            file_name = os.path.basename(path)
            data.append([key, file_name, path])

        df = pd.DataFrame(data, columns=['Config Variable', 'File Name', 'Path'])

        self.main_frame = ttk.Frame(self, style='Main.TFrame')
        self.main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        self.table_frame = ttk.Frame(self.main_frame, style='Main.TFrame')
        self.table_frame.grid(row=0, column=0, sticky="nsew")

        self.tree = ttk.Treeview(self.table_frame, columns=('Config Variable', 'File Name', 'Path'), show='headings', selectmode='extended')
        self.tree.heading('Config Variable', text='Config Variable')
        self.tree.heading('File Name', text='File Name')
        self.tree.heading('Path', text='Path')

        self.tree.column('Config Variable', width=150, stretch=True)
        self.tree.column('File Name', width=150, stretch=True)
        self.tree.column('Path', width=300, stretch=True)

        for idx, row in df.iterrows():
            self.tree.insert('', 'end', values=(row['Config Variable'], row['File Name'], row['Path']))

        self.tree.pack(side='left', fill='both', expand=True)

        scrollbar = ttk.Scrollbar(self.table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side='left', fill='y')

        self.button_frame = ttk.Frame(self.main_frame, style='Main.TFrame')
        self.button_frame.grid(row=0, column=1, sticky="ns")

        close_button = ttk.Button(self.main_frame, text="Next", command=self.load_selected_files)
        close_button.grid(row=1, column=0, columnspan=2, pady=10, sticky="ew")
        close_button.config(width=10)

        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(0, weight=1)

    def load_selected_files(self):
        selected_items = self.tree.selection()
        files_to_load = []
        for item in selected_items:
            values = self.tree.item(item, "values")
            key, path = values[0], values[2]
            if path not in self.selected_files:
                self.selected_files.add(path)
                if path not in self.loaded_files:
                    self.loaded_files.add(path)
                    files_to_load.append((key, path))

        if files_to_load:
            for key, path in files_to_load:
                try:
                    if os.path.exists(path):
                        self.callback(key, path)
                        show_message("File Loaded", f"Successfully loaded {key} from {path}", master=self.master,
                                     custom=True)
                    else:
                        show_message("Error", f"File not found: {path}", type='error', master=self.master, custom=True)
                except Exception as e:
                    show_message("Error", f"Failed to load {key}: {e}", type='error', master=self.master, custom=True)

        self.selected_files.clear()
        self.destroy()
