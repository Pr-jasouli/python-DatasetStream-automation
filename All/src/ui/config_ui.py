import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from functools import partial

from All.src.utils import utils


class ConfigUI(tk.Toplevel):
    def __init__(self, master, tabs):
        super().__init__(master)
        self.title("Configuration")
        self.tabs = tabs
        self.entries = {}
        self.config_manager = master.config_manager
        self.init_ui()

    def init_ui(self):
        """Initializes the user interface components and lays out the window."""
        utils.center_window(self, self.master)
        self.resizable(False, False)
        self.create_menus()
        self.create_frames()
        self.create_frames()
        self.transient(self.master)
        self.grab_set()
        self.wait_window(self)

    def create_menus(self):
        """Creates menu items for the window."""
        menu_items = [
            {"label": "Config", "command": partial(self.toggle_frame, ".config")},
            {"label": "Sources", "command": partial(self.toggle_frame, "sources")}
        ]
        utils.create_menu(self, menu_items)

    def create_frames(self):
        """Initializes frames for configuration settings and source file settings."""
        self.sources_destinations_frame = ttk.Frame(self)
        self.config_frame = ttk.Frame(self)
        self.build_sources_frame()
        self.build_config_frame()
        self.sources_destinations_frame.grid(row=0, column=0, sticky="nsew")
        self.config_frame.grid(row=0, column=0, sticky="nsew")
        self.config_frame.grid_remove()

    def build_sources_frame(self):
        """Creates user interface for source file settings."""
        for idx, tab_name in enumerate(self.tabs):
            self.create_source_widgets(tab_name, idx)
            self.set_initial_path(tab_name)

    def create_source_widgets(self, tab_name, row):
        """Helper function to create label, entry, and button for a source file."""
        ttk.Label(self.sources_destinations_frame, text=f"{tab_name} Reference file:").grid(row=row, column=0, padx=10, pady=5, sticky="ew")
        entry = ttk.Entry(self.sources_destinations_frame)
        entry.grid(row=row, column=1, padx=10, pady=5, sticky="ew")
        ttk.Button(self.sources_destinations_frame, text="Browse", command=lambda e=entry: self.select_file(e)).grid(row=row, column=2, padx=10, pady=5, sticky="ew")
        self.entries[f"{tab_name.lower()}_src"] = entry

    def set_initial_path(self, tab_name):
        """Sets the initial path in the entry field from config."""
        key = f"{tab_name.lower()}_src"
        if key in self.config_manager.config_data:
            self.entries[key].delete(0, tk.END)
            self.entries[key].insert(0, self.config_manager.config_data[key])

    def build_config_frame(self):
        """Creates user interface for configuration management."""
        utils.create_styled_button(self.config_frame, "Load Configuration").pack(pady=10, expand=True, fill="x")
        utils.create_styled_button(self.config_frame, "Save Configuration", command=self.save_configuration).pack(pady=10, expand=True, fill="x")


    def toggle_frame(self, frame_name):
        """Toggles between different frames in the UI based on the user's choice."""
        if frame_name == ".config":
            self.sources_destinations_frame.grid_remove()
            self.config_frame.grid()
        else:
            self.config_frame.grid_remove()
            self.sources_destinations_frame.grid()

    def select_file(self, entry_field):
        """Handles selecting a file and updating the corresponding entry field."""
        file_selected = filedialog.askopenfilename(filetypes=[("All Files", "*.*"), ("Excel Files", "*.xls;*.xlsx"), ("Text Files", "*.txt"), ("CSV Files", "*.csv")])
        if file_selected:
            entry_field.delete(0, tk.END)
            entry_field.insert(0, file_selected)

    def load_configuration(self):
        """Loads the configuration settings."""
        self.config_manager.load_config()
        messagebox.showinfo("Load Configuration", "Configuration loaded successfully!")

    def save_configuration(self):
        """Saves the current configuration settings."""
        for key, entry in self.entries.items():
            cleaned_path = utils.clean_file_path(entry.get())
            self.config_manager.update_config(key, cleaned_path)
        self.config_manager.save_config()
        messagebox.showinfo("Save Configuration", "Configuration saved successfully!")

    def center_window(self):
        """Centers the window on the screen based on the master window's size."""
        window_width, window_height = 420, 650
        x = self.master.winfo_x() + (self.master.winfo_width() - window_width) // 2
        y = self.master.winfo_y() + (self.master.winfo_height() - window_height) // 2
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")

