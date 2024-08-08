import tkinter as tk
from tkinter import ttk

from ui.config_ui import ConfigUI
from ui.audience_tab import AudienceTab
from ui.cost_tab import CostTab
from utilities import config_manager
from utilities.utils import show_message, create_styled_button
from utilities.config_manager import ConfigManager
import os



class ConfigLoaderPopup(tk.Toplevel):
    def __init__(self, master, config_manager, callback):
        super().__init__(master)
        self.config_manager = config_manager
        self.callback = callback
        self.config_data = self.config_manager.get_config()
        self.loaded_files = set()
        self.init_ui()
        center_window(self, master, 800, 400)
        self.transient(master)
        self.grab_set()
        master.wait_window(self)

    def init_ui(self):
        self.title("Load Files")
        self.configure(bg='#f0f0f0')

        files_to_load = {k: v for k, v in self.config_data.items() if os.path.isfile(v)}

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

        load_button = ttk.Button(self.button_frame, text="Load", command=self.load_selected_files)
        load_button.grid(row=0, column=0, padx=5, pady=2, sticky="n")
        load_button.config(width=7)

        close_button = ttk.Button(self.main_frame, text="Next", command=self.destroy)
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
            if path not in self.loaded_files:
                self.loaded_files.add(path)
                files_to_load.append((key, path))

        if not files_to_load:
            show_message("Info", "No new files selected to load.", master=self.master, custom=True)
            return

        for key, path in files_to_load:
            try:
                if os.path.exists(path):
                    self.callback(key, path)
                    show_message("File Loaded", f"Successfully loaded {key} from {path}", master=self.master, custom=True)
                else:
                    show_message("Error", f"File not found: {path}", type='error', master=self.master, custom=True)
            except Exception as e:
                show_message("Error", f"Failed to load {key}: {e}", type='error', master=self.master, custom=True)




class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()
        self.config_manager = ConfigManager()
        self.config_data = self.config_manager.load_config()
        self.config_ui_callback = self.update_config_data

        self.initialize_ui()

    def initialize_ui(self):
        """Initialize the main UI components."""
        self.title("Telecom Application")
        self.create_styles()
        self.configure_geometry()
        self.create_menus()
        self.create_tabs()
        self.create_bottom_frame()

    def configure_geometry(self):
        """Configure window size and properties."""
        self.geometry("970x730")
        self.minsize(600, 450)

    def create_styles(self):
        """Configure the styles for the application."""
        style = ttk.Style(self)
        style.theme_use('xpnative')
        style.configure('TButton', font=('Helvetica', 14), padding=10)
        style.configure('TLabel', font=('Helvetica', 12), background='#f0f0f0')
        style.configure('TNotebook.Tab', font=('Helvetica', 8), padding=[20, 8],
                        background='white', borderwidth=1, relief='solid')
        style.configure('Bottom.TFrame', background='SystemButtonFace')

    def create_menus(self):
        """Set up the application's menu bar."""
        menubar = tk.Menu(self)
        self.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Open", command=self.file_open)
        file_menu.add_command(label="Save Configuration", command=self.save_configuration)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.exit_app)

        edit_menu = tk.Menu(menubar, tearoff=0)
        edit_menu.add_command(label="Undo", command=self.edit_undo)
        edit_menu.add_command(label="Redo", command=self.edit_redo)
        edit_menu.add_separator()
        edit_menu.add_command(label="Preferences", command=self.open_config)

        menubar.add_cascade(label="File", menu=file_menu)
        menubar.add_cascade(label="Edit", menu=edit_menu)

    def setup_file_menu(self):
        """Create and return the file menu."""
        file_menu = tk.Menu(self.menubar, tearoff=0, background='SystemButtonFace', fg='black')
        file_menu.add_command(label="Open", command=lambda: show_message("Open", "Open a file!", type="info", master=self, custom=True))
        file_menu.add_command(label="Save Configuration", command=config_manager.save_config)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.exit_app)
        return file_menu

    def setup_edit_menu(self):
        """Create and return the edit menu."""
        edit_menu = tk.Menu(self.menubar, tearoff=0, background='SystemButtonFace', fg='black')
        edit_menu.add_command(label="Undo", command=lambda: show_message("Undo", "Undo the last action!", type="info", master=self, custom=True))
        edit_menu.add_command(label="Redo", command=lambda: show_message("Redo", "Redo the last undone action!", type="info", master=self, custom=True))
        edit_menu.add_separator()
        edit_menu.add_command(label="Preferences", command=self.open_config)
        return edit_menu

    def create_tabs(self):
        """Initialize tab controls and tabs."""
        self.tab_control = ttk.Notebook(self, padding=10)
        self.audience_tab = AudienceTab(parent=self.tab_control, config_manager=self.config_manager,
                                        config_ui_callback=self.config_ui_callback)
        self.cost_tab = CostTab(self.tab_control, config_manager=self.config_manager, config_ui_callback=self.config_ui_callback)
        self.tab_control.add(self.audience_tab, text='Audience')
        self.tab_control.add(self.cost_tab, text='  Cost  ')
        self.tab_control.pack(expand=1, fill='both', padx=15, pady=(5, 0))

    def create_bottom_frame(self):
        """Create the bottom frame with configuration button."""
        bottom_frame = ttk.Frame(self, padding=10, style='Bottom.TFrame')
        bottom_frame.pack(fill='x', side='bottom')
        create_styled_button(bottom_frame, '\U0001F527', self.open_config).pack(side='left', padx=10)

    def open_config(self):
        """Open the configuration UI."""
        tab_names = [self.tab_control.tab(tab, "text") for tab in self.tab_control.tabs()]
        self.config_ui = ConfigUI(self, tab_names)

    def file_open(self):
        show_message("Open", "Open a file!", type="info", master=self, custom=True)

    def save_configuration(self):
        """Saves the current configuration using the ConfigManager."""
        try:
            self.config_manager.save_config()
            show_message("Configuration", "Configuration saved successfully!", type="info", master=self, custom=True)
        except Exception as e:
            show_message("Error", "Failed to save configuration:\n" + str(e), type="error", master=self, custom=True)

    def exit_app(self):
        self.quit()

    def edit_undo(self):
        show_message("Undo", "Undo the last action!", type="info", master=self, custom=True)

    def edit_redo(self):
        show_message("Open", "Open a file!", type="info", master=self, custom=True)

    def update_config_data(self, key, value):
        self.config_manager.update_config(key, value)


if __name__ == "__main__":
    app = MainApplication()
    app.mainloop()
