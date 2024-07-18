import tkinter as tk
from tkinter import ttk

from All.src.utils import utils, config_manager
from All.src.utils.config_manager import ConfigManager
from All.src.ui.audience_tab import AudienceTab
from All.src.ui import config_ui


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
        self.geometry("1066x800")
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
        file_menu.add_command(label="Open", command=lambda: utils.show_message("Open", "Open a file!", type="info", master=self, custom=True))
        file_menu.add_command(label="Save Configuration", command=config_manager.save_config)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.exit_app)
        return file_menu

    def setup_edit_menu(self):
        """Create and return the edit menu."""
        edit_menu = tk.Menu(self.menubar, tearoff=0, background='SystemButtonFace', fg='black')
        edit_menu.add_command(label="Undo", command=lambda: utils.show_message("Undo", "Undo the last action!"))
        edit_menu.add_command(label="Redo", command=lambda: utils.show_message("Redo", "Redo the last undone action!"))
        edit_menu.add_command(label="Undo", command=lambda: utils.show_message("Undo", "Undo the last action!", type="info", master=self, custom=True))
        edit_menu.add_command(label="Redo", command=lambda: utils.show_message("Redo", "Redo the last undone action!", type="info", master=self, custom=True))
        edit_menu.add_separator()
        edit_menu.add_command(label="Preferences", command=self.open_config)
        return edit_menu

    def create_tabs(self):
        """Initialize tab controls and tabs."""
        self.tab_control = ttk.Notebook(self, padding=10)
        self.audience_tab = AudienceTab(parent=self.tab_control, config_data=self.config_data,
                                        config_ui_callback=self.config_ui_callback)
        # self.pdf_tab = PDFTab(self.tab_control, config_data=self.config_data,
        #                       config_ui_callback=self.config_ui_callback)
        self.tab_control.add(self.audience_tab, text='Audience')
        # self.tab_control.add(self.pdf_tab, text='PDF')
        self.tab_control.pack(expand=1, fill='both', padx=15, pady=15)

    def create_bottom_frame(self):
        """Create the bottom frame with configuration button."""
        bottom_frame = ttk.Frame(self, padding=10, style='Bottom.TFrame')
        bottom_frame.pack(fill='x', side='bottom')
        utils.create_styled_button(bottom_frame, '\U0001F527', self.open_config).pack(side='left', padx=10)

    def open_config(self):
        """Open the configuration UI."""
        tab_names = [self.tab_control.tab(tab, "text") for tab in self.tab_control.tabs()]
        self.config_ui = config_ui.ConfigUI(self, tab_names)

    def file_open(self):
        utils.show_message("Open", "Open a file!", type="info", master=self, custom=True)

    def save_configuration(self):
        """Saves the current configuration using the ConfigManager."""
        try:
            self.config_manager.save_config()
            utils.show_message("Configuration", "Configuration saved successfully!", type="info", master=self, custom=True)
        except Exception as e:
            utils.show_message("Error", "Failed to save configuration:\n" + str(e), type="error", master=self, custom=True)

    def exit_app(self):
        self.quit()

    def edit_undo(self):
        utils.show_message("Undo", "Undo the last action!", type="info", master=self, custom=True)

    def edit_redo(self):
        utils.show_message("Open", "Open a file!", type="info", master=self, custom=True)

    def update_config_data(self, key, value):
        self.config_manager.update_config(key, value)


if __name__ == "__main__":
    app = MainApplication()
    app.mainloop()
