import tkinter as tk
from tkinter import ttk


from ui.tab_audience import AudienceTab
from ui.tab_cost import CostTab
from ui.ui_config import ConfigUI
from utilities.utils import show_message, create_styled_button
from utilities.config_manager import ConfigManager, ConfigLoaderPopup
import os


class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()
        self.config_manager = ConfigManager()
        self.config_data = self.config_manager.load_config()
        self.config_ui_callback = self.update_config_data

        self.initialize_ui()
        self.center_window()
        self.show_file_loader_popup_if_files_exist()

    def center_window(self):
        window_width = 970
        window_height = 780
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        position_top = int(screen_height / 2 - window_height / 2)
        position_right = int(screen_width / 2 - window_width / 2)
        self.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')

    def show_file_loader_popup_if_files_exist(self):
        files_to_load = {k: v for k, v in self.config_data.items() if os.path.isfile(v)}
        if files_to_load:
            print("Files found, showing loader popup.")
            ConfigLoaderPopup(self, self.config_manager, self.load_tab_content)


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
        self.minsize(600, 450)

    def create_styles(self):
        """Configure the styles for the application."""
        style = ttk.Style(self)
        style.theme_use('xpnative')
        style.configure('TButton', font=('Helvetica', 14), padding=10)
        style.configure('TLabel', font=('Helvetica', 12), background='#f0f0f0')
        style.configure('TNotebook.Tab', font=('Helvetica', 8), padding=[20, 8],
                         borderwidth=1, relief='solid')
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
        file_menu.add_command(label="Open",
                              command=lambda: show_message("Open", "Open a file!", type="info", master=self,
                                                           custom=True))
        file_menu.add_command(label="Save Configuration", command=self.save_configuration)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.exit_app)
        return file_menu

    def setup_edit_menu(self):
        """Create and return the edit menu."""
        edit_menu = tk.Menu(self.menubar, tearoff=0, background='SystemButtonFace', fg='black')
        edit_menu.add_command(label="Undo",
                              command=lambda: show_message("Undo", "Undo the last action!", type="info", master=self,
                                                           custom=True))
        edit_menu.add_command(label="Redo",
                              command=lambda: show_message("Redo", "Redo the last undone action!", type="info",
                                                           master=self, custom=True))
        edit_menu.add_separator()
        edit_menu.add_command(label="Preferences", command=self.open_config)
        return edit_menu

    def create_tabs(self):
        """Initialize tab controls and tabs."""
        self.tab_control = ttk.Notebook(self, padding=10)

        # Create tabs without loading files
        self.audience_tab = AudienceTab(parent=self.tab_control, config_manager=self.config_manager,
                                        config_ui_callback=self.config_ui_callback)
        self.cost_tab = CostTab(self.tab_control, config_manager=self.config_manager,
                                config_ui_callback=self.config_ui_callback)

        self.tab_control.add(self.audience_tab, text='Audience')
        self.tab_control.add(self.cost_tab, text='Cost')

        self.tab_control.pack(expand=1, fill='both', padx=15, pady=(5, 0))

    def load_tab_content(self, key, path):
        """Callback function to load content into specific tabs."""
        try:
            if key == "audience_src":
                self.audience_tab.load_file(path)
            elif key == "cost_src":
                self.cost_tab.load_file(path)
        except Exception as e:
            show_message("Error", f"Failed to load {key}: {e}", type='error', master=self, custom=True)

    def create_bottom_frame(self):
        """Create the bottom frame with configuration and process buttons."""
        bottom_frame = ttk.Frame(self, padding=10, style='Bottom.TFrame')
        bottom_frame.pack(side='bottom', fill='x')

        create_styled_button(bottom_frame, '\U0001F527', self.open_config).pack(side='left', padx=10)

        audience_frame = ttk.Labelframe(bottom_frame, text="Audience", padding=5)
        audience_frame.pack(side='left', padx=10, pady=(0, 15))

        process_button = create_styled_button(audience_frame, "Process", self.audience_tab.start_processing,
                                              width=12)
        process_button.pack(side='left', padx=5, pady=5)

        view_result_button = create_styled_button(audience_frame, "View", self.audience_tab.view_result,
                                                  width=12)
        view_result_button.pack(side='left', padx=5, pady=5)

        cost_frame = ttk.Labelframe(bottom_frame, text="Cost", padding=5)
        cost_frame.pack(side='left', padx=10, pady=(0, 15))

        process_button = create_styled_button(cost_frame, "New Deal", self.cost_tab.open_new_deal_popup,
                                              width=12)
        process_button.pack(side='left', padx=5, pady=5)

        update_deal_button = create_styled_button(cost_frame, "Update Deal", self.cost_tab.open_update_deal_popup, width=12)
        update_deal_button.pack(side='left', padx=5, pady=5)



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
        show_message("Redo", "Redo the last undone action!", type="info", master=self, custom=True)

    def update_config_data(self, key, value):
        self.config_manager.update_config(key, value)


if __name__ == "__main__":
    app = MainApplication()
    app.mainloop()
