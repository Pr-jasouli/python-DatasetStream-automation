from datetime import datetime
from tkinter import messagebox
from tkinter import ttk

import pandas as pd

from All.src.utils import utils


class AudienceTab(ttk.Frame):
    def __init__(self, parent, config_data, config_ui_callback=None):
        super().__init__(parent)
        self.config_ui_callback = config_ui_callback
        self.config_data = config_data
        self.df = None
        self.setup_ui()

    def setup_ui(self):
        """Sets up user interface components."""
        self.pack(fill='both', expand=True)
        self.create_styles()
        self.create_widgets()
        self.load_initial_excel()

    def create_styles(self):
        """Configure styles used within the tab."""
        style = ttk.Style(self)
        style.configure('TFrame', background='white')
        style.configure('Title.TLabel', font=('Arial', 12, 'underline'), background='white')
        style.configure('AudienceTab.TButton', padding=[5, 2], font=('Arial', 10))

    def create_widgets(self):
        """Creates the widgets for the Audience tab."""
        ttk.Label(self, text="REFERENCES", style='Title.TLabel').pack(side='top', padx=10, pady=(10, 5))
        self.setup_buttons_and_entries()

    def setup_file_loading(self, parent):
        ttk.Button(parent, text="⇓", command=self.prompt_excel_load, style='AudienceTab.TButton').pack(side='left', padx=10)

    def prompt_excel_load(self):
        filetypes = [("Excel files", "*.xlsx *.xls")]
        utils.select_file(self.load_excel, filetypes)

    def load_excel(self, file_path):
        if file_path:
            try:
                self.df = pd.read_excel(file_path)
                self.update_file_details_label(file_path)
                # AUTO SAVES CONFIG AS YOU LOAD AN EXCEL
                # if self.config_ui_callback:
                #     self.config_ui_callback('audience_src', file_path)
                self.config_data['audience_src'] = file_path
            except Exception as e:
                utils.show_message("Error", f"Failed to load the file:\n{str(e)}", type='error', master=self, custom=True)


    def update_file_details_label(self, file_path):
        rows, cols = self.df.shape
        relative_path = '/'.join(file_path.split('/')[-3:])
        self.file_details_label.config(text=f".../{relative_path} | rows: {rows} ~ columns: {cols}")
        utils.show_message("Success", "Excel file loaded successfully!")

    def setup_buttons_and_entries(self):
        """Setup buttons and entry fields for user interaction."""
        container = ttk.Frame(self)
        container.pack(side='top', fill='x', expand=False, padx=20, pady=10)

        self.setup_file_loading(container)
        self.setup_date_entries(container)
        self.setup_show_columns_button(container)
        self.setup_file_details_display()

    def setup_show_columns_button(self, parent):
        """Sets up a button to show column names from the loaded DataFrame."""
        self.show_columns_button = ttk.Button(parent, text="☱", command=self.show_columns, style='AudienceTab.TButton')
        self.show_columns_button.pack(side='right', padx=10)

    def setup_date_entries(self, parent):
        ttk.Label(parent, text="Date (MM - YYYY):").pack(side='left')
        self.month_entry = ttk.Entry(parent, width=3, validate='key',
                                     validatecommand=(self.register(self.validate_month), '%P'))
        self.month_entry.pack(side='left', padx=(0, 2))
        self.year_entry = ttk.Entry(parent, width=5, validate='key',
                                    validatecommand=(self.register(self.validate_year), '%P'))
        self.year_entry.pack(side='left', padx=(2, 10))
        ttk.Button(parent, text="✓", command=self.validate_date, style='AudienceTab.TButton').pack(side='right', padx=10)

    def setup_file_details_display(self):
        file_details_container = ttk.Frame(self)
        file_details_container.pack(side='top', fill='y', expand=False, padx=20, pady=5)
        self.file_details_label = ttk.Label(file_details_container, text="No file loaded", anchor='w')
        self.file_details_label.pack(side='left', fill='y', expand=False, padx=10)

    def load_initial_excel(self):
        src_audience_path = self.config_data.get('audience_src')
        if src_audience_path:
            self.load_excel(src_audience_path)

    def validate_month(self, P):
        """Validate the month entry to ensure it's empty or a valid month number."""
        if P == "":
            return True
        if P.isdigit() and 1 <= int(P) <= 12:
            return True
        return False

    def validate_year(self, P):
        """Validate the year entry to ensure it meets specified conditions."""
        if P == "":
            return True
        if P.isdigit() and P.startswith("2"):
            if len(P) > 4:
                return False
            if int(P) <= 2044:
                return True
        return False

    def validate_date(self):
        try:
            month = int(self.month_entry.get())
            year = int(self.year_entry.get())
            if datetime(year, month, 1) > datetime.now():
                utils.show_message("Error", "The reference date cannot be in the future.", type='error', master=self, custom=True)
            elif self.df is not None:
                self.check_date_in_data(year, month)
            else:
                utils.show_message("Error", "Load an Excel file first.", type='error', master=self, custom=True)
        except ValueError:
            utils.show_message("Error", "Invalid date. Please enter a valid month and year.", type='error', master=self, custom=True)

    def check_date_in_data(self, year, month):
        """Checks if the date exists in the loaded data and updates the user."""
        mask = (self.df['PERIOD_YEAR'] == year) & (self.df['PERIOD_MONTH'] == month)
        if mask.any():
            utils.show_message("Validation", "Date is valid and found in the file.", type='info', master=self, custom=True)
        else:
            specific_data = self.df[(self.df['PERIOD_YEAR'] == year)]
            utils.show_message("Validation", f"Date not found in the file. Debug: Year({year}), Month({month})\nSample rows where year matches:\n{specific_data.head()}", type='error', master=self, custom=True)

    def show_columns(self):
        """Displays the column names from the loaded DataFrame."""
        if self.df is not None:
            columns = '\n'.join(self.df.columns)
            utils.show_message("Columns", f"Columns in the file:\n{columns}", type='info', master=self, custom=True)
        else:
            utils.show_message("Error", "Load an Excel file first.", type='info', master=self, custom=True)
