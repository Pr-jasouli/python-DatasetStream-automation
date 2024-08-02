import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd

from utilities.config_manager import ConfigManager


class CostTab(ttk.Frame):
    def __init__(self, parent, config_manager=None, config_ui_callback=None):
        super().__init__(parent)
        self.config_manager = config_manager
        self.config_ui_callback = config_ui_callback
        self.file_path = self.config_manager.config_data.get("cost_src", "")
        self.data = None

        self.year_var = tk.StringVar()
        self.network_name_var = tk.StringVar()
        self.prod_en_name_var = tk.StringVar()

        self.init_ui()
        if self.file_path:
            self.load_cost_reference_file(self.file_path)

    def init_ui(self):
        """Initialize the UI components."""
        style = ttk.Style()
        style.configure("Custom.TButton", font=('Helvetica', 10))

        container = ttk.Frame(self)
        container.grid(row=0, column=0, pady=10, padx=10)

        ttk.Label(container, text="CT_BOOK_YEAR:", font=('Helvetica', 10)).grid(row=0, column=0, pady=5, padx=5, sticky="e")
        self.year_dropdown = ttk.Combobox(container, textvariable=self.year_var, state="readonly", font=('Helvetica', 10), width=10)
        self.year_dropdown.grid(row=0, column=1, pady=5, padx=5)

        ttk.Label(container, text="NETWORK_NAME:", font=('Helvetica', 10)).grid(row=0, column=2, pady=5, padx=5, sticky="e")
        self.network_name_dropdown = ttk.Combobox(container, textvariable=self.network_name_var, state="readonly", font=('Helvetica', 10))
        self.network_name_dropdown.grid(row=0, column=3, pady=5, padx=5)

        ttk.Label(container, text="CNT_NAME_GRP:", font=('Helvetica', 10)).grid(row=0, column=4, pady=5, padx=5, sticky="e")
        self.prod_en_name_dropdown = ttk.Combobox(container, textvariable=self.prod_en_name_var, state="readonly", font=('Helvetica', 10))
        self.prod_en_name_dropdown.grid(row=0, column=5, pady=5, padx=5)

        self.load_cost_button = ttk.Button(container, text="Load Cost File", command=self.load_cost_data, width=20, style="Custom.TButton")
        self.load_cost_button.grid(row=1, column=3, columnspan=3, pady=10, padx=5)

        self.get_contracts_button = ttk.Button(container, text="Get Contracts (TODO)", command=self.get_contracts, width=20, style="Custom.TButton")
        self.get_contracts_button.grid(row=1, column=0, columnspan=3, pady=10, padx=5)

        self.year_dropdown.bind("<<ComboboxSelected>>", self.update_dropdowns)
        self.network_name_dropdown.bind("<<ComboboxSelected>>", self.update_dropdowns)
        self.prod_en_name_dropdown.bind("<<ComboboxSelected>>", self.update_dropdowns)

    def load_cost_reference_file(self, file_path):
        """Load the cost reference file and initialize dropdown menus."""
        self.data = self.find_relevant_sheet(file_path)
        if self.data is not None:
            self.populate_dropdowns()

    def find_relevant_sheet(self, file_path):
        """Find the sheet with the required columns and return the data."""
        xls = pd.ExcelFile(file_path)
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            if {'CT_BOOK_YEAR', 'NETWORK_NAME', 'CNT_NAME_GRP'}.issubset(df.columns):
                return df
        return None

    def populate_dropdowns(self):
        """Populate the initial values in the dropdown menus."""
        self.data['CT_BOOK_YEAR'] = self.data['CT_BOOK_YEAR'].astype(str)
        self.data['NETWORK_NAME'] = self.data['NETWORK_NAME'].astype(str)
        self.data['CNT_NAME_GRP'] = self.data['CNT_NAME_GRP'].astype(str)

        years = [''] + sorted(self.data['CT_BOOK_YEAR'].dropna().unique())
        network_names = [''] + sorted(self.data['NETWORK_NAME'].dropna().unique())
        prod_en_names = [''] + sorted(self.data['CNT_NAME_GRP'].dropna().unique())

        self.year_dropdown['values'] = years
        self.network_name_dropdown['values'] = network_names
        self.prod_en_name_dropdown['values'] = prod_en_names

    def update_dropdowns(self, event=None):
        """Update the dropdown values based on selections."""
        year_selected = self.year_var.get()
        network_name_selected = self.network_name_var.get()
        prod_en_name_selected = self.prod_en_name_var.get()

        # Filter current selections
        filtered_data = self.data.copy()
        if year_selected:
            filtered_data = filtered_data[filtered_data['CT_BOOK_YEAR'] == year_selected]
        if network_name_selected:
            filtered_data = filtered_data[filtered_data['NETWORK_NAME'] == network_name_selected]
        if prod_en_name_selected:
            filtered_data = filtered_data[filtered_data['CNT_NAME_GRP'] == prod_en_name_selected]

        # Update year dropdown
        years = [''] + sorted(filtered_data['CT_BOOK_YEAR'].dropna().unique())
        current_year = self.year_var.get()
        self.year_dropdown['values'] = years
        self.year_var.set(current_year if current_year in years else '')

        # Update network name dropdown
        network_names = [''] + sorted(filtered_data['NETWORK_NAME'].dropna().unique())
        current_network_name = self.network_name_var.get()
        self.network_name_dropdown['values'] = network_names
        self.network_name_var.set(current_network_name if current_network_name in network_names else '')

        # Update product name dropdown
        prod_en_names = [''] + sorted(filtered_data['CNT_NAME_GRP'].dropna().unique())
        current_prod_en_name = self.prod_en_name_var.get()
        self.prod_en_name_dropdown['values'] = prod_en_names
        self.prod_en_name_var.set(current_prod_en_name if current_prod_en_name in prod_en_names else '')

        if not year_selected:
            self.network_name_var.set('')
            self.prod_en_name_var.set('')
            self.network_name_dropdown['values'] = [''] + sorted(self.data['NETWORK_NAME'].dropna().unique())
            self.prod_en_name_dropdown['values'] = [''] + sorted(self.data['CNT_NAME_GRP'].dropna().unique())

        if not network_name_selected:
            self.prod_en_name_var.set('')
            filtered_data = self.data.copy()
            if year_selected:
                filtered_data = filtered_data[filtered_data['CT_BOOK_YEAR'] == year_selected]
            self.prod_en_name_dropdown['values'] = [''] + sorted(filtered_data['CNT_NAME_GRP'].dropna().unique())

    def load_cost_data(self):
        """Handle the Load Cost button click."""
        file_path = filedialog.askopenfilename(
            title="Select Cost File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.file_path = file_path
            self.load_cost_reference_file(file_path)

    def get_contracts(self):
        """Handle the Get Contracts button click."""
        year_selected = self.year_var.get()
        network_name_selected = self.network_name_var.get()
        prod_en_name_selected = self.prod_en_name_var.get()
        print(f"Selected Year: {year_selected}")
        print(f"Selected Network Name: {network_name_selected}")
        print(f"Selected Product Name: {prod_en_name_selected}")


if __name__ == "__main__":
    root = tk.Tk()
    config_manager = ConfigManager()
    tab = CostTab(root, config_manager=config_manager)
    tab.pack(expand=1, fill='both')
    root.mainloop()
