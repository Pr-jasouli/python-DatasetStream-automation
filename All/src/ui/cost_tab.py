import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd
import tkinter.font as tkFont

from utilities.config_manager import ConfigManager


class CostTab(ttk.Frame):
    def __init__(self, parent, config_manager=None, config_ui_callback=None):
        super().__init__(parent)
        self.config_manager = config_manager
        self.config_ui_callback = config_ui_callback
        self.file_path = self.config_manager.config_data.get("cost_src", "")
        self.data = None
        self.metadata = None

        self.column_mapping = {
            'Business provider': 'NETWORK_NAME',
            'contract name': 'CNT_NAME_GRP',
            'contract start date': 'CT_STARTDATE',
            'contract_end_date': 'CT_ENDDATE',
            'contract effective date': 'CT_EFFECTIVE_DATE',
            'contract duration': 'CT_DURATION',
            'contract type': 'CT_TYPE',
            'contract book year': 'CT_BOOK_YEAR',
            'contract auto reniew': 'CT_AUTORENEW',
            'contract notice period': 'CT_NOTICE_PER',
            'ollo': ['CT_AVAIL_IN_SCARLET_FR', 'CT_AVAIL_IN_SCARLET_NL'],
            'Fix_fee- Fee': 'CT_FIXFEE',
            'Fix_fee_new- New Fee': 'CT_FIXFEE_NEW'
        }

        self.network_name_var = tk.StringVar()
        self.cnt_name_grp_var = tk.StringVar()
        self.business_model_var = tk.StringVar()

        self.init_ui()
        if self.file_path:
            self.load_cost_reference_file(self.file_path)

    def init_ui(self):
        style = ttk.Style()
        style.configure("Custom.TButton", font=('Helvetica', 10))

        top_frame = ttk.Frame(self)
        top_frame.grid(row=0, column=0, pady=10, padx=10, sticky="ew")

        tree_container = ttk.Frame(self)
        tree_container.grid(row=1, column=0, pady=10, padx=10, sticky="nsew")

        ttk.Label(top_frame, text="Network", font=('Helvetica', 10)).grid(row=0, column=0, pady=5, padx=5, sticky="e")
        self.network_name_dropdown = ttk.Combobox(top_frame, textvariable=self.network_name_var, state="readonly", font=('Helvetica', 10))
        self.network_name_dropdown.grid(row=0, column=1, pady=5, padx=5)

        ttk.Label(top_frame, text="Contract:", font=('Helvetica', 10)).grid(row=0, column=2, pady=5, padx=5, sticky="e")
        self.cnt_name_grp_dropdown = ttk.Combobox(top_frame, textvariable=self.cnt_name_grp_var, state="readonly", font=('Helvetica', 10))
        self.cnt_name_grp_dropdown.grid(row=0, column=3, pady=5, padx=5)

        ttk.Label(top_frame, text="Model:", font=('Helvetica', 10)).grid(row=0, column=4, pady=5, padx=5, sticky="e")
        self.business_model_dropdown = ttk.Combobox(top_frame, textvariable=self.business_model_var, state="readonly", font=('Helvetica', 10))
        self.business_model_dropdown.grid(row=0, column=5, pady=5, padx=5)

        self.load_cost_button = ttk.Button(top_frame, text="Load Cost File", command=self.load_cost_data, width=20, style="Custom.TButton")
        self.load_cost_button.grid(row=1, column=2, columnspan=3, pady=10, padx=5)

        self.tree = ttk.Treeview(tree_container, columns=(), show="headings")
        self.tree.grid(row=0, column=0, sticky="nsew")

        tree_xscroll = ttk.Scrollbar(tree_container, orient="horizontal", command=self.tree.xview)
        tree_xscroll.grid(row=1, column=0, sticky="ew")
        self.tree.configure(xscrollcommand=tree_xscroll.set)

        tree_container.columnconfigure(0, weight=1)
        tree_container.rowconfigure(0, weight=1)

        self.network_name_dropdown.bind("<<ComboboxSelected>>", self.update_dropdowns)
        self.cnt_name_grp_dropdown.bind("<<ComboboxSelected>>", self.update_dropdowns)
        self.business_model_dropdown.bind("<<ComboboxSelected>>", self.update_dropdowns)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

    def load_cost_reference_file(self, file_path):
        self.data = self.find_relevant_sheet(file_path)
        self.metadata = pd.read_excel(file_path, sheet_name="BUSINESS CONTRACT METADATA")
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
