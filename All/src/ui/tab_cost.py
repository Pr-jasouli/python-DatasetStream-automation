import os
import tkinter as tk
import tkinter.font as tkFont
from tkinter import ttk, filedialog

import pandas as pd

from utilities import utils
from utilities.config_manager import ConfigManager
from utilities.utils import show_message, open_file_and_update_config


class CostTab(ttk.Frame):
    def __init__(self, parent, base_dir, config_manager=None, config_ui_callback=None):
        super().__init__(parent)
        self.model_columns = self.get_model_columns()
        self.config_manager = config_manager
        self.config_ui_callback = config_ui_callback
        self.config_data = config_manager.get_config()
        self.file_path = self.config_data.get('cost_src', None)
        self.data = None
        self.network_name_var = tk.StringVar()
        self.cnt_name_grp_var = tk.StringVar()
        self.business_model_var = tk.StringVar()
        self.allocation_var = tk.StringVar()

        self.base_dir = base_dir

        self.init_ui()

    def generate_template(self, new_row):
        working_contracts_file = os.path.join(self.config_data.get('cost_dest', ''), 'working_contracts.xlsx')

        business_model = new_row['Business model'].capitalize()

        if not os.path.exists(working_contracts_file):
            with pd.ExcelWriter(working_contracts_file, engine='openpyxl') as writer:
                new_df = pd.DataFrame([new_row])
                new_df.to_excel(writer, sheet_name=business_model, index=False)
                print(f"Debug: Created new working contracts file with sheet '{business_model}'.")

        else:
            with pd.ExcelWriter(working_contracts_file, engine='openpyxl', mode='a',
                                if_sheet_exists='overlay') as writer:
                try:
                    existing_df = pd.read_excel(working_contracts_file, sheet_name=business_model)
                    updated_df = pd.concat([existing_df, pd.DataFrame([new_row])], ignore_index=True)
                except ValueError:
                    updated_df = pd.DataFrame([new_row])

                updated_df.to_excel(writer, sheet_name=business_model, index=False)
                print(f"Debug: Updated '{business_model}' sheet in the working contracts file.")

        show_message("Success", f"Contract added to {working_contracts_file} under '{business_model}' sheet.",
                     master=self, custom=True)

    def get_model_columns(self):
        return {
            ### OK 100%
            'Fixed fee': ['CT_STARTDATE', 'CT_ENDDATE', 'allocation', 'NETWORK_NAME', 'CNT_NAME_GRP', 'PROD_EN_NAME', 'CT_TYPE',
                          'CT_DURATION', 'CT_NOTICE_DATE', 'CT_AUTORENEW', 'CT_NOTICE_PER',
                          'CT_AVAIL_IN_SCARLET_FR',
                          'CT_AVAIL_IN_SCARLET_NL', 'CT_FIXFEE', 'CT_FIXFEE_NEW', 'Business model', 'variable/fix'],

            'fixed fee': ['CT_BOOK_YEAR', 'allocation', 'NETWORK_NAME', 'CNT_NAME_GRP', 'PROD_EN_NAME', 'CT_TYPE', 'CT_STARTDATE', 'CT_ENDDATE',
                          'CT_DURATION', 'CT_NOTICE_DATE', 'CT_AUTORENEW', 'CT_NOTICE_PER',
                          'CT_AVAIL_IN_SCARLET_FR',
                          'CT_AVAIL_IN_SCARLET_NL', 'CT_FIXFEE', 'CT_FIXFEE_NEW', 'Business model', 'variable/fix'],

            'Free': ['CT_BOOK_YEAR', 'allocation', 'NETWORK_NAME', 'CNT_NAME_GRP', 'CT_TYPE', 'CT_STARTDATE', 'CT_ENDDATE',
                     'CT_DURATION', 'CT_NOTICE_DATE', 'CT_AUTORENEW', 'CT_NOTICE_PER',
                     'CT_AVAIL_IN_SCARLET_FR',
                     'CT_AVAIL_IN_SCARLET_NL', 'Business model', 'variable/fix'],
            'free': ['CT_BOOK_YEAR', 'allocation', 'NETWORK_NAME', 'CNT_NAME_GRP', 'CT_TYPE', 'CT_STARTDATE', 'CT_ENDDATE',
                     'CT_DURATION', 'CT_NOTICE_DATE', 'CT_AUTORENEW', 'CT_NOTICE_PER',
                     'CT_AVAIL_IN_SCARLET_FR',
                     'CT_AVAIL_IN_SCARLET_NL', 'Business model', 'variable/fix'],

            'fixed fee + index': [
                'CT_STARTDATE', 'CT_ENDDATE', 'allocation', 'NETWORK_NAME', 'CNT_NAME_GRP', 'PROD_EN_NAME',
                'CT_TYPE', 'CT_DURATION', 'CT_NOTICE_DATE', 'CT_AUTORENEW', 'CT_NOTICE_PER',
                'CT_AVAIL_IN_SCARLET_FR', 'CT_AVAIL_IN_SCARLET_NL', 'CT_FIXFEE', 'CT_FIXFEE_NEW',
                'Business model', 'variable/fix', 'CT_INDEX'
            ],
            'Fixed fee cogs': [
                'CT_STARTDATE', 'CT_ENDDATE', 'allocation', 'NETWORK_NAME', 'CNT_NAME_GRP', 'PROD_EN_NAME',
                'CT_TYPE', 'CT_DURATION', 'CT_NOTICE_DATE', 'CT_AUTORENEW', 'CT_NOTICE_PER',
                'CT_AVAIL_IN_SCARLET_FR', 'CT_AVAIL_IN_SCARLET_NL', 'CT_FIXFEE', 'CT_FIXFEE_NEW',
                'Business model', 'variable/fix'
            ],
            'fixed fee cogs': [
                'CT_STARTDATE', 'CT_ENDDATE', 'allocation', 'NETWORK_NAME', 'CNT_NAME_GRP', 'PROD_EN_NAME',
                'CT_TYPE', 'CT_DURATION', 'CT_NOTICE_DATE', 'CT_AUTORENEW', 'CT_NOTICE_PER',
                'CT_AVAIL_IN_SCARLET_FR', 'CT_AVAIL_IN_SCARLET_NL', 'CT_FIXFEE', 'CT_FIXFEE_NEW',
                'Business model', 'variable/fix'
            ],
            'CPS Over MG Subs': [
                'CT_STARTDATE', 'CT_ENDDATE', 'allocation', 'NETWORK_NAME', 'CNT_NAME_GRP', 'PROD_EN_NAME',
                'CT_TYPE', 'CT_DURATION', 'CT_NOTICE_DATE', 'CT_AUTORENEW', 'CT_NOTICE_PER',
                'CT_AVAIL_IN_SCARLET_FR', 'CT_AVAIL_IN_SCARLET_NL', 'CT_INDEX', 'CT_MG', 'CT_VARFEE', 'CT_VARFEE_NEW', 'CT_STEP1_SUBS', 'CT_STEP2_SUBS', 'CT_STEP3_SUBS'
            ],
            'CPS Over MG Subs + index': [
                'CT_STARTDATE', 'CT_ENDDATE', 'allocation', 'NETWORK_NAME', 'CNT_NAME_GRP', 'PROD_EN_NAME',
                'CT_TYPE', 'CT_DURATION', 'CT_NOTICE_DATE', 'CT_AUTORENEW', 'CT_NOTICE_PER',
                'CT_AVAIL_IN_SCARLET_FR', 'CT_AVAIL_IN_SCARLET_NL', 'CT_INDEX', 'CT_MG', 'CT_VARFEE', 'CT_VARFEE_NEW', 'CT_STEP1_SUBS', 'CT_STEP2_SUBS', 'CT_STEP3_SUBS'
            ],
            'Revenue share over MG subs': [
                'CT_STARTDATE', 'CT_ENDDATE', 'allocation', 'NETWORK_NAME', 'CNT_NAME_GRP', 'PROD_EN_NAME',
                'CT_TYPE', 'CT_DURATION', 'CT_NOTICE_DATE', 'CT_AUTORENEW', 'CT_NOTICE_PER',
                'CT_AVAIL_IN_SCARLET_FR', 'CT_AVAIL_IN_SCARLET_NL', 'Revenue share %', 'CT_STEP1_SUBS',
                'CT_STEP2_SUBS', 'CT_STEP3_SUBS'
            ],
            'CPS on product park': [
                'CT_STARTDATE', 'CT_ENDDATE', 'allocation', 'NETWORK_NAME', 'CNT_NAME_GRP', 'PROD_EN_NAME',
                'CT_TYPE', 'CT_DURATION', 'CT_NOTICE_DATE', 'CT_AUTORENEW', 'CT_NOTICE_PER',
                'CT_AVAIL_IN_SCARLET_FR', 'CT_AVAIL_IN_SCARLET_NL', 'PROD_PRICE', 'CT_STEP1_SUBS', 'CT_STEP2_SUBS', 'CT_STEP3_SUBS' #, '#Subscriber', 'PROD_PRICE_VAT_EXCL'
            ],
            'CPS on volume regionals + index': [
                'CT_STARTDATE', 'CT_ENDDATE', 'allocation', 'NETWORK_NAME', 'CNT_NAME_GRP', 'PROD_EN_NAME',
                'CT_TYPE', 'CT_DURATION', 'CT_NOTICE_DATE', 'CT_AUTORENEW', 'CT_NOTICE_PER',
                'CT_AVAIL_IN_SCARLET_FR', 'CT_AVAIL_IN_SCARLET_NL', 'CT_INDEX', '#Subscriber',
                'CT_STEP1_SUBS', 'CT_STEP2_SUBS', 'CT_STEP3_SUBS'
            ],
            'Event Based INTEC': [
                'CT_STARTDATE', 'CT_ENDDATE', 'allocation', 'NETWORK_NAME', 'CNT_NAME_GRP', 'PROD_EN_NAME',
                'CT_TYPE', 'CT_DURATION', 'CT_NOTICE_DATE', 'CT_AUTORENEW', 'CT_NOTICE_PER',
                'CT_AVAIL_IN_SCARLET_FR', 'CT_AVAIL_IN_SCARLET_NL', 'CT_VARFEE', 'CT_VARFEE_NEW'
            ],
        }

    def load_file(self, path):
        self.file_path = path
        self.load_cost_reference_file(path)

    def load_cost_reference_file(self, file_path):
        try:
            self.data = pd.read_excel(file_path, sheet_name='all contract cost file')
            self.populate_dropdowns()

            self.network_name_dropdown.config(state='normal')
            self.cnt_name_grp_dropdown.config(state='normal')
            self.business_model_dropdown.config(state='normal')
            self.allocation_dropdown.config(state='normal')
        except Exception as e:
            show_message("Error", f"Failed to load cost file: {e}", type='error', master=self, custom=True)

    def load_cost_data(self):
        file_path = open_file_and_update_config(
            config_manager=self.config_manager,
            config_key='cost_src',
            title="Select Cost File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if isinstance(file_path, str) and file_path:
            self.load_file(file_path)
        else:
            print("Invalid file path:", file_path)

    def init_ui(self):
        style = ttk.Style()
        style.configure("Custom.TButton", font=('Helvetica', 10))

        top_frame = ttk.Frame(self)
        top_frame.grid(row=0, column=0, pady=10, padx=10, sticky="ew")

        tree_container = ttk.Frame(self)
        tree_container.grid(row=1, column=0, pady=10, padx=10, sticky="nsew")

        self.load_cost_button = ttk.Button(
            top_frame, text="Cost File", command=self.load_cost_data, width=9, style='AudienceTab.TButton'
        )
        self.load_cost_button.grid(row=0, column=0, pady=5, padx=4, sticky="w")
        self.load_cost_button.bind("<Enter>", lambda e: self.show_tooltip(
            e,
            "Business model\nsheet: all contract cost file\ncolumns: 'NETWORK_NAME', 'CNT_NAME_GRP', 'Business model'")
                                   )
        self.load_cost_button.bind("<Leave>", lambda e: self.hide_tooltip())

        self.refresh_fields_button = ttk.Button(
            top_frame, text="Refresh", command=self.clear_fields, width=9, style='AudienceTab.TButton'
        )
        self.refresh_fields_button.grid(row=1, column=0, pady=5, padx=4, sticky="w")

        ttk.Label(top_frame, text="Network:", font=('Helvetica', 10)).grid(
            row=0, column=3, pady=5, padx=4, sticky="e"
        )
        self.network_name_dropdown = ttk.Combobox(
            top_frame, textvariable=self.network_name_var, state="disabled", font=('Helvetica', 10)
        )
        self.network_name_dropdown.grid(row=0, column=4, pady=5, padx=4)

        ttk.Label(top_frame, text="Channel:", font=('Helvetica', 10)).grid(
            row=0, column=5, pady=5, padx=4, sticky="e"
        )
        self.cnt_name_grp_dropdown = ttk.Combobox(
            top_frame, textvariable=self.cnt_name_grp_var, state="disabled", font=('Helvetica', 10)
        )
        self.cnt_name_grp_dropdown.grid(row=0, column=6, pady=5, padx=4)

        ttk.Label(top_frame, text="Model:", font=('Helvetica', 10)).grid(
            row=0, column=1, pady=5, padx=4, sticky="e"
        )
        self.business_model_dropdown = ttk.Combobox(
            top_frame, textvariable=self.business_model_var, state="disabled", font=('Helvetica', 10)
        )
        self.business_model_dropdown.grid(row=0, column=2, pady=5, padx=4)

        ttk.Label(top_frame, text="Allocation:", font=('Helvetica', 10)).grid(
            row=1, column=1, pady=5, padx=4, sticky="e"
        )
        self.allocation_var = tk.StringVar()
        self.allocation_dropdown = ttk.Combobox(
            top_frame, textvariable=self.allocation_var, state="disabled", font=('Helvetica', 10)
        )
        self.allocation_dropdown.grid(row=1, column=2, pady=5, padx=4)


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
        self.allocation_dropdown.bind("<<ComboboxSelected>>", self.update_dropdowns)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

    def populate_dropdowns(self):
        if self.data is None or self.data.empty:
            self.network_name_dropdown['values'] = ['']
            self.cnt_name_grp_dropdown['values'] = ['']
            self.business_model_dropdown['values'] = ['']
            self.allocation_dropdown['values'] = ['']

            self.network_name_dropdown.config(state='disabled')
            self.cnt_name_grp_dropdown.config(state='disabled')
            self.business_model_dropdown.config(state='disabled')
            self.allocation_dropdown.config(state='disabled')
        else:
            self.network_name_dropdown.config(state='normal')
            self.cnt_name_grp_dropdown.config(state='normal')
            self.business_model_dropdown.config(state='normal')
            self.allocation_dropdown.config(state='normal')

            self.data['NETWORK_NAME'] = self.data['NETWORK_NAME'].astype(str)
            self.data['CNT_NAME_GRP'] = self.data['CNT_NAME_GRP'].astype(str)
            self.data['Business model'] = self.data['Business model'].astype(str)
            self.data['allocation'] = self.data['allocation'].astype(str)

            network_names = [''] + sorted(self.data['NETWORK_NAME'].dropna().unique())
            cnt_name_grps = [''] + sorted(self.data['CNT_NAME_GRP'].dropna().unique())
            business_models = [''] + sorted(self.data['Business model'].dropna().unique())
            allocations = [''] + sorted(self.data['allocation'].dropna().unique())

            self.network_name_dropdown['values'] = network_names
            self.cnt_name_grp_dropdown['values'] = cnt_name_grps
            self.business_model_dropdown['values'] = business_models
            self.allocation_dropdown['values'] = allocations

    def update_dropdowns(self, event=None):
        if self.data.empty:
            return

        network_name_selected = self.network_name_var.get()
        cnt_name_grp_selected = self.cnt_name_grp_var.get()
        business_model_selected = self.business_model_var.get()
        allocation_selected = self.allocation_var.get()

        filtered_data = self.data.copy()
        if network_name_selected:
            filtered_data = filtered_data[filtered_data['NETWORK_NAME'] == network_name_selected]
        if cnt_name_grp_selected:
            filtered_data = filtered_data[filtered_data['CNT_NAME_GRP'] == cnt_name_grp_selected]
        if business_model_selected:
            filtered_data = filtered_data[filtered_data['Business model'] == business_model_selected]
        if allocation_selected:
            filtered_data = filtered_data[filtered_data['allocation'] == allocation_selected]

        network_names = [''] + sorted(filtered_data['NETWORK_NAME'].dropna().unique())
        current_network_name = self.network_name_var.get()
        self.network_name_dropdown['values'] = network_names
        self.network_name_var.set(current_network_name if current_network_name in network_names else '')

        cnt_name_grps = [''] + sorted(filtered_data['CNT_NAME_GRP'].dropna().unique())
        current_cnt_name_grp = self.cnt_name_grp_var.get()
        self.cnt_name_grp_dropdown['values'] = cnt_name_grps
        self.cnt_name_grp_var.set(current_cnt_name_grp if current_cnt_name_grp in cnt_name_grps else '')

        business_models = [''] + sorted(filtered_data['Business model'].dropna().unique())
        current_business_model = self.business_model_var.get()
        self.business_model_dropdown['values'] = business_models
        self.business_model_var.set(current_business_model if current_business_model in business_models else '')

        allocations = [''] + sorted(filtered_data['allocation'].dropna().unique())
        current_allocation = self.allocation_var.get()
        self.allocation_dropdown['values'] = allocations
        self.allocation_var.set(current_allocation if current_allocation in allocations else '')

        if network_name_selected or cnt_name_grp_selected or business_model_selected or allocation_selected:
            self.display_metadata(network_name_selected, cnt_name_grp_selected, allocation_selected)

    def display_metadata(self, network_name, cnt_name_grp=None, allocation=None):
        # refresh table
        if hasattr(self, 'tree'):
            self.tree.destroy()

        # frame dédié pour le treeview et les scrollbars
        tree_frame = ttk.Frame(self)
        tree_frame.grid(row=2, column=0, columnspan=3, pady=10, padx=10, sticky="nsew")

        #hauteur minimale pour e viter le resizing involontaire
        columns = self.model_columns.get(self.business_model_var.get(), [])
        # hauteur fixe
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=20)

        # vertical scroll
        y_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")

        # horiz scroll
        x_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")

        # scrolbar to treeview
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")

        # permet expand
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # main layout, évite resizing
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)

        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=tkFont.Font().measure(col) + 20)

        # filtre data par copie
        filtered_rows = self.data[self.data['NETWORK_NAME'] == network_name].copy()
        if cnt_name_grp:
            filtered_rows = filtered_rows[filtered_rows['CNT_NAME_GRP'] == cnt_name_grp]
        if allocation:
            filtered_rows = filtered_rows[filtered_rows['allocation'] == allocation]

        cost_dest = self.config_data.get('cost_dest', '')
        working_contracts_file = os.path.join(cost_dest, 'working_contracts.xlsx')
        additional_data = pd.DataFrame()

        if os.path.exists(working_contracts_file) and business_model_selected:
            print(
                f"Debug: Loading additional contracts from {working_contracts_file} (Sheet: {business_model_selected})")
            try:
                additional_data = pd.read_excel(working_contracts_file, sheet_name=business_model_selected)

                additional_data_filtered = additional_data[additional_data['NETWORK_NAME'] == network_name]
                if cnt_name_grp:
                    additional_data_filtered = additional_data_filtered[
                        additional_data_filtered['CNT_NAME_GRP'] == cnt_name_grp]
                additional_data_filtered['Source'] = 'cost_dest'
                additional_data = additional_data_filtered
            except ValueError as e:
                print(f"Debug: {e}. Proceeding with only reference entries.")
                additional_data = pd.DataFrame()

        if not filtered_rows.empty and not additional_data.empty:
            combined_data = pd.concat([filtered_rows, additional_data], ignore_index=True)
        elif not filtered_rows.empty:
            combined_data = filtered_rows
        else:
            combined_data = additional_data

        # différents formats de date (4567 ou 01-12-2024)
        def convert_to_standard_date_format(date):
            if pd.isna(date):
                # si no value retourne no value
                return date
            if isinstance(date, str):
                try:
                    return pd.to_datetime(date, format='%d-%m-%Y', errors='raise').strftime('%d-%m-%Y')
                except (ValueError, TypeError):
                    # retourne la valeure originale si pas formattée
                    return date
            elif isinstance(date, (int, float)):
                # retourne nombres en string
                return str(date)
            return date

        if 'CT_STARTDATE' in combined_data.columns:
            combined_data['CT_STARTDATE'] = combined_data['CT_STARTDATE'].apply(convert_to_standard_date_format)
        if 'CT_ENDDATE' in combined_data.columns:
            combined_data['CT_ENDDATE'] = combined_data['CT_ENDDATE'].apply(convert_to_standard_date_format)

        # ordre par date
        if 'CT_STARTDATE' in combined_data.columns:
            try:
                combined_data.sort_values(by='CT_STARTDATE', ascending=False, inplace=True)
            except Exception as e:
                print(f"Warning: Could not sort by CT_STARTDATE due to mixed data types: {e}")

        for _, row in combined_data.iterrows():
            values = [row.get(col, '') for col in columns]
            if row.get('Source') == 'cost_dest':
                self.tree.insert("", tk.END, values=values, tags=('highlight',))
            else:
                self.tree.insert("", tk.END, values=values)

        self.tree.tag_configure('highlight', background='#ffffcc')

        for col in columns:
            max_width = max((tkFont.Font().measure(str(self.tree.set(item, col))) for item in self.tree.get_children()),
                            default=100)
            self.tree.column(col, width=max_width + 20)

    def open_new_deal_popup(self):
        network_name = self.network_name_var.get()
        cnt_name_grp = self.cnt_name_grp_var.get()
        business_model = self.business_model_var.get()
        allocation = self.allocation_var.get()

        if not business_model:
            show_message("Warning", "Please select a business model", master=self, custom=True)
            return

        columns = self.model_columns.get(business_model, [])
        new_deal_popup = tk.Toplevel(self)
        new_deal_popup.title("New Deal")
        new_deal_popup.geometry("400x650")

        entry_vars = {col: tk.StringVar() for col in columns}

        tk.Label(new_deal_popup, text="Business model").grid(row=0, column=0, padx=10, pady=5, sticky='e')
        business_model_combobox = ttk.Combobox(new_deal_popup, textvariable=tk.StringVar(value=business_model))
        business_model_combobox['values'] = self.business_model_dropdown['values']
        business_model_combobox.grid(row=0, column=1, padx=10, pady=5, sticky='w')
        business_model_combobox.config(state='readonly')

        for i, col in enumerate(columns, start=1):
            tk.Label(new_deal_popup, text=col).grid(row=i, column=0, padx=10, pady=5, sticky='e')

            if col == 'allocation':
                allocation_combobox = ttk.Combobox(new_deal_popup, textvariable=entry_vars[col])
                allocation_combobox['values'] = self.allocation_dropdown['values']
                allocation_combobox.grid(row=i, column=1, padx=10, pady=5, sticky='w')
                allocation_combobox.set(allocation)
                allocation_combobox.config(state='readonly' if allocation else 'normal')
            elif col in ['NETWORK_NAME', 'CNT_NAME_GRP']:
                entry = tk.Entry(new_deal_popup, textvariable=entry_vars[col])
                entry.grid(row=i, column=1, padx=10, pady=5, sticky='w')
                entry.insert(0, network_name if col == 'NETWORK_NAME' else cnt_name_grp)
                entry.config(state='readonly')
            elif col == 'CT_TYPE' and business_model in ['Fixed fee', 'fixed fee']:
                entry = tk.Entry(new_deal_popup, textvariable=entry_vars[col])
                entry.grid(row=i, column=1, padx=10, pady=5, sticky='w')
                entry.insert(0, 'F')
                entry.config(state='readonly')
            else:
                entry = tk.Entry(new_deal_popup, textvariable=entry_vars[col])
                entry.grid(row=i, column=1, padx=10, pady=5, sticky='w')

        def submit_deal():
            new_row = {col: entry_vars[col].get() for col in columns}

            new_row['Business model'] = business_model

            required_fields = ['allocation', 'Business model', 'NETWORK_NAME', 'CNT_NAME_GRP']
            for field in required_fields:
                if not new_row.get(field):
                    show_message("Error", f"Field '{field}' cannot be empty.", master=new_deal_popup, custom=True)
                    return

            if business_model == 'Fixed fee' or business_model == 'fixed fee':
                new_row['variable/fix'] = 'fixed'

            self.generate_template(new_row)
            new_deal_popup.destroy()
            self.display_metadata(self.network_name_var.get(), self.cnt_name_grp_var.get(),
                                  self.business_model_var.get())

        submit_button = ttk.Button(new_deal_popup, text="Save", command=submit_deal)
        submit_button.grid(row=len(columns) + 1, column=0, padx=10, pady=20, sticky='e')

    def open_update_deal_popup(self):
        selected_items = self.tree.selection()
        if not selected_items:
            show_message("Warning", "Please select at least one deal to update", master=self, custom=True)
            return

        self.update_deal_popup = tk.Toplevel(self)
        self.update_deal_popup.title("Update Deal")
        self.update_deal_popup.geometry("400x600")

        self.current_update_index = 0
        self.items_to_update = selected_items

        self.update_deal_entries = {col: tk.StringVar() for col in self.tree["columns"]}

        for i, col in enumerate(self.tree["columns"]):
            tk.Label(self.update_deal_popup, text=col).grid(row=i, column=0, padx=10, pady=5, sticky='e')
            entry = tk.Entry(self.update_deal_popup, textvariable=self.update_deal_entries[col])
            entry.grid(row=i, column=1, padx=10, pady=5, sticky='w')
            if col in ['NETWORK_NAME', 'CNT_NAME_GRP']:
                entry.config(state='readonly')

        def submit_update():
            new_values = {col: self.update_deal_entries[col].get() for col in self.tree["columns"]}
            self.update_deal_row(self.current_update_index, new_values)
            self.current_update_index += 1
            if self.current_update_index < len(self.items_to_update):
                self.populate_update_deal_entries()
            else:
                self.update_deal_popup.destroy()

        def cancel_update():
            self.current_update_index += 1
            if self.current_update_index < len(self.items_to_update):
                self.populate_update_deal_entries()
            else:
                self.update_deal_popup.destroy()

        submit_button = ttk.Button(self.update_deal_popup, text="Save", command=submit_update)
        submit_button.grid(row=len(self.tree["columns"]), column=0, padx=10, pady=20, sticky='e')

        cancel_button = ttk.Button(self.update_deal_popup, text="Cancel", command=cancel_update)
        cancel_button.grid(row=len(self.tree["columns"]), column=1, padx=10, pady=20, sticky='w')

        self.populate_update_deal_entries()

    def populate_update_deal_entries(self):
        item_id = self.items_to_update[self.current_update_index]
        values = self.tree.item(item_id, "values")
        for col, value in zip(self.tree["columns"], values):
            self.update_deal_entries[col].set(value)

    def update_deal_row(self, index, new_values):
        item_id = self.items_to_update[index]
        item_idx = self.tree.index(item_id)

        for col, val in new_values.items():
            self.data.at[item_idx, col] = val

        self.save_updated_data()
        self.data.sort_values(by='CT_BOOK_YEAR', ascending=False, inplace=True)
        self.display_metadata(self.network_name_var.get(), self.cnt_name_grp_var.get(), self.business_model_var.get())

    def add_new_deal_row(self, new_row):
        self.data = pd.concat([self.data, pd.DataFrame([new_row])], ignore_index=True)
        self.save_updated_data()
        self.data.sort_values(by='CT_BOOK_YEAR', ascending=False, inplace=True)
        self.display_metadata(self.network_name_var.get(), self.cnt_name_grp_var.get(), self.business_model_var.get())

    def save_updated_data(self):
        with pd.ExcelWriter(self.file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            self.data.to_excel(writer, sheet_name='all contract cost file', index=False)

    def show_tooltip(self, event, text):
        self.tooltip = utils.tooltip_show(event, text, self)

    def hide_tooltip(self):
        utils.tooltip_hide(self.tooltip)

    def clear_fields(self):
        self.network_name_var.set('')
        self.cnt_name_grp_var.set('')
        self.business_model_var.set('')
        self.allocation_var.set('')

        self.tree.delete(*self.tree.get_children())
        self.populate_dropdowns()


if __name__ == "__main__":
    root = tk.Tk()
    config_manager = ConfigManager()
    tab = CostTab(root, config_manager=config_manager)
    tab.pack(expand=1, fill='both')
    root.mainloop()
