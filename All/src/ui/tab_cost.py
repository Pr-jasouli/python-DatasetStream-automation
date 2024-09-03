import os
import tkinter as tk
import tkinter.font as tkFont
from datetime import timedelta, datetime
from tkinter import ttk

import pandas as pd

from parser.fixedFee_cogs import FixedFeeCogsLevelHandler
from parser.fixedFee_index import FixedFeeIndexLevelHandler
from parser.fixedFee_channelGroupLevel import FixedFeeChannelGroupLevelHandler
from parser.fixedFee_providerLevel import FixedFeeProviderLevelHandler
from parser.free import FreeLevelHandler
from utilities import utils
from utilities.config_manager import ConfigManager
from utilities.utils import show_message, open_file_and_update_config, set_window_icon, center_window


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
        self.prod_en_name_var = tk.StringVar()
        self.business_model_var = tk.StringVar()
        self.allocation_var = tk.StringVar()

        self.base_dir = base_dir

        self.init_ui()

    def generate_template(self, new_row):
        output_dir = os.path.abspath(os.path.join(self.config_data.get('cost_dest', ''), '../outputs'))
        working_contracts_file = os.path.join(output_dir, 'working_contracts.xlsx')

        if not os.path.exists(output_dir):
            print(f"Debug: Output directory '{output_dir}' does not exist. Creating it.")
            os.makedirs(output_dir)
        else:
            print(f"Debug: Output directory '{output_dir}' already exists.")

        print(f"Debug: Working contracts file path: {working_contracts_file}")

        business_model = new_row['Business model'].capitalize()

        # Possible date formats to handle dd-mm-yyyy and dd-mm-yy
        date_formats = ['%d-%m-%Y', '%d-%m-%y']

        def parse_date(date_str):
            for fmt in date_formats:
                try:
                    return pd.to_datetime(date_str, format=fmt, dayfirst=True)
                except ValueError:
                    continue
            raise ValueError(f"Date {date_str} is not in a recognized format")

        try:
            new_row['CT_STARTDATE'] = parse_date(new_row['CT_STARTDATE'])
            new_row['CT_ENDDATE'] = parse_date(new_row['CT_ENDDATE'])
        except Exception as e:
            print(f"Error converting dates: {e}")
            show_message("Error", f"Invalid date format in CT_STARTDATE or CT_ENDDATE: {e}", master=self, custom=True)
            return

        start_year = new_row['CT_STARTDATE'].year
        end_year = new_row['CT_ENDDATE'].year

        # Prepare the rows to be added, one per year
        rows_to_add = []
        for year in range(start_year, end_year + 1):
            row_copy = new_row.copy()
            row_copy['YEAR'] = year

            # Format dates back to dd-mm-yyyy
            row_copy['CT_STARTDATE'] = row_copy['CT_STARTDATE'].strftime('%d-%m-%Y')
            row_copy['CT_ENDDATE'] = row_copy['CT_ENDDATE'].strftime('%d-%m-%Y')

            rows_to_add.append(row_copy)

        new_df = pd.DataFrame(rows_to_add)

        try:
            if not os.path.exists(working_contracts_file):
                print(f"Debug: File '{working_contracts_file}' does not exist. Creating new file.")
                with pd.ExcelWriter(working_contracts_file, engine='openpyxl') as writer:
                    new_df.to_excel(writer, sheet_name=business_model, index=False)
                    print(f"Debug: Created new working contracts file with sheet '{business_model}'.")
            else:
                print(f"Debug: File '{working_contracts_file}' exists. Appending data.")
                with pd.ExcelWriter(working_contracts_file, engine='openpyxl', mode='a',
                                    if_sheet_exists='overlay') as writer:
                    try:
                        existing_df = pd.read_excel(working_contracts_file, sheet_name=business_model)
                        updated_df = pd.concat([existing_df, new_df], ignore_index=True)
                    except ValueError:
                        updated_df = new_df

                    updated_df.to_excel(writer, sheet_name=business_model, index=False)
                    print(f"Debug: Updated '{business_model}' sheet in the working contracts file.")
        except PermissionError:
            show_message("Error", "Please close the 'working_contracts.xlsx' file and try saving again.", master=self,
                         custom=True)
            return

        show_message("Success", f"Contract added to {working_contracts_file} under '{business_model}' sheet.",
                     master=self, custom=True)


    def get_model_columns(self):
        return {
            ### OK 100%
            'Fixed fee': ['YEAR', 'Business model', 'allocation', 'NETWORK_NAME', 'CT_STARTDATE', 'CT_ENDDATE',
                          'CT_DURATION', 'CT_FIXFEE_NEW', 'CT_TYPE', 'CT_AUTORENEW', 'CT_NOTICE_DATE', 'CT_NOTICE_PER',
                          'CT_AVAIL_IN_SCARLET_FR', 'CT_AVAIL_IN_SCARLET_NL'],

            'fixed fee': ['YEAR', 'Business model', 'allocation', 'NETWORK_NAME', 'CT_STARTDATE', 'CT_ENDDATE',
                          'CT_DURATION', 'CT_FIXFEE_NEW', 'CT_TYPE', 'CT_AUTORENEW', 'CT_NOTICE_DATE', 'CT_NOTICE_PER',
                          'CT_AVAIL_IN_SCARLET_FR', 'CT_AVAIL_IN_SCARLET_NL'],

            'Free': ['YEAR', 'Business model', 'allocation', 'NETWORK_NAME', 'CT_STARTDATE', 'CT_ENDDATE',
                     'CT_DURATION', 'CT_TYPE', 'CT_AUTORENEW', 'CNT_NAME_GRP', 'CT_NOTICE_DATE', 'CT_NOTICE_PER',
                     'CT_AVAIL_IN_SCARLET_FR',
                     'CT_AVAIL_IN_SCARLET_NL', ],

            'free': ['YEAR', 'Business model', 'allocation', 'NETWORK_NAME', 'CT_STARTDATE', 'CT_ENDDATE',
                     'CT_DURATION', 'CT_TYPE', 'CT_AUTORENEW', 'CNT_NAME_GRP', 'CT_NOTICE_DATE', 'CT_NOTICE_PER',
                     'CT_AVAIL_IN_SCARLET_FR',
                     'CT_AVAIL_IN_SCARLET_NL', ],

            'fixed fee + index': ['YEAR', 'Business model', 'allocation', 'NETWORK_NAME', 'CT_STARTDATE', 'CT_ENDDATE',
                                  'CT_DURATION', 'CT_FIXFEE_NEW', 'CT_INDEX', 'CT_TYPE', 'CT_AUTORENEW',
                                  'CT_NOTICE_DATE', 'CT_NOTICE_PER', 'CT_AVAIL_IN_SCARLET_FR',
                                  'CT_AVAIL_IN_SCARLET_NL'],

            'Fixed fee cogs': ['YEAR', 'Business model', 'allocation', 'NETWORK_NAME', 'CT_STARTDATE', 'CT_ENDDATE',
                               'CT_DURATION', 'CT_FIXFEE_NEW', 'CT_TYPE', 'CT_AUTORENEW', 'CT_NOTICE_DATE',
                               'CT_NOTICE_PER', 'CT_AVAIL_IN_SCARLET_FR', 'CT_AVAIL_IN_SCARLET_NL'],

            'fixed fee cogs': [
                'CT_STARTDATE', 'CT_ENDDATE', 'allocation', 'NETWORK_NAME', 'CNT_NAME_GRP', 'PROD_EN_NAME',
                'CT_TYPE', 'CT_DURATION', 'CT_NOTICE_DATE', 'CT_AUTORENEW', 'CT_NOTICE_PER',
                'CT_AVAIL_IN_SCARLET_FR', 'CT_AVAIL_IN_SCARLET_NL', 'CT_FIXFEE', 'CT_FIXFEE_NEW',
                'Business model'
            ],

            'CPS Over MG Subs': ['YEAR', 'Business model', 'allocation', 'NETWORK_NAME', 'CT_STARTDATE', 'CT_ENDDATE', 'CT_DURATION', 'CT_FIXFEE_NEW', 'CT_TYPE', 'CT_AUTORENEW', 'CT_NOTICE_DATE', 'CT_NOTICE_PER', 'CT_AVAIL_IN_SCARLET_FR', 'CT_AVAIL_IN_SCARLET_NL'],




            'CPS Over MG Subs + index': [
                'CT_STARTDATE', 'CT_ENDDATE', 'allocation', 'NETWORK_NAME', 'CNT_NAME_GRP', 'PROD_EN_NAME',
                'CT_TYPE', 'CT_DURATION', 'CT_NOTICE_DATE', 'CT_AUTORENEW', 'CT_NOTICE_PER',
                'CT_AVAIL_IN_SCARLET_FR', 'CT_AVAIL_IN_SCARLET_NL', 'CT_INDEX', 'CT_MG', 'CT_VARFEE', 'CT_VARFEE_NEW',
                'CT_STEP1_SUBS', 'CT_STEP2_SUBS', 'CT_STEP3_SUBS'
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
                'CT_AVAIL_IN_SCARLET_FR', 'CT_AVAIL_IN_SCARLET_NL', 'PROD_PRICE', 'CT_STEP1_SUBS', 'CT_STEP2_SUBS',
                'CT_STEP3_SUBS'  #, '#Subscriber', 'PROD_PRICE_VAT_EXCL'
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
            'Revenue share': [
                'CT_TYPE', 'allocation'
            ],
            'CPS- 5 steps': [
                'CT_TYPE', 'allocation'
            ],
        }

    def load_file(self, path):
        self.file_path = path
        self.load_cost_reference_file(path)

    def convert_excel_date(self, excel_serial):
        """Convert an Excel serial date to a Python datetime object."""
        try:
            if pd.isna(excel_serial):
                return None
            if isinstance(excel_serial, (int, float)):
                base_date = datetime(1899, 12, 30)
                return base_date + timedelta(days=int(excel_serial))
            return pd.to_datetime(excel_serial, errors='coerce')
        except Exception as e:
            print(f"Error converting date: {excel_serial} - {e}")
            return None

    def load_cost_reference_file(self, file_path):
        try:
            self.data = pd.read_excel(file_path, sheet_name='all contract cost file')

            self.data['CT_STARTDATE'] = self.data['CT_STARTDATE'].apply(self.convert_excel_date)
            self.data['CT_ENDDATE'] = self.data['CT_ENDDATE'].apply(self.convert_excel_date)

            self.populate_dropdowns()

            self.network_name_dropdown.config(state='normal')
            self.cnt_name_grp_dropdown.config(state='normal')
            self.business_model_dropdown.config(state='normal')
            self.allocation_dropdown.config(state='normal')

            self.display_metadata(self.network_name_var.get(), self.cnt_name_grp_var.get(),
                                  self.business_model_var.get())
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

        # chargement data additionnels
        cost_dest = self.config_data.get('cost_dest', '')
        working_contracts_file = os.path.join(cost_dest, 'working_contracts.xlsx')
        additional_data = pd.DataFrame()

        if os.path.exists(working_contracts_file):
            print(f"Debug: Loading additional contracts from {working_contracts_file}")
            try:
                additional_data = pd.read_excel(working_contracts_file, sheet_name=self.business_model_var.get())

                additional_data_filtered = additional_data[additional_data['NETWORK_NAME'] == network_name].copy()
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


        if 'CT_STARTDATE' in combined_data.columns:
            combined_data['CT_STARTDATE'] = pd.to_datetime(combined_data['CT_STARTDATE'], errors='coerce').dt.strftime(
                '%d-%m-%Y')
        if 'CT_ENDDATE' in combined_data.columns:
            combined_data['CT_ENDDATE'] = pd.to_datetime(combined_data['CT_ENDDATE'], errors='coerce').dt.strftime(
                '%d-%m-%Y')

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

        # treeview occupe la taille maximale
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)

    def open_new_deal_popup(self):
        network_name = self.network_name_var.get()
        allocation = self.allocation_var.get()
        business_model = self.business_model_var.get()

        new_deal_popup = tk.Toplevel(self)
        new_deal_popup.title("New Deal")
        new_deal_popup.geometry("1080x625")
        set_window_icon(new_deal_popup)

        business_model_var = tk.StringVar(value=business_model)  # Pre-fill the business model
        entry_vars = {}

        left_frame = ttk.Frame(new_deal_popup)
        left_frame.grid(row=0, column=0, padx=10, pady=5, sticky='nsew')

        dynamic_widgets = []

        def clear_dynamic_widgets():
            for widget in dynamic_widgets:
                widget.grid_forget()
                widget.destroy()
            dynamic_widgets.clear()

        allocation_combobox = None
        allocation_options = {
            'Fixed fee': ['channel group level', 'product level', 'provider level', 'provider level (gulli excluded)'],
            'fixed fee + index': ['provider & product level', 'provider level', 'provider level & product(basic)'],
            'Fixed fee cogs': ['channel group level', 'channel level', 'product level', 'provider level'],
            'Free': ['channel group level', 'product level'],
            'CPS Over MG Subs': ['provider level', 'provider level & product(options)'],
            'CPS Over MG Subs + index': ['provider & product level'],
            'Revenue share over MG subs': ['provider level'],
            'CPS on product park': [
                'channel group level', 'subscriber based - provider level', 'subscriber based (only basic)',
                'subscriber based (pxs/scarlet)', 'suscriber based FR/NL(only basic)', 'product price level(basic)'
            ],
            'CPS on volume regionals + index': ['provider level', 'subscriber based - provider level'],
            'Event Based INTEC': ['product level', 'provider level'],
            'Revenue share': ['channel group level', 'product level'],
            'CPS- 5 steps': ['provider level']
        }

        def update_allocation_options(selected_model):
            options = allocation_options.get(selected_model, sorted(self.data['allocation'].dropna().unique()))
            allocation_combobox['values'] = options
            if options:
                allocation_combobox.set(options[0])

        def update_fields_based_on_business_model(event=None):
            clear_dynamic_widgets()

            selected_model = business_model_var.get()
            current_columns = self.model_columns.get(selected_model, [])

            field_order = ['Business model', 'allocation', 'NETWORK_NAME', 'DATA_TYPE', 'variable/fix',
                           'CT_FIXFEE_NEW', 'CT_STARTDATE', 'CT_ENDDATE'] + \
                          [col for col in current_columns if col not in [
                              'Business model', 'allocation', 'NETWORK_NAME', 'DATA_TYPE', 'CT_FIXFEE',
                              'CT_FIXFEE_NEW', 'CT_STARTDATE', 'CT_ENDDATE', 'variable/fix']]

            for i, col in enumerate(field_order, start=1):
                if col in ['Business model']:
                    continue

                label = tk.Label(left_frame, text=col)
                label.grid(row=i + 2, column=0, padx=10, pady=5, sticky='e')
                dynamic_widgets.append(label)

                entry_vars[col] = tk.StringVar()

                if col == 'allocation':
                    nonlocal allocation_combobox
                    allocation_combobox = ttk.Combobox(left_frame, textvariable=entry_vars[col])
                    allocation_combobox.grid(row=i + 2, column=1, padx=10, pady=5, sticky='w')
                    dynamic_widgets.append(allocation_combobox)
                    update_allocation_options(
                        selected_model)
                elif col == 'NETWORK_NAME':
                    network_name_combobox = ttk.Combobox(left_frame, textvariable=entry_vars[col])
                    network_name_combobox['values'] = sorted(self.data['NETWORK_NAME'].dropna().unique().tolist())
                    network_name_combobox.grid(row=i + 2, column=1, padx=10, pady=5, sticky='w')
                    dynamic_widgets.append(network_name_combobox)
                    network_name_combobox.set(network_name)
                    network_name_combobox.bind("<<ComboboxSelected>>", lambda e: update_channels_listbox())
                elif col == 'CT_TYPE':
                    if selected_model.lower() == 'fixed fee':
                        entry_vars[col].set('F')
                        ct_type_label = tk.Label(left_frame, text='F')
                        ct_type_label.grid(row=i + 2, column=1, padx=10, pady=5, sticky='w')
                        dynamic_widgets.append(ct_type_label)
                    else:
                        ct_type_combobox = ttk.Combobox(left_frame, textvariable=entry_vars[col])
                        ct_type_combobox['values'] = ['F', 'V']
                        ct_type_combobox.grid(row=i + 2, column=1, padx=10, pady=5, sticky='w')
                        dynamic_widgets.append(ct_type_combobox)
                elif col == 'CT_AUTORENEW':
                    ct_autorenew_combobox = ttk.Combobox(left_frame, textvariable=entry_vars[col])
                    ct_autorenew_combobox['values'] = ['Yes', 'No']
                    ct_autorenew_combobox.grid(row=i + 2, column=1, padx=10, pady=5, sticky='w')
                    dynamic_widgets.append(ct_autorenew_combobox)
                elif col == 'variable/fix':
                    if selected_model.lower() == 'fixed fee':
                        entry_vars[col].set('fixed')
                        variable_fix_label = tk.Label(left_frame, text='fixed')
                        variable_fix_label.grid(row=i + 2, column=1, padx=10, pady=5, sticky='w')
                        dynamic_widgets.append(variable_fix_label)
                    else:
                        variable_fix_combobox = ttk.Combobox(left_frame, textvariable=entry_vars[col])
                        variable_fix_combobox['values'] = ['variable', 'fixed']
                        variable_fix_combobox.grid(row=i + 2, column=1, padx=10, pady=5, sticky='w')
                        dynamic_widgets.append(variable_fix_combobox)
                elif col == 'DATA_TYPE':
                    data_type_combobox = ttk.Combobox(left_frame, textvariable=entry_vars[col])
                    data_type_combobox['values'] = ['ACTUALS', 'FORECAST', 'PLAN']
                    data_type_combobox.grid(row=i + 2, column=1, padx=10, pady=5, sticky='w')
                    dynamic_widgets.append(data_type_combobox)
                else:
                    entry = tk.Entry(left_frame, textvariable=entry_vars[col])
                    entry.grid(row=i + 2, column=1, padx=10, pady=5, sticky='w')
                    dynamic_widgets.append(entry)


        update_fields_based_on_business_model()

        submit_button = ttk.Button(left_frame, text="Save", command=lambda: submit_deal(business_model_var, entry_vars),
                                   style='AudienceTab.TButton')
        submit_button.grid(row=0, column=0, padx=5, pady=5, sticky='w')

        add_button = ttk.Button(left_frame, text="Add Another Selection", command=lambda: add_listbox_pair(),
                                style='AudienceTab.TButton')
        add_button.grid(row=0, column=1, padx=5, pady=5, sticky='e')

        tk.Label(left_frame, text="Business model").grid(row=1, column=0, padx=10, pady=5, sticky='e')
        business_model_combobox = ttk.Combobox(left_frame, textvariable=business_model_var)
        business_model_combobox['values'] = list(self.model_columns.keys())
        business_model_combobox.grid(row=1, column=1, padx=10, pady=5, sticky='w')

        business_model_combobox.bind("<<ComboboxSelected>>", update_fields_based_on_business_model)


        right_frame = ttk.Frame(new_deal_popup)
        right_frame.grid(row=0, column=1, padx=10, pady=5, sticky='nsew')

        dynamic_listbox_pairs = []

        def load_channel_grouping_data():
            try:
                channel_grouping_src = self.config_data.get('channel_grouping_src', None)
                if channel_grouping_src:
                    return pd.read_excel(channel_grouping_src, sheet_name='Content_Channel_Grouping')
                else:
                    return pd.DataFrame()
            except Exception as e:
                print(f"Error loading channel grouping data: {e}")
                return pd.DataFrame()

        def get_channels_for_network(network_name):
            channel_data = load_channel_grouping_data()
            if not channel_data.empty and 'CHANNEL_NETWORK_GROUP' in channel_data.columns:
                filtered_channels = channel_data[channel_data['CHANNEL_NETWORK_GROUP'] == network_name]
                if not filtered_channels.empty:
                    return sorted(filtered_channels['CHANNEL_NAME'].dropna().unique())
            return sorted(channel_data['CHANNEL_NAME'].dropna().unique()) if not channel_data.empty else []

        def update_channels_listbox():
            channels = get_channels_for_network(entry_vars['NETWORK_NAME'].get())
            for pair in dynamic_listbox_pairs:
                pair[0].delete(0, tk.END)
                for item in channels:
                    pair[0].insert(tk.END, item)

        def add_listbox_pair():
            # Ensure channels are fetched using the correct network name
            channels = get_channels_for_network(entry_vars['NETWORK_NAME'].get())

            # Assign channels and TV packs to listboxes
            pair_frame = ttk.Frame(right_frame)
            pair_frame.grid(row=len(dynamic_listbox_pairs) * 2, column=0, padx=5, pady=5, sticky='nsew')

            tk.Label(pair_frame, text="Select Channels:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
            tk.Label(pair_frame, text="Select TV Packs:").grid(row=0, column=1, padx=5, pady=5, sticky='w')

            channels_listbox = tk.Listbox(pair_frame, selectmode='multiple', height=6, exportselection=False)
            channels_listbox.grid(row=1, column=0, padx=5, pady=5, sticky='w')

            for item in channels:
                channels_listbox.insert(tk.END, item)

            packs_listbox = tk.Listbox(pair_frame, selectmode='multiple', height=6, exportselection=False)
            packs_listbox.grid(row=1, column=1, padx=5, pady=5, sticky='w')

            packs = sorted(self.data['PROD_EN_NAME'].dropna().unique())
            for item in packs:
                packs_listbox.insert(tk.END, item)

            dynamic_listbox_pairs.append((channels_listbox, packs_listbox))

        # paires initiales
        add_listbox_pair()

        def submit_deal(business_model_var, entry_vars):
            selected_model = business_model_var.get()
            current_columns = self.model_columns.get(selected_model, [])

            # validation des fields
            # required_fields = {
            #     'CT_STARTDATE': entry_vars.get('CT_STARTDATE', tk.StringVar()).get(),
            #     'CT_ENDDATE': entry_vars.get('CT_ENDDATE', tk.StringVar()).get(),
            #     'allocation': entry_vars.get('allocation', tk.StringVar()).get(),
            #     'NETWORK_NAME': entry_vars.get('NETWORK_NAME', tk.StringVar()).get(),
            #     'Business model': business_model_var.get(),
            # }

            for channels_listbox, packs_listbox in dynamic_listbox_pairs:
                selected_channels = [channels_listbox.get(idx) for idx in channels_listbox.curselection()]
                selected_packs = [packs_listbox.get(idx) for idx in packs_listbox.curselection()]

                if not selected_channels or not selected_packs:
                    continue

                for channel in selected_channels:
                    for pack in selected_packs:
                        new_row = {
                            'NETWORK_NAME': entry_vars.get('NETWORK_NAME', tk.StringVar()).get(),
                            'CNT_NAME_GRP': channel,
                            'PROD_EN_NAME': pack,
                            'Business model': business_model_var.get(),
                            'allocation': entry_vars.get('allocation', tk.StringVar()).get(),
                            'CT_TYPE': entry_vars.get('CT_TYPE', tk.StringVar()).get(),
                            'CT_AUTORENEW': entry_vars.get('CT_AUTORENEW', tk.StringVar()).get(),
                            'variable/fix': entry_vars.get('variable/fix', tk.StringVar()).get(),
                            'CT_STARTDATE': entry_vars.get('CT_STARTDATE', tk.StringVar()).get(),
                            'CT_ENDDATE': entry_vars.get('CT_ENDDATE', tk.StringVar()).get(),
                            'DATA_TYPE': entry_vars.get('DATA_TYPE', tk.StringVar()).get(),
                            'CT_FIXFEE_NEW': entry_vars.get('CT_FIXFEE_NEW', tk.StringVar()).get(),
                        }

                        for col in current_columns:
                            if col not in new_row:
                                new_row[col] = entry_vars.get(col, tk.StringVar()).get()

                        if (new_row['Business model'].lower() == 'fixed fee' and
                                new_row['allocation'].lower() == 'provider level'):
                            handler = FixedFeeProviderLevelHandler(new_row)
                            handler.add_additional_fields()
                        if (new_row['Business model'].lower() == 'fixed fee' and
                                new_row['allocation'].lower() == 'channel group level'):
                            handler = FixedFeeChannelGroupLevelHandler(new_row)
                            handler.add_additional_fields()
                        if new_row['Business model'].lower() == 'free':
                            handler = FreeLevelHandler(new_row)
                            handler.add_additional_fields()
                        if new_row['Business model'].lower() == 'fixed fee + index':
                            handler = FixedFeeIndexLevelHandler(new_row)
                            handler.add_additional_fields()
                        if new_row['Business model'].lower() == 'Fixed fee cogs':
                            handler = FixedFeeCogsLevelHandler(new_row)
                            handler.add_additional_fields()

                        self.generate_template(new_row)

            new_deal_popup.destroy()
            self.display_metadata(self.network_name_var.get(), self.cnt_name_grp_var.get(),
                                  self.business_model_var.get())

        update_fields_based_on_business_model()

    ######## refefefef


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
