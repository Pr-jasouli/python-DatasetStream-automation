import tkinter as tk
import tkinter.font as tkFont
from tkinter import ttk, filedialog

import pandas as pd

from utilities import utils
from utilities.config_manager import ConfigManager
from utilities.utils import show_message


class CostTab(ttk.Frame):
    def __init__(self, parent, config_manager=None, config_ui_callback=None):
        super().__init__(parent)
        self.config_manager = config_manager
        self.config_ui_callback = config_ui_callback
        self.config_data = config_manager.get_config()
        self.file_path = None
        self.data = None
        self.network_name_var = tk.StringVar()
        self.cnt_name_grp_var = tk.StringVar()
        self.business_model_var = tk.StringVar()

        self.init_ui()
        self.model_columns = {

            'Fixed fee': ['CT_BOOK_YEAR', 'NETWORK_NAME', 'CNT_NAME_GRP', 'CT_TYPE', 'CT_STARTDATE', 'CT_ENDDATE',
                          'CT_DURATION', 'CT_NOTICE_DATE', 'CT_AUTORENEW', 'CT_NOTICE_PER', 'CT_AVAIL_IN_SCARLET_FR',
                          'CT_AVAIL_IN_SCARLET_NL', 'CT_FIXFEE', 'CT_FIXFEE_NEW', 'Business model', 'variable/fix'],
            'fixed fee': ['CT_BOOK_YEAR', 'NETWORK_NAME', 'CNT_NAME_GRP', 'CT_TYPE', 'CT_STARTDATE', 'CT_ENDDATE',
                          'CT_DURATION', 'CT_NOTICE_DATE', 'CT_AUTORENEW', 'CT_NOTICE_PER', 'CT_AVAIL_IN_SCARLET_FR',
                          'CT_AVAIL_IN_SCARLET_NL', 'CT_FIXFEE', 'CT_FIXFEE_NEW', 'Business model', 'variable/fix'],
        }

        if self.file_path:
            self.load_cost_reference_file(self.file_path)

    def load_file(self, path):
        self.file_path = path
        self.load_cost_reference_file(path)

    def load_cost_reference_file(self, file_path):
        self.data = pd.read_excel(file_path, sheet_name='all contract cost file')
        self.populate_dropdowns()

    def init_ui(self):
        style = ttk.Style()
        style.configure("Custom.TButton", font=('Helvetica', 10))

        top_frame = ttk.Frame(self)
        top_frame.grid(row=0, column=0, pady=10, padx=10, sticky="ew")

        tree_container = ttk.Frame(self)
        tree_container.grid(row=1, column=0, pady=10, padx=10, sticky="nsew")

        self.load_cost_button = ttk.Button(top_frame, text="Cost File", command=self.load_cost_data, width=9, style='AudienceTab.TButton')
        self.load_cost_button.grid(row=0, column=1, columnspan=3, pady=10, padx=4)
        self.load_cost_button.bind("<Enter>", lambda e: self.show_tooltip(e,"Business model\nsheet: all contract cost file\ncolumns: 'NETWORK_NAME', 'CNT_NAME_GRP', 'Business model'"))
        self.load_cost_button.bind("<Leave>", lambda e: self.hide_tooltip())


        ttk.Label(top_frame, text="Network:", font=('Helvetica', 10)).grid(row=0, column=4, pady=5, padx=4, sticky="e")
        self.network_name_dropdown = ttk.Combobox(top_frame, textvariable=self.network_name_var, state="readonly", font=('Helvetica', 10))
        self.network_name_dropdown.grid(row=0, column=5, pady=5, padx=4)

        ttk.Label(top_frame, text="Contract:", font=('Helvetica', 10)).grid(row=0, column=6, pady=5, padx=4, sticky="e")
        self.cnt_name_grp_dropdown = ttk.Combobox(top_frame, textvariable=self.cnt_name_grp_var, state="readonly", font=('Helvetica', 10))
        self.cnt_name_grp_dropdown.grid(row=0, column=7, pady=5, padx=4)

        ttk.Label(top_frame, text="Model:", font=('Helvetica', 10)).grid(row=0, column=8, pady=5, padx=4, sticky="e")
        self.business_model_dropdown = ttk.Combobox(top_frame, textvariable=self.business_model_var, state="readonly", font=('Helvetica', 10))
        self.business_model_dropdown.grid(row=0, column=9, pady=5, padx=4)

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

    def find_relevant_sheet(self, file_path):
        xls = pd.ExcelFile(file_path)
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            if {'NETWORK_NAME', 'CNT_NAME_GRP', 'Business model'}.issubset(df.columns):
                return df
        return None

    def populate_dropdowns(self):
        self.data['NETWORK_NAME'] = self.data['NETWORK_NAME'].astype(str)
        self.data['CNT_NAME_GRP'] = self.data['CNT_NAME_GRP'].astype(str)
        self.data['Business model'] = self.data['Business model'].astype(str)

        network_names = [''] + sorted(self.data['NETWORK_NAME'].dropna().unique())
        cnt_name_grps = [''] + sorted(self.data['CNT_NAME_GRP'].dropna().unique())
        business_models = [''] + sorted(self.data['Business model'].dropna().unique())

        self.network_name_dropdown['values'] = network_names
        self.cnt_name_grp_dropdown['values'] = cnt_name_grps
        self.business_model_dropdown['values'] = business_models

    def update_dropdowns(self, event=None):
        network_name_selected = self.network_name_var.get()
        cnt_name_grp_selected = self.cnt_name_grp_var.get()
        business_model_selected = self.business_model_var.get()

        filtered_data = self.data.copy()
        if network_name_selected:
            filtered_data = filtered_data[filtered_data['NETWORK_NAME'] == network_name_selected]
        if cnt_name_grp_selected:
            filtered_data = filtered_data[filtered_data['CNT_NAME_GRP'] == cnt_name_grp_selected]
        if business_model_selected:
            filtered_data = filtered_data[filtered_data['Business model'] == business_model_selected]

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

        if business_model_selected:
            self.display_metadata(network_name_selected, cnt_name_grp_selected, business_model_selected)

    def display_metadata(self, network_name, cnt_name_grp, business_model):
        for column in self.tree.get_children():
            self.tree.delete(column)

        columns = self.model_columns.get(business_model, [])

        self.tree["columns"] = columns
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=tkFont.Font().measure(col) + 20)

        filtered_rows = self.data[
            (self.data['NETWORK_NAME'] == network_name) &
            (self.data['CNT_NAME_GRP'] == cnt_name_grp) &
            (self.data['Business model'] == business_model)
        ]

        for _, row in filtered_rows.iterrows():
            values = [row[col] for col in columns]
            self.tree.insert("", tk.END, values=values)

        for col in columns:
            max_width = max((tkFont.Font().measure(str(self.tree.set(item, col))) for item in self.tree.get_children()), default=100)
            self.tree.column(col, width=max_width + 20)

    def load_cost_data(self):
        file_path = filedialog.askopenfilename(
            title="Select Cost File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.file_path = file_path
            self.load_cost_reference_file(file_path)

    def open_new_deal_popup(self):
        network_name = self.network_name_var.get()
        cnt_name_grp = self.cnt_name_grp_var.get()
        business_model = self.business_model_var.get().lower()

        if not business_model:
            show_message("Warning", "Please select a business model", master=self, custom=True)
            return

        columns = self.model_columns.get(business_model, [])

        new_deal_popup = tk.Toplevel(self)
        new_deal_popup.title("New Deal")
        new_deal_popup.geometry("400x600")

        entry_vars = {col: tk.StringVar() for col in columns}

        for i, col in enumerate(columns):
            #colonnes hidden
            if col not in ['Business model', 'variable/fix']:
                tk.Label(new_deal_popup, text=col).grid(row=i, column=0, padx=10, pady=5, sticky='e')
                entry = tk.Entry(new_deal_popup, textvariable=entry_vars[col])
                entry.grid(row=i, column=1, padx=10, pady=5, sticky='w')
                if col in ['NETWORK_NAME', 'CNT_NAME_GRP']:
                    entry.insert(0, network_name if col == 'NETWORK_NAME' else cnt_name_grp)
                    entry.config(state='readonly')

        def submit_deal():
            new_row = {col: entry_vars[col].get() for col in columns}
            # colonnes cachées
            if business_model == 'fixed fee' or business_model == 'Fixed fee':
                new_row['Business model'] = 'Fixed fee'
                new_row['variable/fix'] = 'fixed'
            self.add_new_deal_row(pd.Series(new_row))
            new_deal_popup.destroy()

        def cancel_deal():
            new_deal_popup.destroy()

        submit_button = ttk.Button(new_deal_popup, text="Save", command=submit_deal)
        submit_button.grid(row=len(columns), column=0, padx=10, pady=20, sticky='e')

        cancel_button = ttk.Button(new_deal_popup, text="Cancel", command=cancel_deal)
        cancel_button.grid(row=len(columns), column=1, padx=10, pady=20, sticky='w')

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
        self.display_metadata(self.network_name_var.get(), self.cnt_name_grp_var.get(), self.business_model_var.get())

    def add_new_deal_row(self, new_row):
        self.data = pd.concat([self.data, pd.DataFrame([new_row])], ignore_index=True)
        self.save_updated_data()
        self.display_metadata(self.network_name_var.get(), self.cnt_name_grp_var.get(), self.business_model_var.get())

    def save_updated_data(self):
        with pd.ExcelWriter(self.file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            self.data.to_excel(writer, sheet_name='all contract cost file', index=False)

    def show_tooltip(self, event, text):
        self.tooltip = utils.tooltip_show(event, text, self)

    def hide_tooltip(self):
        utils.tooltip_hide(self.tooltip)

if __name__ == "__main__":
    root = tk.Tk()
    config_manager = ConfigManager()
    tab = CostTab(root, config_manager=config_manager)
    tab.pack(expand=1, fill='both')
    root.mainloop()
