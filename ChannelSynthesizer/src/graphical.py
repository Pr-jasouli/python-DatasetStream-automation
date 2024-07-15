import os
import tkinter as tk
import webbrowser
from tkinter import filedialog, messagebox, Label, Button, OptionMenu, StringVar, Frame, Entry, Toplevel

from parser import PDFProcessor
from strategies import TelenetColumnStrategy, VOOColumnStrategy, OrangeColumnStrategy


class ChannelSynthesizerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Channel Synthesizer")
        self.files_frame = tk.Frame(self.root)
        self.files_frame.pack(padx=10, pady=10, expand=True, fill=tk.BOTH)
        self.add_file_button = Button(self.root, text="Load PDF", command=self.add_file)
        self.add_file_button.pack(pady=10)
        self.config_window = None
        self.process_button = Button(self.root, text="Process All", command=self.process_all_files)
        self.process_button.pack(pady=10)
        self.view_result_button = Button(self.root, text="View Result", command=self.view_result)
        self.view_result_button.pack(pady=10)
        self.files_data = []
        self.file_labels = {}
        self.strategies = [TelenetColumnStrategy(), VOOColumnStrategy(), OrangeColumnStrategy()]
        self.strategy_listbox = None
        self.pdf_processors = {}

    def add_file(self):
        file_path = filedialog.askopenfilename(title='Select PDF', filetypes=[("PDF files", "*.pdf")])
        if file_path:
            self.configure_file(file_path)

    def view_result(self):
        output_file = os.path.join(os.path.abspath('../outputs/xlsx'), 'extracted_data.xlsx')
        try:
            webbrowser.open(output_file)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open the file: {e}")

    def configure_file(self, file_path):
        file_settings_window = Toplevel(self.root)
        file_settings_window.title("Configure Parser")

        self.pdf_processors[file_path] = PDFProcessor(strategies=self.strategies)
        entries = self.pdf_processors[file_path].process_file_gui(file_path, file_settings_window)

        def on_done():
            if entries:
                page_include = [not bool(entry[2].get()) for entry in entries]
                if any(page_include):
                    processed_data = self.pdf_processors[file_path].process_pdf(entries, file_path)
                    if processed_data:
                        strategy_name = self.pdf_processors[file_path].get_strategy_name()
                        self.display_file_data(file_path, processed_data, strategy_name)
            file_settings_window.destroy()

        done_button = Button(file_settings_window, text="Done", command=on_done)
        done_button.pack(pady=20)

    def update_all_strategy_dropdowns(self):
        for file_data in self.files_data:
            file_path, _, strategy_var = file_data
            file_frame = self.file_labels.get(file_path)
            if file_frame:
                strategy_menu = next(
                    (widget for widget in file_frame.winfo_children() if isinstance(widget, tk.OptionMenu)), None)
                if strategy_menu:
                    menu = strategy_menu["menu"]
                    menu.delete(0, "end")
                    options = [s.name for s in self.strategies] + ['Configure strategies...']
                    for option in options:
                        menu.add_command(label=option, command=lambda opt=option, var=strategy_var: var.set(opt))
    #
    # def save_strategy(self, name, entries, details_frame):
    #     try:
    #         new_buffers_start = [int(entry[0].get()) for entry in entries]
    #         new_buffers_end = [int(entry[1].get()) for entry in entries]
    #         if self.config_window is not None:
    #             self.config_window.destroy()
    #             self.config_window = None
    #         messagebox.showinfo("Info", f"New strategy '{name}' created successfully.")
    #         self.update_strategy_list()
    #         self.update_all_strategy_dropdowns()
    #     except ValueError:
    #         messagebox.showerror("Error", "Please enter a valid unsigned integer for all buffer values.")

    def display_file_data(self, file_path, processed_data, strategy_name):
        file_frame = tk.Frame(self.files_frame)
        file_frame.pack(fill=tk.X, padx=5, pady=5)

        remove_button = Button(file_frame, text="❌", command=lambda: self.remove_file(file_path, file_frame))
        remove_button.pack(side=tk.LEFT, padx=5)

        label_text = os.path.basename(file_path) + ": " + ", ".join(
            f"Page {p[0]} ({p[1]} cols)" for p in processed_data)
        label = Label(file_frame, text=label_text)
        label.pack(side=tk.LEFT, expand=True, padx=5)

        strategy_var = StringVar(file_frame, value=strategy_name)
        strategy_var.trace("w", lambda *args: self.on_strategy_select(strategy_var.get(), file_path))

        options = [s.name for s in self.strategies] + ['Configure strategies...']
        strategy_menu = OptionMenu(file_frame, strategy_var, *options)
        strategy_menu.pack(side=tk.RIGHT, padx=5)

        self.file_labels[file_path] = file_frame
        self.files_data.append(
            (file_path, processed_data, strategy_var))

    def on_strategy_select(self, choice, file_path):
        if choice == "Configure strategies...":
            self.open_strategy_configurator()
        else:
            strategy_instance = next((s for s in self.strategies if s.name == choice), None)
            if strategy_instance and file_path in self.pdf_processors:
                self.pdf_processors[file_path].set_column_strategy(strategy_instance)

    def open_strategy_configurator(self):
        if self.config_window and self.config_window.winfo_exists():
            return
        self.config_window = tk.Toplevel(self.root)
        self.config_window.title("Strategy Configuration")
        self.config_window.protocol("WM_DELETE_WINDOW", self.on_config_window_close)

        list_frame = tk.Frame(self.config_window)
        list_frame.pack(side='left', fill='y')
        self.strategy_listbox = tk.Listbox(list_frame, width=30, height=10)
        self.strategy_listbox.pack(padx=5, pady=5)
        self.update_strategy_list()
        details_frame = tk.Frame(self.config_window)
        details_frame.pack(side='right', fill='both', expand=True, pady=10, padx=10)
        Label(details_frame, text="Select a strategy to view details or create a new one.").pack()
        self.strategy_listbox.bind('<<ListboxSelect>>',
                                   lambda event: self.display_strategy_details(event, details_frame))

    def on_config_window_close(self):
        self.config_window.destroy()
        self.config_window = None

    def update_strategy_list(self):
        self.strategy_listbox.delete(0, tk.END)
        for strategy in self.strategies:
            self.strategy_listbox.insert(tk.END, strategy.name)

    def display_strategy_details(self, event, details_frame):
        selection = self.strategy_listbox.curselection()
        if selection:
            index = selection[0]

            strategy = self.strategies[index]
            for widget in details_frame.winfo_children():
                widget.destroy()

            Label(details_frame, text="Strategy Name:").pack()
            name_entry = Entry(details_frame, width=30)
            name_entry.insert(0, strategy.name)
            name_entry.pack()

            Label(details_frame, text="Buffers:").pack()
            entries = []
            for i in range(strategy.max_columns):
                sub_frame = Frame(details_frame)
                sub_frame.pack(fill='x', padx=5, pady=2)
                Label(sub_frame, text=f"Column {i + 1} Start Buffer:").pack(side='left')
                start_entry = Entry(sub_frame, width=5)
                start_entry.insert(0, str(strategy.buffers_start[i]))
                start_entry.pack(side='left', padx=10)
                Label(sub_frame, text=f"Column {i + 1} End Buffer:").pack(side='left')
                end_entry = Entry(sub_frame, width=5)
                end_entry.insert(0, str(strategy.buffers_end[i]))
                end_entry.pack(side='left', padx=10)
                entries = []

                if isinstance(strategy, (TelenetColumnStrategy, VOOColumnStrategy, OrangeColumnStrategy)):
                    start_entry.config(state='disabled')
                    end_entry.config(state='disabled')
                entries.append((start_entry, end_entry))
            save_name_button = Button(details_frame, text="Save Name",
                                      command=lambda: self.save_name(details_frame, name_entry, strategy, entries))
            save_name_button.pack(pady=10)

    def save_name(self, details_frame, name_entry, strategy, entries):
        new_name = name_entry.get()
        if strategy.name != new_name:
            strategy.name = new_name
            messagebox.showinfo("Info", "Strategy name updated successfully.")
            self.update_strategy_list()
            self.update_all_strategy_dropdowns()
            if self.config_window is not None:
                self.config_window.destroy()
                self.config_window = None

    def process_all_files(self):
        all_data = []
        for file_path, processed_data, strategy_var in self.files_data:
            strategy_instance = next((s for s in self.strategies if s.name == strategy_var.get()), None)
            if strategy_instance:
                self.pdf_processors[file_path].set_column_strategy(strategy_instance)
                extracted_data = self.pdf_processors[file_path].extract_data(
                    file_path, [p[1] for p in processed_data], [True] * len(processed_data), all_data)
            else:
                print(f"Strategy '{strategy_var.get()}' not found.")
        if all_data:
            self.pdf_processors[next(iter(self.pdf_processors))].export_to_excel(
                all_data)
        else:
            messagebox.showinfo("Info", "No data to export.")
            print("No data was extracted.")

    def remove_file(self, file_path, file_frame):
        if file_path in self.file_labels:
            del self.file_labels[file_path]
        self.files_data = [(path, data, strat) for path, data, strat in self.files_data if path != file_path]
        file_frame.destroy()
        if file_path in self.pdf_processors:
            del self.pdf_processors[file_path]


if __name__ == "__main__":
    root = tk.Tk()
    app = ChannelSynthesizerGUI(root)
    root.mainloop()
