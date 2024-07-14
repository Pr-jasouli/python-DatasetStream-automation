import tkinter as tk
from tkinter import filedialog, messagebox, Label, Button, OptionMenu, StringVar, Frame
import os
from parser2 import PDFProcessor
from strategies import TelenetColumnStrategy, VOOColumnStrategy


class ChannelSynthesizerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Processing Application")
        self.files_frame = tk.Frame(self.root)
        self.files_frame.pack(padx=10, pady=10, expand=True, fill=tk.BOTH)
        self.add_file_button = Button(self.root, text="Load PDF File", command=self.add_file)
        self.add_file_button.pack(pady=10)
        self.process_button = Button(self.root, text="Process All Files", command=self.process_all_files)
        self.process_button.pack(pady=10)
        self.files_data = []
        self.file_labels = {}

    def add_file(self):
        file_path = filedialog.askopenfilename(title='Select a PDF file', filetypes=[("PDF files", "*.pdf")])
        if file_path:
            self.configure_file(file_path)

    def configure_file(self, file_path):
        file_settings_window = tk.Toplevel(self.root)
        file_settings_window.title("Configure File Settings")
        default_strategy = self.determine_strategy(file_path)
        self.pdf_processor = PDFProcessor(default_strategy)
        entries = self.pdf_processor.process_file_gui(file_path, file_settings_window)

        def on_done():
            if entries:
                num_columns_list = [int(entry[1].get()) for entry in entries]
                page_include = [not bool(entry[2].get()) for entry in entries]
                if any(page_include):
                    processed_data = self.pdf_processor.process_pdf(entries, file_path)
                    if processed_data:
                        self.display_file_data(file_path, processed_data,
                                               default_strategy.__class__.__name__.replace('ColumnStrategy', ''))
            file_settings_window.destroy()

        done_button = Button(file_settings_window, text="Done", command=on_done)
        done_button.pack(pady=20)

    def determine_strategy(self, file_path):
        filename = os.path.basename(file_path).lower()
        if "telenet" in filename:
            return TelenetColumnStrategy()
        elif "voo" in filename:
            return VOOColumnStrategy()
        return TelenetColumnStrategy()

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
        options = ['Telenet', 'VOO']
        strategy_menu = OptionMenu(file_frame, strategy_var, *options)
        strategy_menu.pack(side=tk.RIGHT, padx=5)

        self.file_labels[file_path] = file_frame
        self.files_data.append((os.path.abspath(file_path), processed_data, strategy_var))

    def process_all_files(self):
        all_data = []
        for file_path, processed_data, strategy_var in self.files_data:
            strategy_name = strategy_var.get()
            if strategy_name == "Telenet":
                strategy_instance = TelenetColumnStrategy()
            else:
                strategy_instance = VOOColumnStrategy()
            self.pdf_processor.set_column_strategy(strategy_instance)
            self.pdf_processor.extract_data(file_path, [p[1] for p in processed_data], [True] * len(processed_data),
                                            all_data)

        if all_data:
            self.pdf_processor.export_to_excel(all_data)
        else:
            messagebox.showinfo("Info", "No data to export.")

    def remove_file(self, file_path, file_frame):
        self.files_data = [(path, data, strat) for path, data, strat in self.files_data if
                           path != os.path.abspath(file_path)]
        file_frame.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = ChannelSynthesizerGUI(root)
    root.mainloop()
