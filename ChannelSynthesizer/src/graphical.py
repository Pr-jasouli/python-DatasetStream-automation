import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, Label, Button, Spinbox, Checkbutton
import os
from parser import PDFProcessor


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
        self.pdf_processor = PDFProcessor()
        self.file_labels = {}

    def add_file(self):
        file_path = filedialog.askopenfilename(title='Select a PDF file', filetypes=[("PDF files", "*.pdf")])
        if file_path:
            self.process_file(file_path)

    def process_file(self, file_path):
        file_settings_window = Toplevel(self.root)
        file_settings_window.title("Configure File")
        entries = self.pdf_processor.process_file_gui(file_path, file_settings_window)
        if entries is None:
            return

        def on_done():
            num_columns_list = [int(entry[1].get()) for entry in entries]
            page_include = [not bool(entry[2].get()) for entry in entries]
            processed_data = [(entry[0], entry[1].get(), entry[2].get() == 0) for entry in entries if
                              entry[2].get() == 0]
            if any(page_include):
                self.files_data.append((os.path.abspath(file_path), num_columns_list, page_include))
                self.display_file_data(file_path, processed_data)
            file_settings_window.destroy()

        done_button = Button(file_settings_window, text="Done", command=on_done)
        done_button.pack(pady=20)

    def display_file_data(self, file_path, processed_data):
        if processed_data:
            display_text = os.path.basename(file_path) + ": " + ", ".join(
                f"Page {p} ({c})" for p, c, _ in processed_data)
            file_frame = tk.Frame(self.files_frame)
            file_frame.pack(fill=tk.X, padx=5, pady=5)

            remove_button = Button(file_frame, text="❌", command=lambda: self.remove_file(file_path, file_frame))
            remove_button.pack(side=tk.LEFT, padx=5)

            label = Label(file_frame, text=display_text)
            label.pack(side=tk.LEFT, fill=tk.X, expand=True)

            self.file_labels[file_path] = file_frame

    def process_all_files(self):
        print("Processing these files:",
              [os.path.basename(f[0]) for f in self.files_data])

        all_data = []
        for file_path, num_columns_list, page_include in self.files_data:
            self.pdf_processor.extract_data(file_path, num_columns_list, page_include, all_data)
        if all_data:
            self.pdf_processor.export_to_excel(all_data)
        else:
            messagebox.showinfo("Info", "No data to export.")

    def remove_file(self, file_path, file_frame):
        original_count = len(self.files_data)
        self.files_data = [(path, num_columns, pages) for path, num_columns, pages in self.files_data if
                           path != os.path.abspath(file_path)]
        updated_count = len(self.files_data)

        print(f"Removed: {original_count - updated_count} entry. Total now: {updated_count}")

        file_frame.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = ChannelSynthesizerGUI(root)
    root.mainloop()
