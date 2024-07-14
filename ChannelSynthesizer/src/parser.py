import pdfplumber
import pandas as pd
import os
import re
from tkinter import messagebox, Frame, Label, Spinbox, Checkbutton, IntVar, StringVar
from strategies import ColumnStrategy, TelenetColumnStrategy, VOOColumnStrategy

class PDFProcessor:
    def __init__(self, column_strategy: ColumnStrategy):
        self.column_strategy = column_strategy

    def set_column_strategy(self, strategy: ColumnStrategy):
        self.column_strategy = strategy

    def process_file_gui(self, file_path, window):
        entries = []
        try:
            with pdfplumber.open(file_path) as pdf:
                for page_num, page in enumerate(pdf.pages, start=1):
                    entries += self.create_page_settings(page_num, page, window)
            return entries
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open PDF file: {e}")
            window.destroy()
            return None

    def create_page_settings(self, page_num, page, window):
        frame = Frame(window)
        frame.pack(fill='x', padx=50, pady=5)
        num_columns_var = StringVar(value="2")
        include_page_var = IntVar(value=1)
        Label(frame, text=f'Page {page_num}:').pack(side='left')
        Label(frame, text="Columns:").pack(side='left')
        num_columns_spinbox = Spinbox(frame, from_=1, to=10, textvariable=num_columns_var, width=5)
        num_columns_spinbox.pack(side='left', padx=10)
        checkbutton = Checkbutton(frame, text="Exclude this page?", variable=include_page_var)
        checkbutton.pack(side='left')

        num_columns_spinbox.bind("<Button-1>", lambda *args, var=include_page_var: var.set(0))
        num_columns_spinbox.bind("<Key>", lambda *args, var=include_page_var: var.set(0))

        return [(page_num, num_columns_var, include_page_var)]

    def process_pdf(self, entries, file_path):
        processed_data = []
        num_columns_list = [int(entry[1].get()) for entry in entries]
        page_include = [not bool(entry[2].get()) for entry in entries]
        for entry in entries:
            if not entry[2].get():
                processed_data.append((entry[0], entry[1].get(), True))
        return processed_data if any(page_include) else []

    def define_columns(self, page, num_columns):
        return self.column_strategy.define_columns(page, num_columns)

    def extract_data(self, file_path, num_columns_list, page_include, all_data):
        try:
            with pdfplumber.open(file_path) as pdf:
                for page_number, include in zip(range(1, len(page_include) + 1), page_include):
                    if include:
                        page = pdf.pages[page_number - 1]
                        num_columns = int(num_columns_list[page_number - 1])
                        columns = self.define_columns(page, num_columns)
                        column_index = 1
                        for coords in columns:
                            x0, top, x1, bottom = coords
                            cropped_page = page.crop((x0, top, x1, bottom))
                            text = cropped_page.extract_text(x_tolerance=3, y_tolerance=3)
                            if text:
                                phrases = re.split(r'\.\s+|\n+', text.strip())
                                for phrase in phrases:
                                    if phrase.strip():
                                        all_data.append({
                                            'Filename': os.path.basename(file_path),
                                            'Page': page_number,
                                            'Column': column_index,
                                            'Text': phrase.strip(),
                                            'Top': top,
                                            'Bottom': bottom,
                                            'X0': x0,
                                            'X1': x1
                                        })
                            column_index += 1
        except Exception as e:
            messagebox.showerror("Error", f"Error processing PDF: {e}")

    def export_to_excel(self, all_data):
        output_file = os.path.join(os.path.abspath('../outputs/xlsx'), 'extracted_data.xlsx')
        directory = os.path.dirname(output_file)
        if not os.path.exists(directory):
            os.makedirs(directory, exist_ok=True)
        df = pd.DataFrame(all_data)
        if not df.empty:
            df.to_excel(output_file, index=False)
        else:
            messagebox.showinfo("Info", "No data to export.")
