from tkinter import ttk

class PDFTab(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        ttk.Label(self, text="PDF Operations").pack(pady=20)
