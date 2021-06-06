import tkinter as tk
from tkinter import ttk

class Properties(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        ttk.Frame.__init__(self, parent , *args, **kwargs)
        self.parent_frame = parent

        self.lbl_empty = tk.Label(self.parent_frame, text = "", height = 10)
        self.lbl_empty.grid(row = 0, column = 0, columnspan = 2)
        self.excel_open()

    def excel_open(self):
        self.lbl_session = tk.Label(self.parent_frame, text="Session Name:")
        self.entry_session = tk.Entry(self.parent_frame, width = 30)

        self.lbl_link = tk.Label(self.parent_frame, text="Excel Link:")
        self.entry_link = tk.Entry(self.parent_frame, width = 30)

        self.lbl_workbook = tk.Label(self.parent_frame, text="Select Workbook:")
        self.entry_workbook= tk.Entry(self.parent_frame, width = 30)

        self.lbl_session.grid(row = 1, column = 0, padx = 10, pady = 10)
        self.entry_session.grid(row = 1, column = 1, padx = 10, pady = 10)
        self.lbl_link.grid(row = 2, column = 0, padx = 10, pady = 10)
        self.entry_link.grid(row = 2, column = 1, padx = 10, pady = 10)
        self.lbl_workbook.grid(row = 3, column = 0, padx = 10, pady = 10)
        self.entry_workbook.grid(row = 3, column = 1, padx = 10, pady = 10)