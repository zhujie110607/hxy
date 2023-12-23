import tkinter as tk
from tkinter import filedialog
import os


def select_excel_file(prompt_message):
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(filetypes=[(prompt_message, '*.xlsx')])

    if file_path:
        return os.path.abspath(file_path)
    else:
        return None
