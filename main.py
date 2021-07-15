from tkinter import Tk, ttk
from tkinter.constants import BOTH
from typing import Text
from lib.repeated_barcode import GUI
from lib.cost import CostGUI

root = Tk()
root.title("Dornika Market")
root.geometry("700x500+10+10")
widget = GUI()


notebook = ttk.Notebook(root)
notebook.pack(fill=BOTH, expand=True)

f = ttk.Frame(notebook)
notebook.add(f, text="انتخاب فایل ها")
widget.load_files(f)

f1 = ttk.Frame(notebook)
notebook.add(f1, text = "مشخص کردن ستون ها")
widget.set_column(f1)

f2 = ttk.Frame(notebook)
notebook.add(f2, text="جدا سازی")
widget.seprate_file_gui(f2)


f3 = ttk.Frame(notebook)
notebook.add(f3, text="گزارش گیری")
cost_widget = CostGUI(f3)

root.mainloop()