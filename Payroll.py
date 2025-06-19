#!/usr/bin/env python3
import sys
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
from PIL import Image, ImageTk
import datetime
from copy import copy
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from openpyxl.utils import get_column_letter
import os
from os import path
import sys
from processFiles import processFiles

class Payroll:

    def __init__(self, root):
        self.root = root
        self.root.title("Excel File Processor")
        self.root.geometry("800x600")
        self.label_width = 25

        if hasattr(sys, '_MEIPASS'):
            img_path = os.path.join(sys._MEIPASS, "background.png")
        else:
            img_path = os.path.abspath("background.png")

        bg_image = Image.open(img_path)
        bg_image = bg_image.resize((800, 600) )
        self.bg_photo = ImageTk.PhotoImage(bg_image)

        # Create a canvas to display the background image
        canvas = tk.Canvas(root, width=800, height=600)
        canvas.pack()
        canvas.create_image(0, 0, anchor=tk.NW, image=self.bg_photo)

        # Title label
        # title_label = tk.Label(root, text="Processing Excel Files Made Easy", font=("Helvetica", 20), bg="lightblue")
        # canvas.create_window(400, 50, window=title_label)

        # File selections
        self.file1_var = tk.StringVar()
        self.file2_var = tk.StringVar()

        file1_label = tk.Label(root, text="INI File:", width=self.label_width )
        canvas.create_window(100, 540, window=file1_label)

        file1_entry = tk.Entry(root, textvariable=self.file1_var,  width=50)
        canvas.create_window(350, 540, window=file1_entry)

        file1_button = tk.Button(root, text="Select", command=self.select_file1)
        canvas.create_window(550, 540, window=file1_button)

        self.file2_label = tk.Label(root, text="Excel file:", width=self.label_width)
        canvas.create_window(100, 570, window=self.file2_label)

        file2_entry = tk.Entry(root, textvariable=self.file2_var, width=50)
        canvas.create_window(350, 570, window=file2_entry)

        file2_button = tk.Button(root, text="Select", command=self.select_file2)
        canvas.create_window(550, 570, window=file2_button)

        # Month selection
        month_label = tk.Label(root, text="Select Month:", width=self.label_width)
        canvas.create_window(100, 480, window=month_label)

        self.month_var = tk.StringVar()
        month_combobox = ttk.Combobox(root, textvariable=self.month_var, values=[
            "Jan", "Feb", "Mar", "Apr", "May", "Jun",
            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
        ])
        canvas.create_window(270, 480, window=month_combobox)

        # Year selection
        year_label = tk.Label(root, text="Select Year:", width=self.label_width)
        canvas.create_window(100, 510, window=year_label)

        self.year_var = tk.StringVar()
        self.update_years()  # Populate the year selection
        year_combobox = ttk.Combobox(root, textvariable=self.year_var, values=self.years)
        canvas.create_window(270, 510, window=year_combobox)

        # Process button
        process_button = tk.Button(root, text="Process", command=self.process_data)
        canvas.create_window(650, 570, window=process_button)

        # Exit Button
        exit_button = tk.Button(root, text="Exit",    command=root.quit)
        canvas.create_window(750, 570, window=exit_button)

    def update_years(self):
        current_year = datetime.datetime.now().year
        self.years = [str(current_year - 1), str(current_year), str(current_year + 1)]

    def select_file1(self):
        file_path = filedialog.askopenfilename(title="Select File 1")
        self.file1_var.set(file_path)

    def select_file2(self):
        file_path = filedialog.askopenfilename(title="Select File 2")
        self.file2_var.set(file_path)

    def process_data(self):
        defnfilename = self.file1_var.get()
        exclfilename = self.file2_var.get()
        cldrmnth     = self.month_var.get()
        cldryear     = self.year_var.get()
    
        result = processFiles(defnfilename,exclfilename,cldrmnth,cldryear,False)

        messagebox.showinfo("Done",result)

if __name__ == "__main__":
    root = tk.Tk()
    app = Payroll(root)
    root.mainloop()
