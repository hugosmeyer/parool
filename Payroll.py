#!/usr/bin/env python3
import sys
import os
from os import path
import datetime
import configparser
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
from PIL import Image, ImageTk
from processFiles import processFiles

class Payroll:
    """Main GUI application for Excel payroll file processing.
    
    This class creates a Tkinter-based GUI that allows users to:
    - Select recipe definition files (INI format)
    - Choose company, month, and year
    - Select payroll Excel files
    - Process files to generate formatted output sheets
    """

    def __init__(self, root):
        """Initialize the Payroll GUI application.
        
        Args:
            root: The Tkinter root window
        """
        self.root = root
        self.root.title("Excel File Processor")
        self.root.geometry("800x600")
        
        self.root.resizable(False, False)

        # Load background image with error handling
        try:
            if hasattr(sys, '_MEIPASS'):
                img_path = os.path.join(sys._MEIPASS, "background.png")
            else:
                img_path = os.path.abspath("background.png")

            bg_image = Image.open(img_path)
            bg_image = bg_image.resize((800, 600))
            self.bg_photo = ImageTk.PhotoImage(bg_image)
        except FileNotFoundError:
            print(f"Warning: Background image not found at {img_path}")
            self.bg_photo = None
        except Exception as e:
            print(f"Warning: Could not load background image: {e}")
            self.bg_photo = None

        #
        # Create a canvas to display the background image
        #
        canvas = tk.Canvas(root, width=800, height=600)
        canvas.pack()
        if self.bg_photo:
            canvas.create_image(0, 0, anchor=tk.NW, image=self.bg_photo)

        self.dropdownwdth = 100
        self.filenamewdth = 400
        self.lablwdth     = 100
        self.butnwdth     = 55
        self.ypos     = 440
        self.yposincr = 30

        self.lablxpos = 10
        self.valuxpos = self.lablxpos + self.lablwdth + 10
        self.butnxpos = self.valuxpos + self.filenamewdth + 10
        # 
        # Recipe
        #
        self.rcpeflnmvalu = tk.StringVar()
        rcpeflnmlabl = tk.Label(root, text="Recipe File:", anchor="w")
        canvas.create_window(self.lablxpos, self.ypos, window=rcpeflnmlabl, width=self.lablwdth, anchor="nw" )

        self.rcpeflnmentr = tk.Entry(root, textvariable=self.rcpeflnmvalu)
        canvas.create_window(self.valuxpos, self.ypos, window=self.rcpeflnmentr, width=self.filenamewdth, anchor="nw")

        rcpeflnmbutn = tk.Button(root, text="Select", command=self.rcpeflnmslct)
        canvas.create_window(self.butnxpos, self.ypos, window=rcpeflnmbutn , anchor="nw", width=self.butnwdth)

        self.ypos += self.yposincr
        
        #
        # Company
        #
        compnamelabl = tk.Label(root, text="Company:", anchor="w")
        canvas.create_window(self.lablxpos, self.ypos, window=compnamelabl,width = self.lablwdth, anchor="nw")

        self.compnamevalu = tk.StringVar()
        self.compdropdown = ttk.Combobox(root, textvariable=self.compnamevalu, values=[])
        self.compdropdown.configure(state="disabled")
        canvas.create_window(self.valuxpos, self.ypos, window=self.compdropdown, width=self.dropdownwdth, anchor="nw")
     
        self.ypos += self.yposincr

        #
        # Month selection
        #
        mnthlabl = tk.Label(root, text="Month:", anchor="w")
        canvas.create_window(self.lablxpos, self.ypos, window=mnthlabl,width = self.lablwdth, anchor="nw")

        self.mnthnamevalu = tk.StringVar()
        self.mnthdropdown = ttk.Combobox(root, textvariable=self.mnthnamevalu, values=[
            "Jan", "Feb", "Mar", "Apr", "May", "Jun",
            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
        ])
        canvas.create_window(self.valuxpos, self.ypos, window=self.mnthdropdown, width=self.dropdownwdth, anchor="nw")

        self.ypos += self.yposincr

        #
        # Year selection
        #
        year_label = tk.Label(root, text="Year:", anchor="w")
        canvas.create_window(self.lablxpos, self.ypos, window=year_label,width = self.lablwdth, anchor="nw")

        self.yearnamevalu = tk.StringVar()
        self.yeardropdown = ttk.Combobox(root, textvariable=self.yearnamevalu, values=self.yearvalulist())
        canvas.create_window(self.valuxpos, self.ypos, window=self.yeardropdown, width=self.dropdownwdth, anchor="nw")

        self.ypos += self.yposincr

        #
        # Payroll File
        #
        self.prolflnmvalu = tk.StringVar()
        self.prollabl     = tk.Label(root, text="Payroll File:", anchor="w")
        canvas.create_window(self.lablxpos, self.ypos, window=self.prollabl, width=self.lablwdth , anchor="nw")

        prolflnmentr = tk.Entry(root, textvariable=self.prolflnmvalu)
        canvas.create_window(self.valuxpos, self.ypos, window=prolflnmentr, width=self.filenamewdth, anchor="nw")

        prolflnmbutn = tk.Button(root, text="Select", command=self.prolflnmslct)
        canvas.create_window(self.butnxpos, self.ypos, window=prolflnmbutn, anchor="nw", width=self.butnwdth)
                
        self.butnxpos += self.butnwdth + 10
        #
        # Process button
        #
        process_button = tk.Button(root, text="Process", command=self.process_data)
        canvas.create_window(self.butnxpos, self.ypos, window=process_button, anchor="nw", width=self.butnwdth)
        
        self.butnxpos += self.butnwdth + 10
        #
        # Exit Button
        #
        exit_button = tk.Button(root, text="Exit",    command=root.quit)
        canvas.create_window(self.butnxpos, self.ypos, window=exit_button, anchor="nw", width=self.butnwdth)

    
    def rcpeflnmslct(self):
        """Handle recipe file selection and populate company dropdown.
        
        Opens a file dialog to select an INI file, parses it, and extracts
        the list of companies to populate the company dropdown.
        """
        flnmslct = filedialog.askopenfilename(
            title="Select Recipe File", 
            filetypes=[("INI files", "*.ini"), ("All files", "*.*")]
        )
        if flnmslct:
            self.rcpeflnmvalu.set(flnmslct)
            
            # Parse INI file
            config = configparser.ConfigParser(delimiters=('='))
            config.read(flnmslct)

            self.complist = []
            if "COMPANIES" in config:
                for compkeyn in config["COMPANIES"].keys():
                    self.complist.append(config["COMPANIES"].get(compkeyn))
                self.compdropdown.config(values=self.complist)
                self.compdropdown.configure(state="readonly")
                    
    def prolflnmslct(self):
        """Handle payroll Excel file selection.
        
        Opens a file dialog to select an Excel (.xlsx) file.
        """
        file_path = filedialog.askopenfilename(
            title="Select Payroll Excel File", 
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if file_path:
            self.prolflnmvalu.set(file_path)


    def yearvalulist(self):
        """Generate a list of years (previous, current, next).
        
        Returns:
            list: List of year strings [previous year, current year, next year]
        """
        current_year = datetime.datetime.now().year
        return [str(current_year - 1), str(current_year), str(current_year + 1)]
    
    
    def process_data(self):
        """Process the selected files and generate output sheets."""
        # Validate inputs
        defnfilename = self.rcpeflnmvalu.get()
        compname     = self.compnamevalu.get()
        cldrmnth     = self.mnthnamevalu.get()
        cldryear     = self.yearnamevalu.get()
        exclfilename = self.prolflnmvalu.get()
        
        # Validate all required fields are filled
        if not defnfilename:
            messagebox.showerror("Error", "Please select a Recipe file")
            return
        if not compname:
            messagebox.showerror("Error", "Please select a Company")
            return
        if not cldrmnth:
            messagebox.showerror("Error", "Please select a Month")
            return
        if not cldryear:
            messagebox.showerror("Error", "Please select a Year")
            return
        if not exclfilename:
            messagebox.showerror("Error", "Please select a Payroll file")
            return
    
        # Process files and handle result
        status, result = processFiles(defnfilename, exclfilename, cldrmnth, cldryear, compname, False)
        
        if status == "Failed":
            messagebox.showerror("Error", f"Processing failed: {result}")
        else:
            messagebox.showinfo("Success", "Processing completed successfully!")

if __name__ == "__main__":
    root = tk.Tk()
    app = Payroll(root)
    root.mainloop()
