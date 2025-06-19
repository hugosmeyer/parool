#!/usr/bin/env python3
import sys
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
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
from collections import defaultdict
from openpyxl.styles import Alignment


def processFiles(defnfilename,exclfilename,cldrmnth,cldryear,debugActive):
    def debug(*args):
        if debugActive:
            print(*args)
    #try:
    datafilefldr = path.dirname(exclfilename)
    newxfilename = path.join(datafilefldr, str(path.basename(exclfilename.replace(".xlsx", " Schedules.xlsx"))))

    debug("Running in folder      :", os.getcwd())
    debug("Selected Month         :", cldrmnth)
    debug("Selected Year          :", cldryear)
    debug("Data Folder            :", datafilefldr)
    debug("Definition file(ini)   :", defnfilename)
    debug("Excel Source file      :", os.path.basename(exclfilename))
    debug("New Main Excel file    :", os.path.basename(newxfilename))
    debug("New files created here :", datafilefldr)
    if 1 == 1:
        # Check if a file exists
        def check_file(filename):
            try:
                filehndl = open(filename,"r")
            except IOError:
                return False, "Unable to open type file: " + filename
            else:
                filehndl.close()
                return True, ""

        def defnfileprse(file_path):
            ini_data = dict()

            with open(file_path, 'r') as file:
                currsect = None
                for line in file:
                    line = line.strip()
                    if line.startswith('[') and line.endswith(']'):
                        currsect = line[1:-1]
                        ini_data[currsect] = list()
                    elif '=' in line:
                        key, value = map(str.strip, line.split('=', 1))
                        if not key == "":
                            if key is None:
                                key = ""
                            ini_data[currsect].append([key, value])
            return ini_data

        # Define the underline formats for Total headers and values
        defnbrdrthin = Side(border_style="thin",  color="000000")
        defnbrdrthck = Side(border_style="thick", color="000000")
        cellbrdrthin = Border(bottom=defnbrdrthin)
        cellbrdrthck = Border(bottom=defnbrdrthck)

        def copycellvalu(fromcell, destcell):
            destcell.value = fromcell.value

        def copycellfrmt(fromcell, destcell):
            destcell.font          = copy(fromcell.font)
            destcell.border        = copy(fromcell.border)
            destcell.fill          = copy(fromcell.fill)
            destcell.number_format = copy(fromcell.number_format)
            destcell.protection    = copy(fromcell.protection)
            destcell.alignment     = copy(fromcell.alignment)

        def makefontbold(cell):
            thisfont = copy(cell.font)
            thisfont.bold = True
            cell.font = thisfont

        def fontsizenrml(cell):
            cell.font       = Font(name='Arial',size=11,bold=None,italic=False,vertAlign=None,underline=None,strike=False,color='FF000000')

        def fontsizelrge(cell):
            cell.font       = Font(name='Arial',size=14,bold=None,italic=False,vertAlign=None,underline=None,strike=False,color='FF000000')

        def maketextcntr(cell):
            cell.alignment = Alignment(horizontal='general',vertical='center',text_rotation=0,wrap_text=True,shrink_to_fit=False,indent=0)

        def fillcellcolr(cell):
            cell.fill   = PatternFill(fill_type="lightGray",start_color='FFCCCCCC',end_color='FFCCCCCC')

        def frmttotltitl(cell):
            fontsizelrge(cell)
            cell.alignment = Alignment(horizontal='general',vertical='bottom',text_rotation=0,wrap_text=True,shrink_to_fit=False,indent=0)
            fillcellcolr(cell)

        def frmttotlvalu(cell):
            fontsizelrge(cell)
            cell.border        = cellbrdrthck
            cell.number_format = "0.00"
            cellabov           = cell.parent.cell(row=cell.row - 1, column=cell.column)
            cellabov.border    = cellbrdrthin 



        # Main, start of the program

        #Assume all files are present
        filepass = True
   
        # Check if definition file exists
        filechckbool, filechckrslt = check_file(defnfilename)
        if not filechckbool:
            filepass = False

        # Check if Excel file exists
        filechckbool, filechckrslt = check_file(exclfilename)
        if not filechckbool:
            filepass = False

        if not filepass:
            return("Failed","Invalid files selected")

        # Use the definition file's name as business unit name
        busnunitname   = os.path.basename(defnfilename).split(".")[0]






        # Parse the definition file
        defndict = defnfileprse(defnfilename)

        # Open the input Excel sheet and request the result of instead of
        # the formulas itself and take some measurements.
        exclmainbook = load_workbook(exclfilename,data_only=True)
        exclmainshet = exclmainbook[exclmainbook.sheetnames[0]]
        exclrowsused = exclmainshet.max_row
        exclcolsused = exclmainshet.max_column
        
        # Make a list of column headers from the input sheet
        colmcntr = 1
        maincolmhdrs = {}
        while colmcntr <= exclcolsused:
            maincolmhdrs[exclmainshet.cell(row = 1, column = colmcntr).value] = colmcntr
            colmcntr += 1

        # Determine the font in A1 of the input sheet
        cellfontrrefr = exclmainshet.cell(row = 1, column = 1)

        # Process each section in the INI file
        for defn in defndict:
            thisdefn = defndict[defn]
            
            # Space needed for the Totals at the top.
            rowsstrt = 5
   
            # Create a list of input and output column mappings where the
            # input column exists.

            nzrocols = []
            totlcols = []
            repldefn = list()
            colmcntr = 1
            for maincolm, thiscolm in thisdefn:

                # Ignore columns that cannot be found in the input sheet.
                if maincolm not in maincolmhdrs:
                    continue

                # _NZ_ will have rows with an empty or zero value in this column to to be excluded
                if "_NZ_" in thiscolm:
                    nzrocols.append(colmcntr)
                    thiscolm = thiscolm.replace("_NZ_", "").strip()

                # _SUM_ will have the values in this column added up and shown below
                if "_SUM_" in thiscolm:
                    totlcols.append(colmcntr)
                    thiscolm = thiscolm.replace("_SUM_", "").strip()
                colmcntr += 1

                repldefn.append([maincolm, thiscolm])

            thisdefn = repldefn
            del repldefn
   
            # Definition names end in .TAB or .FILE 
            #   .TAB will result in a new tab in a copy of the Input Excel
            #   .FILE will result in a new Excel file wit only the data in
            #   this definition in it.
            defnname, defntype = defn.split(".")

            # Create a sheet in which to work
            if defntype == "FILE":
                destbook = Workbook()
                destshet = destbook.active
            else:
                destshet = exclmainbook.create_sheet(title = defnname)

            # Give it a title 
            destshet.title = defnname

            # Add the title to A1
            titlcell = destshet.cell(row = 1, column = 1)
            fontsizelrge(titlcell)
            makefontbold(titlcell)
            titlcell.value = busnunitname + " - " + defnname + " - " + cldrmnth + " " + cldryear
            titlcell.alignment = Alignment(horizontal='general',vertical='center',text_rotation=0,wrap_text=False,shrink_to_fit=False,indent=0)

            destshet.row_dimensions[1].height = None

            chdralgn = alignment=Alignment(horizontal='general',vertical='bottom',text_rotation=0,wrap_text=True,shrink_to_fit=False,indent=0)

            destcolm = 1
            dlterows = []

            # Create the Column headers
            for maincolm,thiscolm in thisdefn:
                # Insert the header
                destcell = destshet.cell(row = rowsstrt, column = destcolm)
                destcell.value = thiscolm
                frmttotltitl(destcell)

                # If it is a total column, add the header above also
                if destcolm in totlcols:
                    destcell = destshet.cell(row = 1, column = destcolm)
                    destcell.value = thiscolm
                    frmttotltitl(destcell)

                # Copy the cells below the column headers
                rowscntr = 2
                while rowscntr <= exclrowsused:
                    destcell = destshet.cell(row = rowscntr + rowsstrt - 1, column = destcolm)
                    destshet.row_dimensions[rowscntr].height = None
                    fromcell=exclmainshet.cell(row = rowscntr, column = maincolmhdrs[maincolm])
                    destcell.value = fromcell.value 
                    if fromcell.has_style:
                        copycellfrmt(fromcell, destcell)
                    if destcolm in nzrocols:
                        if isinstance(destcell.value, (int, float)) and not isinstance(destcell.value, bool):
                            if abs(destcell.value) < 1e-12:
                                if destcell.row not in dlterows:
                                    dlterows.append(destcell.row)
                        elif isinstance(destcell.value, str):
                            if destcell.value.strip()  == "":
                                if destcell.row not in dlterows:
                                    dlterows.append(destcell.row)
                        elif destcell.value is None:
                            if destcell.row not in dlterows:
                                dlterows.append(destcell.row)
                    rowscntr += 1
                destcolm += 1

            if len(totlcols) > 1:
                destcell = destshet.cell(row = 5, column = destcolm)
                destcell.value = "Total"
                frmttotltitl(destcell)

            # Remove the zero value ones _NZ_ columns from the sheet
            dlterows.sort(reverse = True)
            if dlterows:
                for dlterown in dlterows:
                    destshet.delete_rows(dlterown,1)
                    print(destshet.max_row)

            # Add up the totals if there are any
            if totlcols:
                rowsused = destshet.max_row
                colsused = destshet.max_column
                colmnmbr = 1
                lasttotlvalu = 0

                totllinecels = defaultdict(list)
                totlglobcels = []

                while colmnmbr <= colsused:

                    totlcellblow = destshet.cell(row = rowsused + 1, column = colmnmbr)

                    # Colour all the cells in the total line
                    fillcellcolr(totlcellblow)

                    if colmnmbr in totlcols:
                        frmttotlvalu(totlcellblow)

                        totlcellabov = destshet.cell(row = 2           , column = colmnmbr)
                        frmttotlvalu(totlcellabov)

                        ctrlcellabov = destshet.cell(row = 3           , column = colmnmbr)
                        fontsizelrge(ctrlcellabov)

                        rowscntr = rowsstrt + 1
                        totlcolmcels = defaultdict(list)
                        while rowscntr <= rowsused:
                            valucell = destshet.cell(row = rowscntr, column = colmnmbr)
                            if type(valucell.value) in [int, float]:
                                totlcolmcels[colmnmbr].append(valucell.coordinate)
                                totllinecels[rowscntr].append(valucell.coordinate)
                                totlglobcels.append(valucell.coordinate)
                                lasttotlvalu += valucell.value
                            rowscntr += 1

                        # load the cell value into bottom and above cells
                        totlcellblow.value = "=SUM(" + totlcolmcels[colmnmbr][0] + ":" + totlcolmcels[colmnmbr][-1] + ")"
                        totlcellabov.value = "=SUM(" + totlcolmcels[colmnmbr][0] + ":" + totlcolmcels[colmnmbr][-1] + ")"
                        ctrlcellabov.value = "=IF(" + totlcellblow.coordinate+ "=" + totlcellabov.coordinate + ",TRUE,FALSE)"
                        
                    if colmnmbr == 1:
                        totlcellblow.value = "Grand Total"
                        frmttotltitl(totlcellblow)

                    colmnmbr += 1

                if len(totlcols) > 1:
                    sidetotlcolm = colsused 
                    sidetotlcell = destshet.cell(row = 1, column = sidetotlcolm)
                    sidetotlcell.value = "Total"
                    frmttotltitl(sidetotlcell)

                    sidetotlcell = destshet.cell(row = 2, column = sidetotlcolm)
                    sidetotlcell.value = "=SUM(" + ",".join(totlglobcels) + ")"
                    frmttotlvalu(sidetotlcell)


            rowscntr = rowsstrt + 1 
            rowslast = destshet.max_row - 1
            colmlast = destshet.max_column 
            sidetotllist = []
            print("totlcols=",totlcols)
            
            while rowscntr <= rowslast:
                sidetotlcels = []
                for colmnmbr in totlcols:
                    thiscell = destshet.cell(row = rowscntr, column = colmnmbr)
                    print("++++")
                    sidetotlcels.append(thiscell.coordinate)

                print("nnn",sidetotlcels)
                thiscell = destshet.cell(row = rowscntr, column = colmlast)
                thiscell.value = "=SUM(" + ",".join(sidetotlcels) + ")"
                thiscell.number_format = "0.00"
                sidetotllist.append(thiscell.coordinate)
                rowscntr += 1

            if len(totlcols) > 1:
                # Bottom total on the right
                thiscell = destshet.cell(row = rowscntr, column = colmlast)
                thiscell.value = "=sum(" + sidetotllist[0] + ":" + sidetotllist[-1] + ")"
                frmttotlvalu(thiscell)

                # Control at top right
                sidectrlcell = destshet.cell( row = 3, column=sidetotlcolm)
                sidectrlcell.value = "=IF(" + sidetotlcell.coordinate+ "=" + thiscell.coordinate + ",TRUE,FALSE)"
                fontsizelrge(sidectrlcell)


            # Insert the pesky "Total" before the first total column
            thiscell = destshet.cell( row = 2, column = totlcols[0] - 1)
            thiscell.value = "Total"
            fontsizelrge(thiscell)
            makefontbold(thiscell)

            # Merge the cells for the title on the left.
            destshet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=totlcols[0] - 2)

            # Make all the columns auto sizing
            for cols in destshet.columns:
                col = get_column_letter(cols[0].column)
                destshet.column_dimensions[col].auto_size = True

            # Save it if it is a file type
            if defntype == "FILE":
                destbook.save(filename=os.path.join(datafilefldr, busnunitname + " " + defnname + " " + str(cldrmnth) + " " + str(cldryear) + ".xlsx"))

        # Save the main sheet anyway
        exclmainbook.save(newxfilename)

    #except Exception as e:
    #    status = "Failed"
    #    result = e
    #else:
        status = "Success"
        result = ""

    return(status, result)
