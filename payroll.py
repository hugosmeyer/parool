#!/usr/bin/python3

from copy import copy
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from openpyxl.utils import get_column_letter
import os
import sys
import configparser


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
        current_section = None
        for line in file:
            line = line.strip()
            if line.startswith('[') and line.endswith(']'):
                current_section = line[1:-1]
                ini_data[current_section] = list()
            elif '=' in line:
                key, value = map(str.strip, line.split('=', 1))
                if not key == "":
                    if key is None:
                       key = ""
                    ini_data[current_section].append([key, value])
    return ini_data


# Main, start of the program
if __name__ == "__main__":

    #Assume all files are present
    filepass = True

    # Check if definition file exists
    defnfilename = sys.argv[1]
    filechckbool, filechckrslt = check_file(defnfilename)
    if not filechckbool:
        filepass = False
    
    # Check if Excel file exists
    exclfilename = sys.argv[2]
    filechckbool, filechckrslt = check_file(exclfilename)
    if not filechckbool:
        filepass = False
    
if not filepass:
    sys.exit(1)

# Parse the definition file
defndict = defnfileprse(defnfilename)

exclmainbook = load_workbook(exclfilename,data_only=True)

exclmainname = exclmainbook.sheetnames[0]
exclmainshet = exclmainbook[exclmainbook.sheetnames[0]]
exclrowsused = exclmainshet.max_row
#print("max_row = ",exclrowsused)
exclcolsused = exclmainshet.max_column
#print("max_col = ",exclcolsused)

colmcntr = 1
maincolmhdrs = {}
while colmcntr <= exclcolsused:
    maincolmhdrs[exclmainshet.cell(row = 1, column = colmcntr).value] = colmcntr
    colmcntr += 1

for defn in defndict:
    defnname, defntype = defn.split(".")
    if defntype == "FILE":
        destbook = Workbook()
        destshet = destbook.active
        destshet.title = defnname
    else:
        destshet = exclmainbook.create_sheet(title = defnname)
    
    chdrfont = Font(name='Arial',size=10,bold=True,italic=False,vertAlign=None,underline='none',strike=False,color='FF000000')
    chdrfill = PatternFill(fill_type="lightGray",start_color='FF555555',end_color='FF555555')
    chdralgn = alignment=Alignment(horizontal='general',vertical='center',text_rotation=0,wrap_text=True,shrink_to_fit=False,indent=0)
    destcolm = 1
    for defncolm in defndict[defn]:
        maincolm,thiscolm = defncolm
        
        # insert the header
        destcell = destshet.cell(row = 1, column = destcolm)
        destshet.row_dimensions[1].height = None
        if maincolm in maincolmhdrs:
            fromcell=exclmainshet.cell(row = 1, column = maincolmhdrs[maincolm])
            destcell.value = thiscolm
        else: 
            destcell.value = maincolm
        destcell.font      = chdrfont
        destcell.fill      = chdrfill
        destcell.alignment = chdralgn

        # load the values below the header
        rowscntr = 2
        while rowscntr <= exclrowsused:
            #print("maincolm = ",maincolm)
            #print("rowscntr = ",rowscntr)
            #print("destcolm = ",destcolm)
            destcell = destshet.cell(row = rowscntr, column = destcolm)
            destshet.row_dimensions[rowscntr].height = None
            
            #print (maincolm ,":", maincolm in maincolmhdrs)
            if maincolm in maincolmhdrs:
                #print("fromcolm = ",maincolmhdrs[maincolm])
                fromcell=exclmainshet.cell(row = rowscntr, column = maincolmhdrs[maincolm])
                #print("fromcell.value = ",fromcell.value)
                if fromcell.has_style:
                    destcell.font          = copy(fromcell.font)
                    destcell.border        = copy(fromcell.border)
                    destcell.fill          = copy(fromcell.fill)
                    destcell.number_format = copy(fromcell.number_format)
                    destcell.protection    = copy(fromcell.protection)
                    destcell.alignment     = copy(fromcell.alignment)
                    
                destcell.value = fromcell.value 
            else:
                destcell.value = thiscolm
            rowscntr += 1
        destcolm += 1

    for cols in destshet.columns:
        col = get_column_letter(cols[0].column)
        destshet.column_dimensions[col].auto_size = True

    if defntype == "FILE":
        destbook.save(filename=defnname+".xlsx")

exclmainbook.save(filename="_"+exclfilename)
sys.exit()