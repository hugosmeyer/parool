# Makefile for generating and viewing Excel data in terminal

EXCEL_FILE=devtest.xlsx
EXCEL_OUT="devtest DullUSA Jun 2025.xlsx"
CSV_FILE=devtest.csv
MONTH=Jun
YEAR=2025
DEFINITION=devtest.ini

.PHONY: all run convert view clean

all: run convert view

run:
	./Cmdline.py --defn=$(DEFINITION) --excl=$(EXCEL_FILE) --month=$(MONTH) --year=$(YEAR)

convert:
	xlsx2csv $(EXCEL_OUT) > $(CSV_FILE)

view:
	column -s, -t $(CSV_FILE) | less -#2 -N -S

clean:
	rm -f $(CSV_FILE)

