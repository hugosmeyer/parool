PYTHON_VERSION = 3.11.4
PYTHON_EXE = python-$(PYTHON_VERSION).exe
PYTHON_URL = https://www.python.org/ftp/python/$(PYTHON_VERSION)/$(PYTHON_EXE)
PYTHON_PATH = C:\\Python311\\python.exe
PYINSTALLER = C:\\Python311\\Scripts\\pyinstaller.exe
SCRIPT = Payroll.py
ASSETS = background.png

all: build

python:
	wget -nc $(PYTHON_URL)
	WINEARCH=win32 wine $(PYTHON_EXE)

deps:
	wine $(PYTHON_PATH) -m ensurepip
	wine $(PYTHON_PATH) -m pip install --upgrade pip
	wine $(PYTHON_PATH) -m pip install pillow pyinstaller openpyxl

clean:
	rm -rf build dist __pycache__ *.spec

build: clean
	wine $(PYINSTALLER) --onefile --add-data "$(ASSETS);." $(SCRIPT)

run:
	wine dist/$(basename $(SCRIPT)).exe


