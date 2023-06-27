# income_report
Income Reconciliation Report Generator


## Instruction
-----
1. Copy `hook-tkinterdnd2.py` file and `tkinterdnd2` folder in https://github.com/Eliav2/tkinterdnd2 to be able to use TkInterDnD2 with pyinstaller. More details in https://github.com/Eliav2/tkinterdnd2
2. Create `file_version_info.txt` using: `create-version-file metadata.yml --outfile file_version_info.txt`
3. Create EXE file using: `pyinstaller --onefile --add-data "icon.ico;." -i icon.ico --noconsole --version-file=file_version_info.txt --additional-hooks-dir=. main.py`
4. Run EXE file to start the program
5. Drag and drop report excel files into the text box or select from the menu.
6. Select `Generate Report` button to create the report.


## Limitation
-----
1. Income reports in Excel files must be downloaded beforehand manually.
2. The first two digits in the student's code is their year's code.


## Libraries
----
1. Pyinstaller
2. tkinter
3. tkinterdnd2
4. openpyxl
5. pandas