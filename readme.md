# Automation Tools for D2L Marking Spreadsheets
This repository countains a suite of tools for automatically integrating spreadsheet-based marking reports into the D2L Brightspace gradebook.
##Test
## Functionality

### 1. Update master spreadsheet with data from individual spreadsheets
##### CLI
- Run `update_master.py [.../master_sheet.xlsx] [.../group_sheets/] [names cell] [grades cell]`
##### GUI
Run `gui.pyw`  
- Select your master spreadsheet and your group sheets directory
- Specify cells
- Click run

### 2. Scrape D2L
...

### Future functionality
- Pull data from D2L
- Create subdirectory for each groups files (marking report, original submission, marked submission)

## Requirements
Python 3 and the following packages:
- `pandas`
- `openpyxl`

Run `pip install pandas openpyxl` in terminal to install packages

## Authors
Developed by Dane Beliveau and Wayne Eberly at the University of Calgary.