# Automation Tools for D2L Marking Spreadsheets

This repository countains a suite of tools for automatically integrating spreadsheet-based marking reports into the D2L Brightspace gradebook.

### Current functionality

#### 1. Update master spreadsheet with data from individual spreadsheets

`update_master.py` in a directory which contains a provided master spreadsheet called `master_sheet.xlsx` and a sub-directory `group_sheets` which contains the individual group marking reports.
    - Test data provided

### Future functionality

- Pull data from D2L
- Create subdirectory for each groups files (marking report, original submission, marked submission)
- GUI

## Requirements

Python 3 and the following packages:

- `pandas`
- `openpyxl`

Run `pip install pandas openpyxl` in terminal to install packages

### Authors

- Dane Beliveau
- Wayne Eberly