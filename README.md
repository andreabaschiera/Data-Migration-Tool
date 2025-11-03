# VSME Template Migration Tool

The VSME Template Migration Tool automates the migration of data filled in older versions of the VSME Digital Template (Excel-based) into the latest template version.
It ensures that all Named Ranges (NRs) and data cells with missing ranges are transferred correctly and consistently between template versions — minimizing manual effort and the risk of errors during updates.

## Key Features
- Automatic version detection of the old template
- Checks for proper template fill before processing
- Transfers all user-entered data from the old template to the new one
- Handles renamed NRs and the non-referenced data cells
- Saves a new .xlsx ("result_from_1.*.*.xlsx") file ready for use with the latest VSME Digital Template

## Repository Structure
vsme-template-migration/
│
├── main.py                         # Main migration script
├── Template/
│   └── VSME-Digital-Template-1.1.1.xlsx    # Latest VSME template version
│
├── outils/                         # Utility modules for core operations
│   ├── check_template_fill.py
│   ├── months.py
│   ├── access_missingNR_table.py
│   ├── apply_changes_NR.py
│   ├── clean_NR_with_no_data.py
│   ├── paste_values.py
│   ├── access_NR_table.py
│   ├── copy_values.py
│
├── pickles/
│   └── missingNR_df.py             # Pickled with references of data cells with no NRs
│
└── README.md                       # Documentation 

## Requirements
1. Python environment and dependencies:
   - openpyxl: Read/write Excel 2010 xlsx/xlsm/xltx/xltm files
   - https://pandas.pydata.org/ : open source data analysis and manipulation tool
2. Filled-out Excel template with unmodified name ranges (see [digital templates](https://github.com/EFRAG-EU/Digital-Template-to-XBRL-Converter/tree/main/digital-templates))

## Notes & Warnings
The tool does not alter your original file — it creates a new copy.
If the selected workbook is not recognized as a valid VSME template, the script exits safely with a warning.
Some formatting (like checkboxes and arrows to navigate tabs) are not transferred.

## Authors
Developed by EFRAG's Digital Team
Maintained by EFRAG's Digital Team
For inquiries or support, contact: andrea.baschiera@efrag.org / thibault.magro@efrag.org
