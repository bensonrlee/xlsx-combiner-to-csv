# XLSX Combiner and Converter to CSV
Combines Excel files (of the same structure) and convert them to one CSV.  If the Excel files do not have the same structures, the conversion will not be attempted.  The script only processes the first worksheet of each Excel file.

## Dependencies
The script attempts to install depenencies.  If the user does not have the necessary permissions, the dependencies below need to be installed manually.
1. Python
2. `pandas`
3. `openpyxl`

## Usage
The inputs can be provided as command line arguments.  If any argument is missing, the script prompts for it.

`python xlsxcombiner.py`
