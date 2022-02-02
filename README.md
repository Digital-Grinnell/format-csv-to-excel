# format-csv-to-excel

This Python3 utility reads a named .csv file and outputs a correspondingly named .xlsx file with formatting.

## Use

The script, `main.py`, uses a Python VENV (Virtual ENVironment) and requires one command-line argument, the path to the .csv file that is to be processed.  An example command-line on my Mac might be:

```markdown
/Users/mcfatem/GitHub/format-csv-to-excel/venv/bin/python /Users/mcfatem/GitHub/format-csv-to-excel/main.py /Users/mcfatem/Downloads/ihcQ_grinnell_FULL-Feb-1-2022_ih.csv
```

## Output

The script will attempt to produce a new .xlsx file with the same path as the input .csv file.  For example, the output from the above example command-line was:

```markdown
/Users/mcfatem/Downloads/ihcQ_grinnell_FULL-Feb-1-2022_ih.xlsx has been created.
```

The named output file was indeed created as indicated.

## CSV with Formatting

Formats in the .csv are expressed as %CODE% prefixes in each cell.  A simple example excerpt from a prepared .csv file might be:

| | | | | | | |  
|--- |--- |--- |--- |--- |--- |--- |  
| %ORANGE%31911 | %BOLD%grinnell:31911 | Sin for a Sign? | [1] sp_pdf | [2] grinnell:student-scholarship | application/pdf | [0] None | %SOFT%[0] None |  
| 31910 | %BOLD%grinnell:31910 | Divinity and Power | [1] sp_pdf | [2] grinnell:student-scholarship | application/pdf | [0] None" | %SOFT%[0] None |

## Valid CODEs

The following `%CODE%` values are valid and transform into `ExcelWriter.workbook.add_format` definitions as shown below.

| Code | Cell Format |
| ---  | --- |
| %BOLD% | { 'bold': True } |
| %ORANGE% | { 'bold': True, 'bg_color': '#FFB347' } | 
| %YELLOW% | { 'bold': True, 'bg_color': '#FFFFBF' } |
| %RED% | { 'bold': True, 'font_color': 'red', 'bg_color': '#FFFFBF' } |
| %SOFT% | { 'bold': False, 'font_color': 'gray' } |
| %GREEN% | { 'bold': False, 'font_color': 'green' } |
| %NORMAL% | No cell formatting is applied. |

## Sources

- CSV-to-Excel read/write logic is pulled from https://www.dev2qa.com/python-pandas-read-write-csv-file-and-convert-to-excel-file-example/  
- XLSX cell formatting gleaned from https://xlsxwriter.readthedocs.io/tutorial02.html  
- Cell format codes from https://xlsxwriter.readthedocs.io/format.html  
