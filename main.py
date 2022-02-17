# format-csv-to-excel
# CSV-to-Excel read/write logic is pulled from https://www.dev2qa.com/python-pandas-read-write-csv-file-and-convert-to-excel-file-example/
# XLSX cell formatting gleaned from https://xlsxwriter.readthedocs.io/tutorial02.html
# Cell format codes from https://xlsxwriter.readthedocs.io/format.html

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

import pandas
import os
import sys
import argparse


# Read csv file use pandas module.
def read_csv_file_by_pandas(csv_file):
  df = None
  if os.path.exists(csv_file):
    df = pandas.read_csv(csv_file)
  else:
    print(csv_file + " does not exist.")
  return df

# Write pandas.DataFrame object to an excel file.
def write_to_excel_file_by_pandas(excel_file_path, frame, dg):
  (nrows, ncols) = frame.shape
  excel_writer = pandas.ExcelWriter(excel_file_path, engine='xlsxwriter')
  # frame.to_excel(excel_writer, 'From CSV')
    
  # My stuff...
  x = excel_writer      # x is the ExcelWriter object
  workbook = x.book     # workbook applies to all the sheets and characteristics within x
  sheet = workbook.add_worksheet('From a Formatted CSV')   # sheet is my worksheet
    
  # Create my cell formats...
  bold = workbook.add_format({'bold': True})
  orange = workbook.add_format({'bold': True, 'bg_color': '#FFB347'})
  yellow = workbook.add_format({'bold': True, 'bg_color': '#FFFFBF'})
  red = workbook.add_format({'bold': True, 'font_color': 'red', 'bg_color': '#FFFFBF'})
  soft = workbook.add_format({'bold': False, 'font_color': 'gray'})
  green = workbook.add_format({'bold': False, 'font_color': 'green'})
    
  codes = { 'NORMAL':False, 'BOLD':bold, 'ORANGE':orange, 'YELLOW':yellow, 'RED':red, 'SOFT':soft, 'GREEN':green }

  width = [0] * ncols

  # Write out the column headings
  for col in range(0, ncols):
    val = frame.columns[col]
    sheet.write(0, col, val, bold)

  # My logic to iterate over `excel_writer`
  for row in range(0, nrows):
    for col in range(0, ncols):
      val = frame.values[row, col]
      code = False
      final = val
            
      # Test the cell contents
      if val.startswith('%'):
        parts = val.split("%", 3)
        final = parts[2]
        code = codes[parts[1]]

      # Apply special processing if enabled
      if dg:
        final = dg_processing(col, final)

      # Write the cell contents
      if code:
        sheet.write(row+1, col, final, code)
      else:
        sheet.write(row+1, col, final)
   
      # Set column width to len(final)+1 with a maximum width of 20
      w = min(20, len(final)+1)
      if w > width[col]:
        sheet.set_column(col, col, w)
        width[col] = w
   
  x.save()
  print(excel_file_path + ' has been created.')
  
# Define special processing for Digital.Grinnell
def dg_processing(c, v):
  # Turn the full PID into a hyperlink
  if c == 1:
    link = "https://digital.grinnell.edu/islandora/object/" + v
    formula = '=HYPERLINK("' + link + '", "' + v + '")'
    return formula
  # Nothing to process... return unchanged
  return v;


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
  parser = argparse.ArgumentParser()
  parser.add_argument("csvfile", type=str, help="Specify the path of the .csv file for processing")
  parser.add_argument("--verbose", action="store_true", help="Increase output verbosity")
  parser.add_argument("--dg", action="store_true", help="Apply 'special' Digital.Grinnell IHC rules")
  args = parser.parse_args()
  if args.verbose:
    print("Verbose output selected.")
  if args.dg:
    print("Special Digital.Grinnell processing is selected.")
  csv = args.csvfile
 
  if os.path.exists(csv):
    xlsx = os.path.splitext(csv)[0] + '.xlsx'
    data_frame = read_csv_file_by_pandas(csv)
    write_to_excel_file_by_pandas(xlsx, data_frame, args.dg)
  else:
    sys.exit('Sorry, file ' + csv + ' was not found.')
    
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
