# format-csv-to-excel
# CSV-to-Excel read/write logic is pulled from https://www.dev2qa.com/python-pandas-read-write-csv-file-and-convert-to-excel-file-example/
# XLSX cell formatting gleaned from https://xlsxwriter.readthedocs.io/tutorial02.html
# Cell format codes from https://xlsxwriter.readthedocs.io/format.html

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

import pandas
import os
import sys


# Read csv file use pandas module.
def read_csv_file_by_pandas(csv_file):
  df = None
  if os.path.exists(csv_file):
    df = pandas.read_csv(csv_file)
  else:
    print(csv_file + " does not exist.")
  return df

# Write pandas.DataFrame object to an excel file.
def write_to_excel_file_by_pandas(excel_file_path, frame):
  (nrows, ncols) = frame.shape
  excel_writer = pandas.ExcelWriter(excel_file_path, engine='xlsxwriter')
  frame.to_excel(excel_writer, 'From CSV')
    
  # My stuff...
  x = excel_writer      # x is the ExcelWriter object
  workbook = x.book     # workbook applies to all the sheets and characteristics within x
  sheet = x.sheets['From CSV']   # sheet is my worksheet
    
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
  for col in range(0, ncols-1):
    val = frame.columns[col]
    sheet.write(0, col, val, bold)

  # My logic to iterate over `excel_writer`
  for row in range(0, nrows):
    for col in range(0, ncols-1):
      val = frame.values[row, col]
      code = False
      final = val
            
      # Test the cell contents
      if val.startswith('%'):
        parts = val.split("%", 3)
        final = parts[2]
        code = codes[parts[1]]

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

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
  csvfile = sys.argv[1]
  if os.path.exists(csvfile):
    xlsx = os.path.splitext(csvfile)[0] + '.xlsx'
    data_frame = read_csv_file_by_pandas(csvfile)
    write_to_excel_file_by_pandas(xlsx, data_frame)
  else:
    sys.exit('Sorry, file ' + csvfile + ' was not found.')
    
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
