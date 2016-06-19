import xlrd
import unicodecsv
import glob

def xls2csv (xls_filename, csv_filename):
    # Converts an Excel file to a CSV file.
    # If the excel file has multiple worksheets, only the first worksheet is converted.
    # Uses unicodecsv, so it will handle Unicode characters.
    # Uses a recent version of xlrd, so it should handle old .xls and new .xlsx equally well.

    wb = xlrd.open_workbook(xls_filename)
    sh = wb.sheet_by_index(0)

    fh = open(csv_filename,"wb")
    csv_out = unicodecsv.writer(fh, encoding='utf-8')

    for row_number in xrange (sh.nrows):
        csv_out.writerow(sh.row_values(row_number))

    fh.close()

for files in glob.glob("*.xls"):
    xlsname = files
    csvname = files.split(".xls")[0]+".csv"
    xls2csv (xlsname,csvname)
