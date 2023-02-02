import sys
import openpyxl as xl
import csv

def is_excel_file(filename):
    return filename.split(".")[-1] == "xlsx"

def check_args():
    success = False
    if len(sys.argv)<2:
        print("Too few command-line arguments")
    elif len(sys.argv)>2:
        print("Too many command-line arguments")
    elif not is_excel_file(sys.argv[1]):
        print("Not a xlsx file")
    else: success = True
    return success

def csv_from_excel(xlsxfile):
    wb = xl.load_workbook(xlsxfile)
    sh = wb.active
    your_csv_file = open(f"{xlsxfile.split('.')[0]}.csv", 'w')
    wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)
    for row in sh.iter_rows(min_row=3,values_only=True):
        wr.writerow(row)
    your_csv_file.close()

if __name__=="__main__":
    if check_args():
        csv_from_excel(sys.argv[1])
    else:
        sys.exit(1)

