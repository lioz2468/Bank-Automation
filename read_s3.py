import openpyxl, sys

path = sys.argv[1]
wb = openpyxl.load_workbook(path, data_only=False)
ws = wb.worksheets[2]
print(f"Sheet: {ws.title}")
for row in ws.iter_rows():
    for cell in row:
        if cell.value is not None:
            print(f"{cell.coordinate}\t{repr(cell.value)}")
