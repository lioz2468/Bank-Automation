import openpyxl

for label, path in [('CORRECT', r'C:\Users\lioz_\Downloads\463579.xlsx'),
                    ('WRONG',   r'C:\Bank_Automation\output.xlsx')]:
    wb = openpyxl.load_workbook(path, data_only=False)
    ws = wb.worksheets[2]
    with open(f'C:\\Bank_Automation\\{label}_s3.txt', 'w', encoding='utf-8') as f:
        f.write(f'Sheet: {ws.title}\n')
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    f.write(f'  {cell.coordinate}: {repr(cell.value)}\n')
    print(f'{label} done')
