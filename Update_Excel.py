from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import NamedStyle
from openpyxl.cell.cell import TIME_TYPES  # For date detection

# Update the Excel Dashboard
def openwb(loc):
    return load_workbook(f"{loc}/Excel_Dashboard.xlsx")

def writewb(wb, df, sheet_name, start_row=1, start_col=1, clear_sheet=False, 
           autofit=True, hide_gridlines=True, date_col="Date"):
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.create_sheet(sheet_name)

    if clear_sheet:
        for row in ws.iter_rows():
            for cell in row:
                cell.value = None

    # Headers
    for i, col in enumerate(df.columns, start=start_col):
        ws.cell(row=start_row, column=i, value=col)

    # Data (Polars dates auto-convert)
    data = df.to_dicts()
    for r_idx, row in enumerate(data, start=start_row + 1):
        for c_idx, col in enumerate(df.columns, start=start_col):
            cell = ws.cell(row=r_idx, column=c_idx)
            value = row.get(col)
            cell.value = value
            # **FORMAT DATE COLUMN**
            if col.lower() == date_col.lower() and isinstance(value, TIME_TYPES):
                cell.number_format = 'dd/mm/yyyy'

    if hide_gridlines:
        ws.sheet_view.showGridLines = False

    if autofit:
        autofit_worksheet(ws)

def savewb(wb, loc):
    wb.save(f"{loc}/Excel_Dashboard.xlsx")
    wb.close()

def autofit_worksheet(ws, min_width=10, max_width=50, padding=2):
    for col_idx in range(1, ws.max_column + 1):
        max_len = 0
        col_letter = get_column_letter(col_idx)

        for row_idx in range(1, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is not None:
                max_len = max(max_len, len(str(value)))

        ws.column_dimensions[col_letter].width = min(max(max_len + padding, min_width), max_width)