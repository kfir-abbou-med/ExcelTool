import win32com.client as win32
import pandas as pd
import numpy as np
from pathlib import Path
import re
import sys
win32c = win32.constants


def pivot_table(wb: object, ws1: object, pt_ws: object, ws_name: str, pt_name: str, pt_rows: list, pt_filters: list,
                pt_fields: list):
    """
    wb = workbook1 reference
    ws1 = worksheet1 that contain the data
    pt_ws = pivot table worksheet number
    ws_name = pivot table worksheet name
    pt_name = name given to pivot table
    pt_rows, pt_cols, pt_filters, pt_fields: values selected for filling the pivot tables
    """

    # pivot table location
    pt_loc = len(pt_filters) + 2

    # grab the pivot table source data
    pc = wb.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=ws1.UsedRange)

    # create the pivot table object
    pc.CreatePivotTable(TableDestination=f'{ws_name}!R{pt_loc}C1', TableName=pt_name)

    # selecte the pivot table work sheet and location to create the pivot table
    pt_ws.Select()
    pt_ws.Cells(pt_loc, 1).Select()

    # Sets the rows, columns and filters of the pivot table
    for field_list, field_r in ((pt_filters, win32c.xlPageField),
                                (pt_rows, win32c.xlRowField)):
        for i, value in enumerate(field_list):
            pt_ws.PivotTables(pt_name).PivotFields(value).Orientation = field_r
            pt_ws.PivotTables(pt_name).PivotFields(value).Position = i + 1

    # Sets the Values of the pivot table
    for field in pt_fields:
        pt_ws.PivotTables(pt_name).AddDataField(pt_ws.PivotTables(pt_name).PivotFields(field[0]), field[1],
                                                field[2]).NumberFormat = field[3]

    # Visiblity True or Valse
    pt_ws.PivotTables(pt_name).ShowValuesRow = True
    pt_ws.PivotTables(pt_name).ColumnGrand = True


def run_excel(f_path: Path, f_name: str, sheet_name: str):
    filename = f_path / f_name

    # create excel object
    excel = win32.gencache.EnsureDispatch('Excel.Application')

    # excel can be visible or not
    excel.Visible = True  # False

    # try except for file / path
    try:
        wb = excel.Workbooks.Open(filename)
    except com_error as e:
        if e.excepinfo[5] == -2146827284:
            print(f'Failed to open spreadsheet.  Invalid filename or location: {filename}')
        else:
            raise e
        sys.exit(1)

    # set worksheet
    ws1 = wb.Sheets('Sales')

    # Setup and call pivot_table
    ws2_name = 'pivot_table'
    wb.Sheets.Add().Name = ws2_name
    ws2 = wb.Sheets(ws2_name)

    # update the pt_name, pt_rows, pt_cols, pt_filters, pt_fields at your preference
    pt_name = 'example'  # pivot table name, must be a string
    pt_rows = ['Genre']  # rows of pivot table, must be a list
    # pt_cols = []  # columns of pivot table, must be a list
    pt_filters = ['Year']  # filter to be applied on pivot table, must be a list
    # [0]: field name [1]: pivot table column name [3]: calulation method [4]: number format (explain the list item of pt_fields below)
    pt_fields = [['North America', 'Total Sales in North America', win32c.xlSum, '0'],  # must be a list of lists
                 ['Europe', 'Total Sales in Europe', win32c.xlSum, '0'],
                 ['Japan', 'Total Sales in Japan', win32c.xlSum, '0'],
                 ['Rest of World', 'Total Sales in Rest of World', win32c.xlSum, '0'],
                 ['Global', 'Total Global Sales', win32c.xlSum, '0']]
    # calculation method: xlAverage, xlSum, xlCount
    pivot_table(wb, ws1, ws2, ws2_name, pt_name, pt_rows, pt_filters, pt_fields)
    wb.Save()  # save the pivot table created


#    wb.Close(True)
#    excel.Quit()
