import Constants
import openpyxl
from openpyxl.styles import Border, Side, Alignment, Font

def validate_excel_file_data(workbook):
    return True


def find_last_period_col(sheet):
    for row in sheet.iter_rows(min_row=3, min_col=3, max_row=3):
        for cell in row:
            c = cell
    return c.column


def find_last_product_row(sheet):
    for row in sheet.iter_rows(min_row=5, min_col=1, max_col=1):
        for cell in row:
            r = cell.row
            if cell.value is None or cell.value == Constants.grand_total_text:
                return cell.row
    return r


def calc_total_for_period(sheet ,last_row, last_col):
    first_row = 5

    for row in sheet.iter_rows(min_row=5, max_row=last_row, min_col=3, max_col=last_col):
        total_period = 0
        for cell in row:
            if cell.value is not None:
                total_period = float(cell.value) + total_period
        write_results(sheet, row[0].row, last_col + 1, total_period)


def merge_cells(sheet, start_row, start_col, end_row, end_col):
    sheet.merge_cells(start_row=start_row, start_column=start_col, end_row=end_row, end_column=end_col)
    return sheet


def remove_borders(sheet):
    any_side = Side(border_style=None)
    border = Border(top=any_side, left=any_side, right=any_side, bottom=any_side)
    for row in sheet.iter_rows(min_row=1, min_col=1):
        for cell in row:
            col = num_hash(cell.column)
            sheet[f'{col}{row[0].row}'].border = border
    return sheet


def set_border_above_total_row(sheet, row, col):
    any_side = Side(border_style=None)
    top = Side(border_style='thin')
    border = Border(top=top, left=any_side, right=any_side, bottom=any_side)
    for row in sheet.iter_rows(min_row=row, max_row=row, min_col=1, max_col=col):
        for cell in row:
            col = num_hash(cell.column)
            sheet[f'{col}{row[0].row}'].border = border
    return sheet


def set_alignment(sheet, min_row, max_row, min_col, max_col, horizontal, vertical):
    for row in sheet.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            col = num_hash(cell.column)
            sheet[f'{col}{row[0].row}'].alignment = Alignment(horizontal=horizontal, vertical=vertical)
    return sheet


def set_bold_text(sheet, min_row=1, max_row=None, min_col=1, max_col=None, is_bold=False):
    for row in sheet.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            col = num_hash(cell.column)
            sheet[f'{col}{row[0].row}'].font = Font(bold=is_bold)
    return sheet


def set_cell_format_number(sheet, min_row, max_row, min_col, max_col):
    for row in sheet.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            col = num_hash(cell.column)
            sheet[f'{col}{row[0].row}'].number_format = '"$"#,##0.00;("$"#,##0)'
            if cell.value is not None:
                cell.value = float(cell.value)
    return sheet


def calc_total_for_product(sheet, last_row, last_col):
    first_row = 5
    total_product = {}
    for row in sheet.iter_rows(min_row=5, max_row=last_row, min_col=3, max_col=last_col):
        for cell in row:
            if cell.value is not None:
                if total_product.__contains__(cell.column):
                    total_product[cell.column] = total_product[cell.column] + cell.value
                else:
                    total_product[cell.column] = cell.value

    for key in total_product.keys():
        col = num_hash(key)
        row = last_row + 1
        sheet[f'{col}{row}'] = total_product[key]


def write_results(sheet, row, col, total):
    col_letter = num_hash(col)
    sheet[f'{col_letter}{row}'] = total


alpha = Constants.alpha


def num_hash(num):
    if num < 26:
        return alpha[num-1]
    else:
        q, r = num//26, num % 26
        if r == 0:
            if q == 1:
                return alpha[r-1]
            else:
                return num_hash(q-1) + alpha[r-1]
        else:
            return num_hash(q) + alpha[r-1]
