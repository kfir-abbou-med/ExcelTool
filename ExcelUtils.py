import Constants
import openpyxl
from openpyxl.styles import Border, Side, Alignment, Font


def get_last_row_column(ws):
    row = ws.max_row
    col = ws.max_column
    return row, col


def get_cell_row_col_with_value(sheet, value):
    for row in sheet.iter_rows(min_row=1, min_col=1):
        for cell in row:
            if cell.value == value:
                print(cell)
                return row, cell
    return 0, 0


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


def set_border_under_row(sheet, min_row, max_row, min_col, max_col):
    any_side = Side(border_style=None)
    bottom = Side(border_style='thin')
    border = Border(top=any_side, left=any_side, right=any_side, bottom=bottom)
    for row in sheet.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
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


def set_months_title(sheet, last_col):
    last_month_int = sheet.cell(row=3, column=last_col).value
    pre_last_month_int = sheet.cell(row=3, column=last_col - 1).value
    last_month_name = Constants.months[last_month_int]
    result_cell_letter = num_hash(last_col+1)

    if str(pre_last_month_int).isnumeric():
        pre_last_month_name = Constants.months[pre_last_month_int]
        sheet[f'{result_cell_letter}4'] = f'{last_month_name} vs {pre_last_month_name}'
    else:
        sheet[f'{result_cell_letter}4'] = f'{last_month_name}'
    return sheet


def calc_months_difference(sheet, min_row, max_row, min_col, max_col):
    if max_col - min_col > 1:
        for r in range (min_row, max_row):
            for c in range(max_col-1, max_col):
                current_month_val = sheet.cell(row=r, column=max_col).value
                previous_month_val = sheet.cell(row=r, column=max_col-1).value
                if current_month_val is not None and previous_month_val is not None:
                    sheet.cell(row=r, column=max_col+1).value = current_month_val - previous_month_val

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


def set_fill_on_area(sheet, min_row, max_row, min_col, max_col, color_key):
    for row in range(min_row, max_row):
        for col in range(min_col, max_col + 1):
            sheet[f'{Constants.num_hash(col)}{row}'].fill = Constants.get_fill(color_key)


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
