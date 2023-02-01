import Constants
from openpyxl.styles import Border, Side, Alignment, Font
import re


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


def set_absolute_text(sheet, comments_col_num, total_row_num):
    sheet[f'{num_hash(comments_col_num)}4'] = Constants.comments_text
    sheet[f'A{total_row_num}'] = Constants.grand_total_text
    sheet['C1'] = ''
    sheet['C2'] = ''


def calc_total_for_period(sheet, last_row, last_col):
    first_row = 5

    for row in sheet.iter_rows(min_row=5, max_row=last_row, min_col=3, max_col=last_col):
        total_period = 0
        for cell in row:
            if cell.value is not None:
                total_period = float(cell.value) + total_period
        write_results(sheet, row[0].row, last_col + 1, total_period)


# def merge_cells(sheet, start_row, start_col, end_row, end_col):
#     sheet.merge_cells(start_row=start_row, start_column=start_col, end_row=end_row, end_column=end_col)
#     return sheet

def copy_data_to_new_sheet(sheet, new_sheet):
    mr = sheet.max_row
    mc = sheet.max_column

    col_offset = new_sheet.max_column

    for i in range(1, mr + 1):
        for j in range(1, mc + 1):
            # reading cell value from source excel file
            c = sheet.cell(row=i, column=j)
            if c.has_style:
                new_sheet.cell(row=i, column=j + col_offset + 1)._style = c._style

            # writing the read value to destination excel file
            new_sheet.cell(row=i, column=j + col_offset + 1).value = c.value


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


def is_float(string):
    # Compile a regular expression pattern to match valid float values
    pattern = r"^[-+]?[0-9]*\.?[0-9]+$"

    # Use re.match to check if the string matches the pattern
    # Returns a match object if there is a match, else None
    match = re.match(pattern, string)

    # Convert the match object to a boolean value
    # Returns True if there is a match, else False
    return bool(match)


def calc_months_difference(sheet, min_row, max_row, min_col, max_col):
    for r in range(min_row, max_row):
        # for c in range(max_col-1, max_col):
        current_month_val = sheet.cell(row=r, column=max_col).value
        previous_month_val = sheet.cell(row=r, column=max_col-1).value
        prev_month_text = str(sheet.cell(row=r, column=max_col-1).value)
        is_prev_float = is_float(prev_month_text)

        if prev_month_text is None or not is_prev_float:
            previous_month_val = 0
        if current_month_val is None:
            current_month_val = 0
        sheet.cell(row=r, column=max_col+1).value = current_month_val - previous_month_val
    return sheet


def set_bold_text(sheet, min_row=1, max_row=None, min_col=1, max_col=None, is_bold=False):
    for row in sheet.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            col = num_hash(cell.column)
            sheet[f'{col}{row[0].row}'].font = Font(bold=is_bold)
    return sheet


def set_cell_number_format(sheet, min_row, max_row, min_col, max_col):
    num_format = '#,##0.00;"-"#,##0.00'
    # num_format = '#,##0.00$;"-"#,##0$'

    for row in sheet.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            col = num_hash(cell.column)
            sheet[f'{col}{row[0].row}'].number_format = num_format
            if cell.value is None:
                cell.value = 0

            cell.value = float(cell.value)

    return sheet


def set_fill_on_area(sheet, min_row, max_row, min_col, max_col, color_key):
    for row in range(min_row, max_row):
        for col in range(min_col, max_col + 1):
            sheet[f'{num_hash(col)}{row}'].fill = Constants.get_fill(color_key)


def calc_and_set_total_for_product(sheet, last_row, last_col):
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


def num_hash(num):
    alpha = Constants.alpha

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
