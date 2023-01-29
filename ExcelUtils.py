import Constants


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
