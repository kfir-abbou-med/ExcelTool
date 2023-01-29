import pandas as pd
import openpyxl.utils.cell
import openpyxl
import Constants
import ExcelUtils
excel_dir = r'C:\Temp\ExcelPivotInput'


def read_excel(file):
    df = pd.read_excel(file, sheet_name='Data base')  # can also index sheet by name or fetch all sheets
    return df


def set_auto_fit_width(file):
    with pd.ExcelWriter(file, engine='openpyxl', mode='a', if_sheet_exists="overlay") as writer:
        for sheet in writer.sheets:
            out_df = pd.read_excel(file, sheet_name=str(sheet))
            for column in out_df:
                col_width = max(out_df[column].astype(str).map(len).max(), len(column))
                col_idx = out_df.columns.get_loc(column)
                col_letter = openpyxl.utils.cell.get_column_letter(col_idx + 1)
                writer.sheets[str(sheet)].column_dimensions[str(col_letter)].width = col_width


def main():
    df = read_excel(f'{excel_dir}\\1.xlsx')
    cost_centers = df['Cost Center'].tolist();
    cost_centers = set(cost_centers)
    tmp_output_file = str(f'{excel_dir}\\tmp_out.xlsx')
    output_file = str(f'{excel_dir}\\out.xlsx')
    pivots = {}

    # Create pivot data and write to file
    with pd.ExcelWriter(tmp_output_file, engine='openpyxl') as writer_pivot:
        for cc in cost_centers:
            pivots[cc] = df[df['Cost Center'] == cc].pivot_table(index=['Cost Element', 'Cost element name'],
                                                                 columns=['Period'], values=['Val/COArea Crcy'],
                                                                 aggfunc=['sum'])
            center = str(cc)
            center_int = int(cc)
            pivots[center_int].to_excel(writer_pivot, sheet_name=center)



    # load excel file
    workbook = openpyxl.load_workbook(filename=tmp_output_file, data_only=False)

    # open workbook
    for sheet in workbook.sheetnames:
        # modify the desired cell

        s = workbook[sheet]
        data = get_last_row_column(s)
        s["A1"] = 'Cost Center'
        cost_center_name = Constants.cost_centers[int(sheet)]
        s["B2"] = cost_center_name + ' $'
        s["A3"] = 'Sum of Val/COArea Crcy'
        s['B1'] = sheet

        last_row = ExcelUtils.find_last_product_row(s)
        last_col = ExcelUtils.find_last_period_col(s)

        ExcelUtils.calc_total_for_product(s, last_row, last_col)
        ExcelUtils.calc_total_for_period(s, last_row+1, last_col)

        s[f'{Constants.num_hash(last_col+1)}4'] = Constants.grand_total_text
        s[f'{Constants.num_hash(last_col+3)}4'] = Constants.comments_text
        s[f'A{data[0]+1}'] = Constants.grand_total_text

        for row in range(1, 5):
            for col in range(1, last_col+2):
                s[f'{Constants.num_hash(col)}{row}'].fill = Constants.get_fill('title')

        s['B2'].fill = Constants.get_fill('cc')
        s = ExcelUtils.remove_borders(s)
        s = ExcelUtils.set_border_under_row(s, last_row, last_row, 1, last_col + 1)
        s = ExcelUtils.set_border_under_row(s, 4, 4, last_col + 2, last_col + 2 + 1)
        s = ExcelUtils.set_alignment(s, 1, last_row+1, 1, last_col+1, 'left', 'center')
        s = ExcelUtils.set_bold_text(sheet=s, min_row=1, max_row=last_row + 1, min_col=1, max_col=last_col + 1, is_bold=False)
        s = ExcelUtils.set_cell_format_number(sheet=s, min_row=5, max_row=last_row+1, min_col=3, max_col=last_col+1)
        s = ExcelUtils.set_months_title(sheet=s, last_col=last_col)
        s = ExcelUtils.calc_months_difference(sheet=s, min_row=5, max_row=last_row, min_col=3,max_col=last_col)

    # save the file
    workbook.save(filename=output_file)
    set_auto_fit_width(output_file)

    s = workbook['511200']
    d = s['C3'].value
    print(d)
    workbook.close()


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


main()
