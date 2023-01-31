import glob
import os.path
import pathlib
import shutil

import pandas as pd
import openpyxl.utils.cell
import openpyxl
import Constants
import ExcelUtils


def read_excel(file, sheet_name='Data base'):
    df = pd.read_excel(file, sheet_name=sheet_name)  # can also index sheet by name or fetch all sheets
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


def set_hard_coded_text(sheet, cost_center):
    cost_center_name = Constants.cost_centers[int(cost_center)]
    sheet["A1"] = Constants.cost_center_text
    sheet["B2"] = cost_center_name + ' $'
    sheet["A3"] = Constants.sum_of_val_text
    sheet['B1'] = cost_center


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


def main():
    excel_dir = r'C:\Temp\ExcelPivotInput'
    files = glob.glob(f'{pathlib.Path().absolute()}\\*.xlsx')
    input_file = files[0]
    if not os.path.exists(excel_dir):
        os.makedirs(excel_dir)

    df = read_excel(input_file)
    cost_centers = df[Constants.cost_center_text].tolist()
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
        sh = workbook[sheet]
        set_hard_coded_text(sh, sheet)

        last_cell_occupied = ExcelUtils.get_last_row_column(sh)
        last_row = last_cell_occupied[0] # ExcelUtils.find_last_product_row(sh)
        last_col = last_cell_occupied[1] # ExcelUtils.find_last_period_col(sh)

        # Set text
        ExcelUtils.calc_and_set_total_for_product(sh, last_row, last_col)
        sh[f'{Constants.num_hash(last_col + 2)}4'] = Constants.comments_text
        sh[f'A{last_cell_occupied[0] + 1}'] = Constants.grand_total_text
        sh['C1'] = ''
        sh['C2'] = ''


        # Set some style issues
        ExcelUtils.set_fill_on_area(sh, min_row=1, max_row=5,min_col=1, max_col=last_col, color_key='title')
        sh['B2'].fill = Constants.get_fill('cc')

        ExcelUtils.remove_borders(sh)
        ExcelUtils.set_border_under_row(sh, last_row, last_row, 1, last_col)
        ExcelUtils.set_border_under_row(sh, 4, 4, last_col + 1, last_col + 2)
        ExcelUtils.set_alignment(sh, 1, last_row + 1, 1, last_col + 1, 'left', 'center')
        ExcelUtils.set_bold_text(sheet=sh, min_row=1, max_row=last_row + 1, min_col=1, max_col=last_col + 1, is_bold=False)
        ExcelUtils.set_bold_text(sheet=sh, min_row=last_row+1, max_row=last_row + 1, min_col=1, max_col=last_col, is_bold=True)
        ExcelUtils.set_cell_format_number(sheet=sh, min_row=5, max_row=last_row + 1, min_col=3, max_col=last_col + 1)
        ExcelUtils.set_months_title(sheet=sh, last_col=last_col)
        ExcelUtils.calc_months_difference(sheet=sh, min_row=5, max_row=last_row+1, min_col=3, max_col=last_col)

    # save the file
    workbook.save(filename=output_file)
    set_auto_fit_width(output_file)
    new_sheet = workbook.create_sheet('results')

    for sheet in workbook.sheetnames:
        if sheet == 'results':
            break
        # change xxx with the sheet name that includes the data
        file = Constants.output_file_name
        ws2 = new_sheet

        # calculate total number of rows and
        # columns in source excel file
        s = workbook[sheet]

        copy_data_to_new_sheet(s, ws2)

        # delete sheet
        del workbook[sheet]

    # saving the destination excel file
    res_sheet = workbook['results']
    res_sheet.delete_cols(1,2)
    workbook.save(str(file))
    set_auto_fit_width(file)
    shutil.rmtree(excel_dir)

main()
