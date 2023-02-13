import glob
import os.path
import pathlib
import shutil
import sys

import pandas as pd
import openpyxl.utils.cell
import openpyxl

import AutoFitTool
import Constants
import ExcelUtils


def read_excel(file, sheet_name='Data base'):
    df = pd.read_excel(file, sheet_name=sheet_name)  # can also index sheet by name or fetch all sheets
    return df


# def autofit_cols(file, sheet, name):
#     writer = pd.ExcelWriter(file, engine="xlsxwriter")
#
#         # Loop through columns of current worksheet,
#         # and set correct width for each one
#     for column in sheet:
#         column_width = max(sheet[column].astype(str).map(len).max(), len(column))
#         col_idx = sheet.columns.get_loc(column)
#         writer.sheets[0].set_column(col_idx, col_idx, column_width)
#
#     writer.save()


# def set_auto_fit_width1(sheet):
#     for column in sheet.columns:
#         # col_width = (max(sheet[column].astype(str).map(len).max(), len(column)) + 2) * 3.3
#         col_width = (column[5])
#         # col_idx = sheet.columns.get_loc(column)
#         # col_letter = openpyxl.utils.cell.get_column_letter(col_idx + 1)
#         col_letter = column[2].column_letter
#         # sheet.column_dimensions[str(col_letter)].width = col_width
#         sheet.column_dimensions[str(col_letter)].bestFit = True


# def get_max_text_width(sheet, col_index):
#     width = 0
#     for row in sheet.max_rows:
#         cell = sheet.cell(row, col_index)


def set_auto_fit_width(file, sheet_name, ws):
    with pd.ExcelWriter(file, engine='openpyxl', mode='a', if_sheet_exists="overlay") as writer:
        # for sheet in writer.sheets:
        out_df = pd.read_excel(file, sheet_name=str(sheet_name))
        for column in out_df:
            col_width = (max(out_df[column].astype(str).map(len).max(), len(column)) + 2)
            col_idx = out_df.columns.get_loc(column)
            col_letter = openpyxl.utils.cell.get_column_letter(col_idx + 1)
            writer.sheets[str(sheet_name)].column_dimensions[str(col_letter)].width = col_width
            ws.column_dimensions[str(col_letter)].width = col_width


def set_hard_coded_text(sheet, cost_center):
    cost_center_name = Constants.cost_centers[int(cost_center)]
    sheet["A1"] = Constants.cost_center_text
    sheet["B2"] = cost_center_name + ' $'
    sheet["A3"] = Constants.sum_of_val_text
    sheet['B1'] = cost_center


def get_all_total_per_month(sheet):
    totals = {}
    # Number of columns shifted
    factor = 2
    for colidx in sheet.iter_cols(min_row=sheet.max_row, min_col=3, max_row=sheet.max_row, max_col=sheet.max_column):
        key = colidx[0].col_idx
        val = sheet[colidx[0].coordinate].value
        month = sheet[f'{colidx[0].column_letter}3'].value
        months_range = range(1, 12, 1)
        if month is not None and month in months_range:
            if val is None:
                val = 0
            totals[key-factor] = val
    return totals


all_sheets_total_per_month = {}


def init_all_sheets_total_per_month():
    for i in range(1, 13):
        all_sheets_total_per_month[i] = 0


def add_to_all_sheets_total(single_sheet_total):
    for key in single_sheet_total.keys():
        if all_sheets_total_per_month.__contains__(key):
            all_sheets_total_per_month[key] = all_sheets_total_per_month[key] + single_sheet_total[key]
        else:
            all_sheets_total_per_month[key] = single_sheet_total[key]


def main():
    excel_dir = r'C:\Temp\ExcelPivotInput'
    # files = glob.glob(f'{pathlib.Path().absolute()}\\*.xlsx')
    init_all_sheets_total_per_month()

    input_file = sys.argv[1] #files[0]
    print(f'Loaded input: {input_file}')
    if not os.path.exists(excel_dir):
        os.makedirs(excel_dir)

    df = read_excel(input_file)
    cost_centers = df[Constants.cost_center_text].tolist()
    cost_centers = set(cost_centers)
    tmp_output_file = str(f'{excel_dir}\\tmp_out.xlsx')
    pivots = {}

    # Create pivot data and write to file
    with pd.ExcelWriter(tmp_output_file, engine='openpyxl') as writer_pivot:
        for cc in cost_centers:
            pivots[cc] = df[df['Cost Center'] == cc].pivot_table(index=['Cost Element', 'Cost element name'],
                                                                 columns=['Period'],
                                                                 values=['Val/COArea Crcy'],
                                                                 aggfunc=['sum'])
            center = str(cc)
            center_int = int(cc)
            pivots[center_int].to_excel(writer_pivot, sheet_name=center)

    # load excel file
    workbook = openpyxl.load_workbook(filename=tmp_output_file, data_only=False)

    workbook.create_sheet(Constants.totals_text)
    totals_sheet = workbook[Constants.totals_text]
    ExcelUtils.set_const_text_sum_sheet(totals_sheet)

    # open workbook
    for sheet in workbook.sheetnames:
        if sheet == Constants.totals_text:
            continue

        curr_sheet = workbook[sheet]
        set_hard_coded_text(curr_sheet, sheet)

        last_cell_occupied = ExcelUtils.get_last_row_column(curr_sheet)
        last_row = last_cell_occupied[0]
        last_col = last_cell_occupied[1]
        ExcelUtils.set_totals_for_budget(totals_sheet, workbook[sheet], last_row, last_col)

        # Set text
        ExcelUtils.calc_and_set_total_for_product(curr_sheet, 5, last_row, 3, last_col)
        ExcelUtils.set_absolute_text(curr_sheet, last_col + 2, last_cell_occupied[0] + 1)

        totals_for_sheet = get_all_total_per_month(workbook[sheet])
        add_to_all_sheets_total(totals_for_sheet)

        # Set some style issues
        ExcelUtils.set_fill_on_area(curr_sheet, min_row=1, max_row=5, min_col=1, max_col=last_col, color_key='title')
        curr_sheet['B2'].fill = Constants.get_fill('cc')
        ExcelUtils.remove_borders(curr_sheet)
        ExcelUtils.set_border_under_row(curr_sheet, last_row, last_row, 1, last_col)
        ExcelUtils.set_border_under_row(curr_sheet, 4, 4, last_col + 1, last_col + 2)
        ExcelUtils.set_alignment(curr_sheet, 1, last_row + 1, 1, last_col + 1, 'left', 'center')
        ExcelUtils.set_bold_text(sheet=curr_sheet, min_row=1, max_row=last_row + 1, min_col=1, max_col=last_col + 1,
                                 is_bold=False)
        ExcelUtils.set_bold_text(sheet=curr_sheet, min_row=last_row + 1, max_row=last_row + 1, min_col=1,
                                 max_col=last_col, is_bold=True)
        ExcelUtils.set_months_title(sheet=curr_sheet, last_col=last_col)
        ExcelUtils.calc_months_difference(sheet=curr_sheet, min_row=5, max_row=last_row + 2, min_col=3,
                                          max_col=last_col)
        ExcelUtils.set_all_sheet_numbers_to_number_format(curr_sheet, min_row=4, min_col=3)

    ExcelUtils.set_all_totals(totals_sheet, all_sheets_total_per_month)
    ExcelUtils.set_alignment(totals_sheet, 1, totals_sheet.max_row, 1, totals_sheet.max_column,  'center', 'center')
    ExcelUtils.set_fill_on_area(totals_sheet, 1, 4, 1, totals_sheet.max_column, 'title')
    ExcelUtils.set_bold_text(totals_sheet, 8, 8, 1, totals_sheet.max_column, True)
    ExcelUtils.set_all_sheet_numbers_to_number_format(totals_sheet, min_row=5)

    # Copy all results to a single sheet
    results_sheet = workbook.create_sheet(Constants.results_text)
    all_sheets_but_results = (s_ for s_ in workbook.sheetnames if s_ != Constants.results_text and
                              s_ != Constants.totals_text)

    for sheet in all_sheets_but_results:
        active_sheet = workbook[sheet]
        ExcelUtils.copy_data_to_new_sheet(sheet=active_sheet, new_sheet=results_sheet)
        # delete sheet
        del workbook[sheet]

    results_sheet.delete_cols(1, 2)

    # saving the destination Excel file
    AutoFitTool.auto_fit_cols(results_sheet)
    AutoFitTool.auto_fit_cols(totals_sheet)
    workbook.save(str(Constants.output_file_name))
    shutil.rmtree(excel_dir)


main()
# pyinstaller --noconfirm --onefile --console --icon "C:/Temp/ExcelPivotInput - Copy/App/images.ico"
# --hidden-import "pandas"  "C:/Users/abbouk2/PycharmProjects/ExcelTool/main.py"