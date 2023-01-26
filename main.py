import pandas as pd
import openpyxl.utils.cell

excel_dir = r'C:\Temp\ExcelPivotInput'


def read_excel(file):
    df = pd.read_excel(file, sheet_name='Data base')  # can also index sheet by name or fetch all sheets

    return df


def main():
    df = read_excel(f'{excel_dir}\\1.xlsx')
    cost_centers = df['Cost Center'].tolist();
    cost_centers = set(cost_centers)
    output_file = str(f'{excel_dir}\\out.xlsx')
    pivots = {}

    # Create pivot data and write to file
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer_pivot:
        for cc in cost_centers:
            pivots[cc] = df[df['Cost Center'] == cc].pivot_table(index=['Cost Element', 'Cost element name'],
                                                                 columns=['Period'], values=['Val/COArea Crcy'],
                                                                 aggfunc=['sum'])
            center = str(cc)
            center_int = int(cc)
            pivots[center_int].to_excel(writer_pivot, sheet_name=center)

    out_df = pd.read_excel(output_file, sheet_name=0)  # can also index sheet by name or fetch all sheets
    with pd.ExcelWriter(output_file, engine='openpyxl', mode='a') as writer:
        for sheet in cost_centers:
            for column in out_df:
                col_width = max(out_df[column].astype(str).map(len).max(), len(column))
                col_idx = out_df.columns.get_loc(column)
                col_letter = openpyxl.utils.cell.get_column_letter(col_idx + 1)
                writer.sheets[str(sheet)].column_dimensions[str(col_letter)].width = col_width


main()
# for cc in cost_centers:
#     writer.sheets[str(cc)].column_dimensions['A'].width = 100
