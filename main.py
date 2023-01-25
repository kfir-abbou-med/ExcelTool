import pandas as pd
import numpy as np


excel_file = r'C:\Temp\ExcelPivotInput'


def read_excel(file):
    df = pd.read_excel(file, sheet_name='Data base')  # can also index sheet by name or fetch all sheets

    return df


def main():
    df = read_excel(f'{excel_file}\\1.xlsx')

    # pivot = df.pivot_table(index=['Cost Element', 'Cost element name'], columns=['Period'], values=['Val/COArea Crcy'],
    #                        aggfunc=['sum'])

    cost_centers = df['Cost Center'].tolist();
    cost_centers = set(cost_centers)
    output_file = f'{excel_file}\\out.xlsx'
    pivots = {}
    writer = pd.ExcelWriter(output_file, engine='openpyxl')
    for cc in cost_centers:
        pivots[cc] = df[df['Cost Center'] == cc].pivot_table(index=['Cost Element', 'Cost element name'],
                                                             columns=['Period'], values=['Val/COArea Crcy'],
                                                             aggfunc=['sum'])
        center = str(cc)
        center_int = int(cc)
        pivots[center_int].to_excel(writer, sheet_name=center)
    writer.save()

    for cc in cost_centers:
        df = pd.read_excel(output_file, sheet_name=str(cc))
        writer.sheets[str(cc)].column_dimensions['A'].width = 200

        # for column in df:
        #     column_width = max(df[column].astype(str).map(len).max(), len(column))
            # col_idx = df.columns.get_loc(column)
            # writer.sheets[str(cc)].set_column(col_idx, col_idx, column_width)

    writer.save()


main()
