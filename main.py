import pandas as pd
import openpyxl.utils.cell

excel_dir = r'C:\Temp\ExcelPivotInput'


def read_excel(file):
    df = pd.read_excel(file, sheet_name='Data base')  # can also index sheet by name or fetch all sheets

    return df


def set_titles(df, col_name):
    df.at[0, col_name] = 'kfkfkfkfkfk'
    return df


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

    out_df = pd.read_excel(tmp_output_file)  # can also index sheet by name or fetch all sheets

    with pd.ExcelWriter(tmp_output_file, engine='openpyxl', mode='a', if_sheet_exists="overlay") as writer:
        out_df = set_titles(out_df, 'Unnamed: 0')
        out_df.to_excel(writer, index=False)

    print(out_df)
    out_df.to_excel(tmp_output_file, index=False, header=False)
    #     for sheet in cost_centers:

            # for column in out_df:
            #     col_width = max(out_df[column].astype(str).map(len).max(), len(column))
            #     col_idx = out_df.columns.get_loc(column)
            #     col_letter = openpyxl.utils.cell.get_column_letter(col_idx + 1)
            #     writer.sheets[str(sheet)].column_dimensions[str(col_letter)].width = col_width
            # out_df.to_excel(writer, index=False)
main()
