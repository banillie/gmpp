'''Programme to move internal master data into the same order as gmpp master data

Output - Excel master file with aggregate internally reported data in the same order as gmpp master data'''

from openpyxl import load_workbook, Workbook
from bcompiler.utils import project_data_from_master

def create_internal_master(dft_data, gmpp_dm):
    ws = gmpp_dm.active

    for i, name in enumerate(dft_data):
        print(name)
        ws.cell(row=1, column=4+i).value = name  # place project names in file

        dm_list = []
        for row_num in range(2, ws.max_row+1):
            key = ws.cell(row=row_num, column=1).value
            dm_list.append(key)

        # for loop for placing data into the worksheet
        for row_num in range(2, ws.max_row+1):
            key = ws.cell(row=row_num, column=1).value
            # this loop places all latest raw data into the worksheet
            try:
                if key in dft_data[name].keys():
                    ws.cell(row=row_num, column=4+i).value = dft_data[name][key]
                else:
                    pass
                    #ws.cell(row=row_num, column=4 + i).value = dft_data[name][key]
            except KeyError:
                pass

    return gmpp_dm


gmpp_datamap = load_workbook("C:\\Users\\Standalone\\Will\\masters folder\\dms\\gmpp_dm_merged_excel_master.xlsx")

#gmpp_master = project_data_from_master("C:\\Users\\Standalone\\Will\\masters folder\\core data\\gmpp_master_3"
#                                       "_2018.xlsx")

dft_master = project_data_from_master("C:\\Users\\Standalone\\Will\\masters folder\\core data\\master_3_2018.xlsx")

output = create_internal_master(dft_master, gmpp_datamap)

output.save("C:\\Users\\Standalone\\Will\\output_testing.xlsx")