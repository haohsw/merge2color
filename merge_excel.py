import os
from types import new_class
import pandas as pd
import pprint
from color import color_coordinated, fill_cell
import xlrd
ROOT_PATH = os.path.dirname(os.path.abspath(__file__))
# MERGE_IN = f'{ROOT_PATH}/convert/csv/'
MERGE_IN = f'{ROOT_PATH}/merge/merge_in/'
MERGE_OUT = f'{ROOT_PATH}/merge/merge_out/'
try:
    files = os.listdir(MERGE_IN)
    files = sorted(files)
    pprint.pprint(files)
    # Manual input files
#     files = ['output_JN8MM_202207131357.xls',
#  'output_P758W_202207131357.xls',
#  'output_XRMVP_202207131357.xls']
    writer = pd.ExcelWriter(MERGE_OUT + 'pandas_multiple.xls', engine='xlsxwriter')
    w_count = 1
    for file in files:
        file_path = MERGE_IN + file
        sheet_count = len(list(pd.read_excel(file_path, sheet_name = None)))
        print(sheet_count)
        writer_b = pd.ExcelWriter(file_path, engine="openpyxl")
        for count in range(int(sheet_count)):
            df = pd.read_excel(file_path, sheet_name = str(count + 1), header=None)
            color_coordinates = color_coordinated(writer_b, str(w_count))
            print(color_coordinates)
            if len(color_coordinates) > 0:
                for color_coordinate in color_coordinates:
                    fill_cell(writer, str(w_count), )
            df.to_excel(writer, sheet_name= str(w_count), index=False, header=False)
            w_count = w_count +1
    writer.save()        
except Exception as ex:
    print(ex)