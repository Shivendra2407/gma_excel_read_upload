import pandas as pd
import sys, os
import xlrd
import pdb
# from dbcon import get_connection


def read_parent_sheet(complete_file_path):
    try:
        xls = pd.ExcelFile(complete_file_path)
        # parent_sheet = pd.read_excel(xls,'Overview')
        overview_df = xls.parse("Overview",index_col=None,na_values=['NA'])
        overview_row_count, overview_col_count, = overview_df.shape
        for row in overview_df.iloc[0:overview_row_count]:

            import pdb;pdb.set_trace()

    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)
        print("Exception-> "+str(e))


def get_workbook(complete_file_path):
    wb = None
    try:
        wb = xlrd.open_workbook(complete_file_path)
    except Exception as e:
        print("Exception occurred while opening workbook-> "+str(e))
        if wb:
            wb.release_resources()
            del wb
    finally:
        return wb


def read_sheet(workbook,sheet_name):
    table = {}
    try:
        sheet = workbook.sheet_by_name(sheet_name)
        keys = [sheet.cell_value(0, i) for i in range(0, sheet.ncols)]
        # sheet = wb.sheet_by_index(0)
        if sheet_name.lower() == 'overview':
            for i in range(1, sheet.nrows):
                table[str(sheet.cell_value(i, 0))] = {str('_'.join(x.lower() for x in heading.split(" "))):sheet.cell_value(i,keys.index(heading)) for heading in keys}
        else:
            levels = sheet.col_values(0,1)
            table = {level_name:[] for level_name in levels}
            for i in range(1, sheet.nrows):
                table[str(sheet.cell_value(i, 0))].append({
                str('_'.join(x.lower() for x in heading.split(" "))): sheet.cell_value(i, keys.index(heading)) for
                heading in keys})

    except Exception as e:
        print("Reading sheet-> '"+str(sheet_name)+"' raised exception -> "+str(e))
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)
        table = {}
    finally:
        return table


def get_sheet_names(workbook):
    sheets = None
    try:
        sheets = workbook.sheet_names()
    except Exception as e:
        print("Reading sheet names raised exception -> "+str(e))
    finally:
        return sheets


if __name__ == "__main__":

    wb = get_workbook('/home/shivendra/Documents/cd_excel.xlsx')
    overview = read_sheet(wb,'Overview')
    pdb.set_trace()
