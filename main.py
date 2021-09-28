# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
from os import walk
from shutil import copy
from pdUtil import *
from openpyxl import load_workbook


def replace_excel_title(excel, sheetname):
    workbook = load_workbook(excel)
    sheet = workbook[sheetname]
    sheet.cell(1, 1).value = 'name'
    sheet.cell(1, 2).value = 'company'
    sheet.cell(1, 3).value = 'service'
    sheet.cell(1, 4).value = 'atttime'
    sheet.cell(1, 5).value = 'terminal'
    workbook.save(excel)


def get_source_excel():
    for root, dirs, files in walk('.'):
        for fl in files:
            if fl.endswith('.xlsx'):
                return fl


def backup_source_excel(source):
    dst = source.split('.xlsx')[0] + '_备份.xlsx'
    copy(source, dst)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    excel = get_source_excel()
    backup_source_excel(excel)
    src_sheetname = '原始打卡记录'
    replace_excel_title(excel=excel, sheetname=src_sheetname)
    read_excel_sqlite(excel='原始打卡记录导出.xlsx', sheetname='原始打卡记录')
    res_excel = excel.split('.xlsx')[0] + '_结果.xlsx'
    count = write_with_start_end(excel=res_excel, sheetname='打卡记录')
    write_only_start_or_end(excel=res_excel, sheetname='打卡记录', startrow=count)
