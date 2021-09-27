import pandas as pd
import sqlite3


def read_excel(excel, sheetname):
    df = pd.read_excel(io=excel, sheet_name=sheetname, header=0, engine='openpyxl')
    print(df)
    engine = sqlite3.connect('atthelper.db')
    df.to_sql(name='attendance', con=engine, if_exists='replace')


if __name__ == '__main__':
    read_excel(excel='原始打卡记录导出.xlsx', sheetname='原始打卡记录')
