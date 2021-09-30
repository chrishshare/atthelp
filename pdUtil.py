from pandas import read_excel, read_sql
from sqlite3 import connect
from openpyxl import load_workbook


def read_excel_sqlite(excel, sheetname):
    df = read_excel(io=excel, sheet_name=sheetname, header=0, engine='openpyxl')
    print(df)
    engine = connect('atthelper.db')
    df.to_sql(name='attendance', con=engine, if_exists='replace')


def write_with_start_end(excel, sheetname):
    sql = """select t2.name as 姓名, t2.company as 公司, t2.service as 外包服务编号, t2.start as 上班时间, t2.end as 下班时间, t2.terminal as 终端, t2.attdate as 上班日期
from (select t.name, t.company, t.service, min(t.atttime) as start, max(t.atttime) as end, t.terminal, strftime('%Y-%m-%d', t.atttime) as attdate
      from attendance t
      group by name, strftime('%Y%m%d', t.atttime)
      order by name) t2
where t2.start <> t2.end
order by name;"""
    engine = connect('atthelper.db')
    df = read_sql(sql=sql, con=engine)
    df.to_excel(excel_writer=excel, sheet_name=sheetname)
    return df.shape[0]


def write_only_start_or_end(excel, sheetname, startrow):
    sql = """select t2.name as 姓名, t2.company as 公司, t2.service as 外包服务编号, t2.start as 上班时间, t2.end as 下班时间, t2.terminal as 终端, t2.attdate as 上班日期
from (select t.name, t.company, t.service, min(t.atttime) as start, max(t.atttime) as end, t.terminal, strftime('%Y-%m-%d', t.atttime) as attdate
      from attendance t
      group by name, strftime('%Y%m%d', t.atttime)
      order by name) t2
where t2.start = t2.end
order by name;"""
    engine = connect('atthelper.db')
    df = read_sql(sql=sql, con=engine)

    workbook = load_workbook(excel)
    # sheet = workbook.get_sheet_by_name(sheetname)
    sheet = workbook[sheetname]

    for index, rows in df.iterrows():
        row_nu = startrow + 2 + index
        # name
        sheet.cell(row_nu, 2).value = rows[0]
        # company
        sheet.cell(row_nu, 3).value = rows[1]
        # service
        sheet.cell(row_nu, 4).value = rows[2]

        # terminal
        sheet.cell(row_nu, 7).value = rows[5]
        # 上班日期
        sheet.cell(row_nu, 8).value = rows[6]

        if int(rows[3][11:13]) < 18:
            # starttime
            sheet.cell(row_nu, 5).value = rows[3]
            # endtime
            sheet.cell(row_nu, 6).value = ''
        else:
            # starttime
            sheet.cell(row_nu, 5).value = ''
            # endtime
            sheet.cell(row_nu, 6).value = rows[3]

    sheet.delete_cols(1)
    workbook.save(excel)


if __name__ == '__main__':
    read_excel_sqlite(excel='原始打卡记录导出.xlsx', sheetname='原始打卡记录')
    count = write_with_start_end(excel='result.xlsx', sheetname='打卡记录')
    write_only_start_or_end(excel='result.xlsx', sheetname='打卡记录', startrow=count)
