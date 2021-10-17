from pandas import read_excel, read_sql
from sqlite3 import connect
from openpyxl import load_workbook


def read_excel_to_sqlite(excel, sheetname):
    df = read_excel(io=excel, sheet_name=sheetname, header=0, engine='openpyxl')
    print(df)
    engine = connect('atthelper.db')
    df.to_sql(name='attendance', con=engine, if_exists='replace')


def deal_to_sqlite():
    """
    查同时有上班打卡和下班打开记录的考勤信息
    :param excel:
    :param sheetname:
    :return:
    """
    sql = """select t2.name, t2.company, t2.service, t2.start, t2.end, t2.terminal, t2.attdate
from (select t.name, t.company, t.service, min(t.atttime) as start, max(t.atttime) as end, t.terminal, strftime('%Y-%m-%d', t.atttime) as attdate
      from attendance t
      group by name, strftime('%Y%m%d', t.atttime)
      order by name) t2
order by name;"""

    engine = connect('atthelper.db')
    df = read_sql(sql=sql, con=engine)
    df.to_sql(name='attendance_result', con=engine, if_exists='replace')


def update_start_or_end():
    """
    更新只有一次打卡记录的数据
    :return:
    """
    conn = connect('atthelper.db')
    cursor = conn.cursor()
    update_start = """update attendance_result set start=null where start=end and strftime('%H', start) >= 18;"""
    update_end = """update attendance_result set end=null where start=end and strftime('%H', start)<18;"""
    cursor.execute(update_start)
    conn.commit()

    cursor.execute(update_end)
    conn.commit()

    cursor.close()
    conn.close()


def update_midnight():
    conn = connect('atthelper.db')
    cursor = conn.cursor()
    update_start = """update attendance_result set start=null where start=end and strftime('%H', start) >= 18;"""
    update_end = """update attendance_result set end=null where start=end and strftime('%H', start)<18;"""
    cursor.execute(update_start)
    conn.commit()

    cursor.execute(update_end)
    conn.commit()

    cursor.close()
    conn.close()


def write_to_excel(excel, sheetname='打卡记录', startrow=1):
    sql = """select t2.name as 姓名, t2.company as 公司, t2.service as 外包服务编号, t2.start as 上班时间, t2.end as 下班时间, t2.terminal as 终端, t2.attdate as 打卡日期
from attendance_result t2
order by t2.name;"""

    conn = connect('atthelper.db')

    df = read_sql(sql=sql, con=conn)
    df.to_excel(excel_writer=excel, sheet_name=sheetname)



if __name__ == '__main__':
    read_excel_to_sqlite(excel='原始打卡记录导出.xlsx', sheetname='原始打卡记录')
    deal_to_sqlite()
    update_midnight()
    write_to_excel(excel='ttt')
