from pandas import read_excel, read_sql
from sqlite3 import connect
from openpyxl import load_workbook
from datetime import datetime, timedelta
from logUtil import init_logging
logger = init_logging()


class PdUtilV2(object):
    def __init__(self):
        self.db = 'atthelper.db'

    def db_operator(self, sql, op):
        """
        数据库操作，
        :param sql:  sql语句
        :param op:  操作类型， s：查询， iu：insert update
        :return: query result or None
        """
        logger.info('SQL语句为: %s' % sql)
        conn = connect(self.db)
        cursor = conn.cursor()
        res = ''
        if 's' == op:
            logger.info('走查询流程-------------------------')
            res = cursor.execute(sql).fetchall()
            logger.info('SQL查询结果为: %s' % res)
        elif 'iu' == op:
            logger.info('走insert、update流程-------------------------')
            cursor.execute(sql)
            conn.commit()
            res = None
        cursor.close()
        conn.close()
        return res

    def read_excel_to_sqlite(self, excel, sheet):
        """
        将excel读取到sqlite
        :param excel:  excel路径
        :param sheet: 工作表名称
        :return: 无
        """
        df = read_excel(io=excel, sheet_name=sheet, header=0, engine='openpyxl')
        engine = connect(self.db)
        df.to_sql(name='attendance', con=engine, if_exists='replace')

    def update_yesterday(self, atttuple):
        """
        将当天打卡时间早于5点的打卡数据作为头一天的下班打卡记录
        :param atttuple:
        :return:
        """
        # 将传入的日期减一，并只取日期
        to_datetime = datetime.strptime(atttuple[-1], '%Y-%m-%d %H:%M:%S') + timedelta(days=-1)
        yesterday = to_datetime.date()
        logger.info('头一天的日期为: %s' % yesterday)

        # 查询传入的人员头一天的考勤记录sql
        yesterday_his_sql = "select name, service, end from attendance_result where name='{name}' and service='{service}' and strftime('%Y-%m-%d', end) = '{yest}';".format(
            name=atttuple[0], service=atttuple[1], yest=yesterday)
        query_res = self.db_operator(sql=yesterday_his_sql, op='s')
        logger.info('头一天的考勤记录查询结果为: %s' % query_res)

        if len(query_res) == 1:
            update_sql = "update attendance_result set end='{endtime}' where name='{name}' and service='{service}' and strftime('%Y-%m-%d %H:%M:%S', end)='{attime}';".format(
                endtime=atttuple[-1], name=atttuple[0], service=atttuple[1], attime=query_res[0][-1])
            self.db_operator(sql=update_sql, op='iu')
        else:
            logger.error('头一天的考勤记录查询结果不唯一')

    def deal_bf5(self):
        bf5_sql = "select name, service, start from attendance_result where strftime('%H', start) < '05';"
        res = self.db_operator(sql=bf5_sql, op='s')

        # [('剁椒', 'Z00003', '2021-09-08 04:19:31')]
        if res is not None and len(res) != 0:
            for i in range(len(res)):
                self.update_yesterday(res[i])

    def deal_to_result_table(self):
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

        engine = connect(self.db)
        df = read_sql(sql=sql, con=engine)
        df.to_sql(name='attendance_result', con=engine, if_exists='replace')

    def update_start_or_end(self):
        """
        更新只有一次打卡记录的数据
        早于18点按上班打卡处理
        晚于18点按下班打卡处理
        :return:
        """
        conn = connect(self.db)
        cursor = conn.cursor()
        update_start = """update attendance_result set start=null where start=end and strftime('%H', start) >= 18;"""
        update_end = """update attendance_result set end=null where start=end and strftime('%H', start)<18;"""

        self.db_operator(sql=update_start, op='iu')
        self.db_operator(sql=update_end, op='iu')

    def write_to_excel(self, excel, sheetname='打卡记录', startrow=1):
        sql = """select t2.name as 姓名, t2.company as 公司, t2.service as 外包服务编号, t2.start as 上班时间, t2.end as 下班时间, t2.terminal as 终端, t2.attdate as 打卡日期
    from attendance_result t2
    order by t2.name;"""

        conn = connect(self.db)

        df = read_sql(sql=sql, con=conn)
        df.to_excel(excel_writer=excel, sheet_name=sheetname)


if __name__ == '__main__':
    pd2 = PdUtilV2()
    pd2.deal_bf5()