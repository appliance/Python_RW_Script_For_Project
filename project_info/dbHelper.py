import pymysql
from db_config import db_config

class DBHelper:

    def __init__(self):
        self.conn = None
        self.cur = None


    def connect_database(self):
        try:
            self.conn = pymysql.connect(db_config['host'], db_config['username'],
                                        db_config['password'], db_config['database'],
                                        charset=db_config['charset'])
        except:
            print('connnection failed')
            return False
        self.cur = self.conn.cursor()
        return True


    # 关闭数据库
    def close_database(self):
        if self.conn and self.cur:
            self.cur.close()
            self.conn.close()
        return True


    # 执行sql语句
    def excute(self, sql, params):
        try:
            if self.conn and self.cur:
                self.cur.execute(sql, params)
                self.conn.commit()
        except:
            print("execute failed: " + sql)
            print("params: " + params)
            self.close_database()
            return False
        return True


    # 检索数据库
    def select(self, sql, params):
        try:
            if self.conn and self.cur:
                self.cur.execute(sql, params)
                result = self.cur.fetchall()
                return result
        except:
            print("execute failed: " + sql)
            print("params: " + params)
            self.close_database()
            return ''



