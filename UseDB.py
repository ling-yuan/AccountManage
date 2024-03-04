import sqlite3


class MyDB:

    #初始化对象
    def __init__(self):
        # conn = sqlite3.connect("./DataBases/test.db")
        conn = sqlite3.connect("database.db")
        cur = conn.cursor()
        # sqlCreateInformation = "CREATE TABLE information(id NUMBER,name TEXT,account TEXT,password TEXT);"
        sqlCreateInformation = '''
            CREATE TABLE IF NOT EXISTS `information`(
               `id` INTEGER PRIMARY KEY,
               `name` VARCHAR(100) NOT NULL,
               `account` CHAR,
               `password` CHAR,
               `remark` TEXT
            );
            '''
        cur.execute(sqlCreateInformation)
        conn.commit()
        conn.close()

    #创建连接对象
    def getConnection(self):
        # conn = sqlite3.connect("./DataBases/test.db")
        conn = sqlite3.connect("database.db")
        # conn = sqlite3.connect(':memory:')
        cur = conn.cursor()
        return conn, cur

    #查询，删除，更新
    def executeUpdate(self, sql, args=None):
        conn, cur = self.getConnection()
        rows = None
        if args == None:
            rows = cur.execute(sql)
        else:
            if isinstance(args, tuple):
                rows = cur.execute(sql, args)
            elif isinstance(args, list):
                rows = cur.executemany(sql, args)
        conn.commit()
        conn.close()
        return rows

    #查询
    def executeQuery(self, sql):
        conn, cur = self.getConnection()
        cur.execute(sql)
        result = cur.fetchall()
        conn.close()
        return result
