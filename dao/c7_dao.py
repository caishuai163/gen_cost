import sqlite3

from config.excel_conf import DB_URL


class C7Tb:
    def __init__(self):
        self.code = None
        self.user_id = None
        self.month_val = None
        self.rmb = None


def insert(user_id, month, rmb):
    conn = sqlite3.connect(DB_URL)
    execute = conn.cursor()
    execute.execute("INSERT INTO c7_tb (user_id, month_val, rmb) VALUES(?, ?, ?);",
                    (user_id, month, rmb))
    conn.commit()
    execute.close()


def select_by_uid(user_id):
    conn = sqlite3.connect(DB_URL)
    execute = conn.cursor()
    execute.execute("select * from c7_tb where user_id = ? order by code asc;",
                    (user_id, ))
    res_list = execute.fetchall()
    execute.close()
    val_list = []
    for res_one in res_list:
        e = C7Tb()
        e.code = res_one[0]
        e.user_id = res_one[1]
        e.month_val = res_one[2]
        e.rmb = res_one[3]
        val_list.append(e)
    return val_list


def select_month_range(uid_list):
    """
    获取数据的月份范围
    :param uid_list: uid 列表
    :param data_type_list MonthDataType list
    :return: 月份中文数组列表
    """
    sql = """select min(dt.code),max(dt.code) from c7_tb cdt 
join date_tb dt on dt.date_str = cdt.month_val  
where cdt.user_id  in ("""
    for i in range(len(uid_list)):
        uid = uid_list[i]
        if i != 0:
            sql = sql + ","
        sql = sql + str(uid)
    sql = sql + ")"

    conn = sqlite3.connect(DB_URL)
    execute = conn.cursor()
    execute.execute(sql)
    exec_data = execute.fetchone()
    min_code = exec_data[0]
    max_code = exec_data[1]
    execute.execute("select date_str from date_tb where code >= ? and code <= ? order by code asc",
                    (min_code, max_code))
    date_list = execute.fetchall()
    str_list = []
    for data_one in date_list:
        str_list.append(data_one[0])
    return str_list
