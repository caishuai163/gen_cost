import sqlite3

from config import excel_conf
from config.excel_conf import DB_URL
from dao import date_dao


class CheckCost:
    def __init__(self):
        self.code = None
        self.user_id = None
        self.file_name = None
        # 盘点后丢失电池数量
        self.confirm_lost = None
        # 扣款金额
        self.avg_cost = None
        # 扣款月数
        self.cost_month = None
        # 月扣款金额
        self.month_cost = None
        # 找回数量
        self.find_count = None
        # 盘点后丢失电池数量 找回数量后面那列
        self.confirm_after_find = None


class CheckCostDetail:
    def __init__(self):
        self.code = None
        self.cost_code = None
        self.month_val = None
        self.rmb = None


def insertCost(user_id, file_name, confirm_lost, find_count, confirm_after_find, avg_cost, cost_month, month_cost):
    conn = sqlite3.connect(DB_URL)
    execute = conn.cursor()
    execute.execute("INSERT INTO cost_tb(user_id, file_name, confirm_lost, avg_cost, cost_month, "
                    "month_cost, find_count, confirm_after_find) VALUES(?, ?, ?, ?, ?, ?, ?, ?);",
                    (
                        user_id, file_name, confirm_lost, avg_cost, cost_month, month_cost, find_count,
                        confirm_after_find))
    conn.commit()
    execute.execute("select max(code) from cost_tb")
    code = execute.fetchone()[0]
    execute.close()
    return code


def insertCostDetail(code, month, rmb):
    conn = sqlite3.connect(DB_URL)
    execute = conn.cursor()
    execute.execute("INSERT INTO cost_detail_tb (cost_code, month_val, rmb) VALUES(?, ?, ?);",
                    (code, month, rmb))
    conn.commit()
    execute.close()


def select_month_range(uid_list):
    sql = """select min(dt.code),max(dt.code) from cost_tb ct 
join cost_detail_tb cdt on ct.code =cdt.cost_code 
join date_tb dt on dt.date_str = cdt.month_val  
where ct.user_id  in ("""

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

    if len(str_list) == 0:
        str_list = excel_conf.default_month
    if len(str_list) < len(excel_conf.default_month):
        str_list = date_dao.append_pre(str_list, len(excel_conf.default_month) - len(str_list))

    return str_list


def select_cost_list_title_data(user_id, file_name):
    conn = sqlite3.connect(DB_URL)
    execute = conn.cursor()
    execute.execute("select * from cost_tb where user_id = ? and file_name = ? order by code asc;",
                    (user_id, file_name))
    res_list = execute.fetchall()
    execute.close()
    cost_list = []
    for res_one in res_list:
        e = CheckCost()
        e.code = res_one[0]
        e.user_id = res_one[1]
        e.confirm_lost = res_one[3]
        e.avg_cost = res_one[4]
        e.cost_month = res_one[5]
        e.month_cost = res_one[6]
        e.find_count = res_one[7]
        e.confirm_after_find = res_one[8]
        cost_list.append(e)
    return cost_list


def select_by_user_and_file(user_id, file_name):
    conn = sqlite3.connect(DB_URL)
    execute = conn.cursor()
    execute.execute("select * from cost_tb where user_id = ? and file_name = ? order by code asc;",
                    (user_id, file_name))
    res_one = execute.fetchone()
    execute.close()
    if res_one is None:
        return None
    if res_one[0] is None:
        return None
    e = CheckCost()
    e.code = res_one[0]
    e.user_id = res_one[1]
    e.confirm_lost = res_one[3]
    e.avg_cost = res_one[4]
    e.cost_month = res_one[5]
    e.month_cost = res_one[6]
    e.find_count = res_one[7]
    e.confirm_after_find = res_one[8]
    return e


def select_detail_by_cost_code(code):
    conn = sqlite3.connect(DB_URL)
    execute = conn.cursor()
    execute.execute("select * from cost_detail_tb where cost_code = ? order by code asc;",
                    (code,))
    res_list = execute.fetchall()
    execute.close()
    cost_detail_list = []
    for res_one in res_list:
        e = CheckCostDetail()
        e.code = res_one[0]
        e.cost_code = res_one[1]
        e.month_val = res_one[2]
        e.rmb = res_one[3]
        cost_detail_list.append(e)
    return cost_detail_list


def select_sum_month_range(uid_list):
    print("cost.py" + "select_sum_month_range" + str(uid_list))
    sql = """select min(dt.code),max(dt.code) from cost_tb ct 
join cost_detail_tb cdt on ct.code =cdt.cost_code 
join date_tb dt on dt.date_str = cdt.month_val  
where ct.user_id  in ("""

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
    print(exec_data)
    min_code = exec_data[0]
    max_code = exec_data[1]

    sql = """select min(dt.code),max(dt.code) from p2p_tb ppt  
        join date_tb dt on dt.date_str = ppt.cost_month 
        where ppt.user_id  in ("""

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
    if exec_data is not None:
        print(exec_data)
        if min_code is None or max_code is None:
            if exec_data[0] is not None:
                min_code = exec_data[0]
            if exec_data[1] is not None:
                max_code = exec_data[1]
        else:
            if exec_data[0] is not None and exec_data[0] < min_code:
                min_code = exec_data[0]
            if exec_data[1] is not None and exec_data[1] > max_code:
                max_code = exec_data[1]

    sql = """select min(dt.code),max(dt.code) from spec_tb ppt  
        join date_tb dt on dt.date_str = ppt.cost_month 
        where ppt.user_id  in ("""

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
    if exec_data is not None:
        if min_code is None or max_code is None:
            if exec_data[0] is not None:
                min_code = exec_data[0]
            if exec_data[1] is not None:
                max_code = exec_data[1]
        else:
            if exec_data[0] is not None and exec_data[0] < min_code:
                min_code = exec_data[0]
            if exec_data[1] is not None and exec_data[1] > max_code:
                max_code = exec_data[1]

    execute.execute("select date_str from date_tb where code >= ? and code <= ? order by code asc",
                    (min_code, max_code))
    date_list = execute.fetchall()
    str_list = []
    for data_one in date_list:
        str_list.append(data_one[0])

    if len(str_list) == 0:
        str_list = excel_conf.default_month
    if len(str_list) < len(excel_conf.default_month):
        str_list = date_dao.append_pre(str_list, len(excel_conf.default_month) - len(str_list))
    return str_list
