import sqlite3
from datetime import datetime

from dateutil.relativedelta import relativedelta

from config.excel_conf import DB_URL


def init():
    print("init date")
    end = datetime.now() + relativedelta(months=24)
    conn = sqlite3.connect(DB_URL)
    execute = conn.cursor()
    execute.execute("select max(date) from date_tb ;")
    date_str = execute.fetchone()[0]
    if date_str is None:
        start = datetime.fromisoformat("2021-12-01")
    else:
        start = datetime.fromisoformat(date_str)
    while True:
        start = start + relativedelta(months=1)
        if start.timestamp() > end.timestamp():
            break
        execute.execute("INSERT INTO date_tb (date_str,date) VALUES(?,?);",
                        (str(start.year) + "年" + str(start.month) + "月份", start))
    conn.commit()
    execute.close()
    print("init date success")


def append_pre(arr, fill):
    conn = sqlite3.connect(DB_URL)
    execute = conn.cursor()
    min_str = arr[0]
    execute.execute("select code from date_tb  where date_str = ?;", (min_str,))
    code = execute.fetchone()[0]
    pre_code = code - fill
    execute.execute("select date_str from date_tb where code >= ? and code < ? order by code asc",
                    (pre_code, code))
    date_list = execute.fetchall()

    execute.close()
    str_list = []
    for data_one in date_list:
        str_list.append(data_one[0])
    str_list.extend(arr)
    return str_list
