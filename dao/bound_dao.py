import sqlite3

from config.excel_conf import DB_URL


class BoundTb:
    def __init__(self):
        self.code = None
        self.user_id = None
        self.title_id = None
        self.rmb = None


def insert(uid, title_id, rmb):
    conn = sqlite3.connect(DB_URL)
    execute = conn.cursor()
    execute.execute("INSERT INTO bound_tb (user_id, title_id, rmb) VALUES(?, ?, ?);",
                    (uid, title_id, rmb))
    conn.commit()
    execute.close()


def select_bound_by_uid(uid):
    bound_list = []
    conn = sqlite3.connect(DB_URL)
    execute = conn.cursor()
    execute.execute("select * from bound_tb where user_id = ? order by title_id", (uid,))
    res_list = execute.fetchall()
    execute.close()
    for res_obj in res_list:
        e = BoundTb()
        e.code = res_obj[0]
        e.user_id = res_obj[1]
        e.title_id = res_obj[2]
        e.rmb = res_obj[3]
        bound_list.append(e)
    return bound_list
