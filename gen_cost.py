import os
import shutil
import sqlite3

from config import excel_conf
from config.excel_conf import DB_URL
from dao import date_dao
from service import cost_service, p2p_service, a_b_service, output_excel, bound_service, user_service


def main():
    # 重新加载user_tb, 从excel中
    # user_service.reload_file()
    if os.path.exists("./output"):
        shutil.rmtree("./output")
    os.mkdir("./output")
    excel_conf.init()
    clear_db()

    date_dao.init()
    cost_service.load_cost()

    p2p_service.load_ab()
    a_b_service.load_ab()

    # 加载奖金
    bound_service.load_bonus()

    output_excel.print_data()


def clear_db():
    print("clear old data")
    conn = sqlite3.connect(DB_URL)
    execute = conn.cursor()
    execute.execute("delete from bound_tb;")
    execute.execute("delete from cost_tb;")
    execute.execute("delete from cost_detail_tb;")

    execute.execute("delete from p2p_tb;")
    execute.execute("delete from ab_tb;")
    conn.commit()
    execute.close()
    print("clear old data success")


if __name__ == '__main__':
    main()
