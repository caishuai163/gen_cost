import os
import shutil

from config import excel_conf
from dao import date_dao, init
from service import cost_service, p2p_service, a_b_service, output_excel, bound_service, user_service


def main():
    init.init()
    # 重新加载user_tb, 从excel中
    user_service.reload_file()
    if os.path.exists("./output"):
        shutil.rmtree("./output")
    os.mkdir("./output")
    excel_conf.init()
    init.clear_db()

    date_dao.init()
    cost_service.load_cost()

    p2p_service.load_ab()
    a_b_service.load_ab()

    # 加载奖金
    bound_service.load_bonus()

    output_excel.print_data()


if __name__ == '__main__':
    main()
