import os
import shutil
import traceback

from config import excel_conf
from dao import date_dao, init
from service import cost_service, p2p_service, a_b_service, output_excel, bound_service, user_service, spec_service, \
    c7_service


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
    spec_service.load_spec()
    c7_service.load_c7()

    # 加载奖金
    bound_service.load_bonus()

    output_excel.print_data()


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print("计算出错，抛出异常")
        traceback.print_exc()
        print(e)

    # 等待用户输入
    user_input = input("计算结束，按回车结束: ")

