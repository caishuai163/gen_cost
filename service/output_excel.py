import os

import openpyxl

from dao import basic_user
from service import cost_service, cost_other_service, a_b_service, p2p_service, cost_sum_service, bound_service, \
    spec_service, c7_service


def print_data():
    user_list = basic_user.get_bound_users()
    last_region = None
    region_group_cache = []
    for one_user in user_list:
        region_combine = one_user.region + "-" + one_user.region_manager
        if last_region is not None and last_region != region_combine:
            print_one_region(region_group_cache)
            region_group_cache = []
        region_group_cache.append(one_user)
        last_region = region_combine

    if len(region_group_cache) > 0:
        print_one_region(region_group_cache)


def print_one_region(user_list):
    # 遍历所有数组元素
    last_part = None
    part_arr = []
    for one_user in user_list:
        if one_user.work == "大区经理":
            last_part = None
            print("输出:" + one_user.region + "-" + one_user.region_manager)
            print_region_manager(one_user)
            # 执行大区经理
        else:
            if last_part is not None and last_part != one_user.part:
                # 执行分区分组
                print_region_part(part_arr)
                part_arr = []
            part_arr.append(one_user)
            last_part = one_user.part

    if len(part_arr) > 0:
        print_region_part(part_arr)


def print_region_manager(one_user):
    dir_name = "./output/" + one_user.region + "-" + one_user.region_manager
    if not os.path.exists(dir_name):
        os.mkdir(dir_name)

    xlsx_file = openpyxl.workbook.Workbook()
    sheet = xlsx_file.active

    cur_row, point_seq = cost_service.print_cost(sheet, user_list=(one_user,), column_num=1,
                                                 point_seq=1)

    cur_row, point_seq = cost_other_service.print_cost_other(sheet=sheet, user_list=(one_user,), cur_row=cur_row,
                                                             point_seq=point_seq)
    cur_row, point_seq = p2p_service.print_line(sheet=sheet, user_list=(one_user,), is_give=False, column_num=cur_row,
                                                point_seq=point_seq)

    cur_row, point_seq = p2p_service.print_line(sheet=sheet, user_list=(one_user,), is_give=True, column_num=cur_row,
                                                point_seq=point_seq)
    # A+B
    cur_row, point_seq = a_b_service.print_line(sheet=sheet, user_list=(one_user,), is_after=0, column_num=cur_row,
                                                point_seq=point_seq)

    cur_row, point_seq = a_b_service.print_line(sheet=sheet, user_list=(one_user,), is_after=1, column_num=cur_row,
                                                point_seq=point_seq)

    # 业务员名下电池扣款
    cur_row, point_seq = spec_service.print_line(sheet=sheet, user_list=(one_user,), column_num=cur_row,
                                                 point_seq=point_seq)
    # C类推广员百分之7推广奖励.xlsx
    cur_row, point_seq = c7_service.print_line(sheet=sheet, user_list=(one_user,), column_num=cur_row,
                                               point_seq=point_seq)

    cur_row, point_seq = cost_sum_service.print_line(sheet=sheet, user_list=(one_user,), column_num=cur_row,
                                                     point_seq=point_seq)
    cur_row, point_seq = bound_service.print_line(sheet=sheet, user_list=(one_user,), column_num=cur_row,
                                                  point_seq=point_seq)
    xlsx_file.save("output\\" + one_user.region + "-" + one_user.region_manager + "\\" + one_user.true_name + ".xlsx")


def print_region_part(user_list):
    dir_name = "./output/" + user_list[0].region + "-" + user_list[0].region_manager
    if not os.path.exists(dir_name):
        os.mkdir(dir_name)
    xlsx_file = openpyxl.workbook.Workbook()
    sheet = xlsx_file.active

    cur_row, point_seq = cost_service.print_cost(sheet, user_list=user_list, column_num=1,
                                                 point_seq=1)

    cur_row, point_seq = cost_other_service.print_cost_other(sheet=sheet, user_list=user_list, cur_row=cur_row,
                                                             point_seq=point_seq)
    cur_row, point_seq = p2p_service.print_line(sheet=sheet, user_list=user_list, is_give=False, column_num=cur_row,
                                                point_seq=point_seq)

    cur_row, point_seq = p2p_service.print_line(sheet=sheet, user_list=user_list, is_give=True, column_num=cur_row,
                                                point_seq=point_seq)
    # A+B
    cur_row, point_seq = a_b_service.print_line(sheet=sheet, user_list=user_list, is_after=0, column_num=cur_row,
                                                point_seq=point_seq)

    cur_row, point_seq = a_b_service.print_line(sheet=sheet, user_list=user_list, is_after=1, column_num=cur_row,
                                                point_seq=point_seq)

    # 业务员名下电池扣款
    cur_row, point_seq = spec_service.print_line(sheet=sheet, user_list=user_list, column_num=cur_row,
                                                 point_seq=point_seq)
    # C类推广员百分之7推广奖励.xlsx
    cur_row, point_seq = c7_service.print_line(sheet=sheet, user_list=user_list, column_num=cur_row,
                                               point_seq=point_seq)

    cur_row, point_seq = cost_sum_service.print_line(sheet=sheet, user_list=user_list, column_num=cur_row,
                                                     point_seq=point_seq)
    cur_row, point_seq = bound_service.print_line(sheet=sheet, user_list=user_list, column_num=cur_row,
                                                  point_seq=point_seq)
    xlsx_file.save("output\\" + user_list[0].region + "-" + user_list[0].region_manager + "\\" + user_list[0].part
                   + ".xlsx")
