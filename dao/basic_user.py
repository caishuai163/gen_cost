import sqlite3

from config.excel_conf import DB_URL


class UserTb:
    def __init__(self):
        self.code = None
        self.true_name = None
        self.region = None
        self.region_manager = None
        self.part = None
        self.company = None
        self.work = None
        self.dealer_name = None


def insert(user_name, region, region_manager, part, company, work, dealer_name, platform, need_generate):

    conn = sqlite3.connect(DB_URL)
    execute = conn.cursor()
    execute.execute("""
    INSERT INTO user_tb
(user_name, region, region_manager, part, company, "work", dealer_name, platform, need_generated)
VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?);
""",
                    (user_name, region, region_manager, part, company, work, dealer_name, platform, need_generate))
    conn.commit()


def getUserInfoByName(user_name, company, work):

    conn = sqlite3.connect(DB_URL)
    execute = conn.cursor()
    execute.execute("select * from user_tb where user_name = ?",
                    (str(user_name),))
    result_list = execute.fetchall()
    execute.close()
    if len(result_list) == 0:
        return None
    u = UserTb()
    if len(result_list) == 1:
        result_one = result_list[0]
        u.code = result_one[0]
        u.true_name = result_one[1]
        u.region = result_one[2]
        u.region_manager = result_one[3]
        u.part = result_one[4]
        u.company = result_one[5]
        u.work = result_one[6]
        u.dealer_name = result_one[7]
        return u

    tmp_list = []
    for result_one in result_list:
        if result_one[6] == work:
            tmp_list.append(result_one)

    if len(tmp_list) == 0:
        print("user_tb相同名字存在多个，通过职位未成功匹配."+user_name)
        return None
    if len(tmp_list) == 1:
        result_one = tmp_list[0]
        u.code = result_one[0]
        u.true_name = result_one[1]
        u.region = result_one[2]
        u.region_manager = result_one[3]
        u.part = result_one[4]
        u.company = result_one[5]
        u.work = result_one[6]
        u.dealer_name = result_one[7]
        return u

    tmp_list_2 = []
    for result_one in tmp_list:
        if result_one[5] == company:
            tmp_list_2.append(result_one)

    if len(tmp_list_2) == 0:
        print("user_tb相同名字存在多个，通过分区未成功匹配."+user_name)
        return None
    if len(tmp_list_2) == 1:
        result_one = tmp_list_2[0]
        u.code = result_one[0]
        u.true_name = result_one[1]
        u.region = result_one[2]
        u.region_manager = result_one[3]
        u.part = result_one[4]
        u.company = result_one[5]
        u.work = result_one[6]
        u.dealer_name = result_one[7]
        return u
    print("user_tb相同名字存在多个，通过职位+分区未成功匹配." + user_name)
    return None


def getUserInfoByNameAndPart(user_name, region, region_manager, part, work):

    conn = sqlite3.connect(DB_URL)
    execute = conn.cursor()
    execute.execute("select * from user_tb where user_name = ?",
                    (str(user_name),))
    result_list = execute.fetchall()
    execute.close()
    user_list = convert_bean(result_list)
    if len(user_list) == 0:
        return None
    if len(user_list) == 1:
        return user_list[0]

    tmp_list = []
    for user_one in user_list:
        if user_one.work == work:
            tmp_list.append(user_one)

    if len(tmp_list) == 0:
        print("相同名字存在多个，通过职位未成功匹配."+user_name)
        return None
    if len(tmp_list) == 1:
        return tmp_list[0]

    tmp_list_2 = []
    for user_one in tmp_list:
        if region_manager is None:
            if user_one.region == region:
                tmp_list_2.append(user_one)
        else:
            if user_one.region == region and user_one.region_manager == region_manager:
                tmp_list_2.append(user_one)

    if len(tmp_list_2) == 0:
        print("相同名字存在多个，通过分区未成功匹配."+user_name)
        return None
    if len(tmp_list_2) == 1:
        return tmp_list_2[0]

    tmp_list_3 = []
    for user_one in tmp_list_2:
        if user_one.part == part:
            tmp_list_3.append(user_one)

    if len(tmp_list_3) == 0:
        print("相同名字存在多个，通过分区未成功匹配."+user_name)
        return None
    if len(tmp_list_3) == 1:
        return tmp_list_2[0]
    print("相同名字存在多个，通过职位+分区未成功匹配." + user_name)
    return None


def getUserInfoByNameAndRegion(user_name, region):
    conn = sqlite3.connect(DB_URL)
    execute = conn.cursor()
    execute.execute("select * from user_tb where user_name = ? and region = ?",
                    (str(user_name), str(region),))
    result_list = execute.fetchall()
    execute.close()
    user_list = convert_bean(result_list)
    if len(user_list) == 0:
        return None
    if len(user_list) == 1:
        return user_list[0]
    print("相同名字存在多个，通过名字+大区未成功匹配." + user_name)
    return None


def get_bound_users():
    conn = sqlite3.connect(DB_URL)
    execute = conn.cursor()
    execute.execute("""select ut.* from user_tb ut 
    where need_generated = 1
    order by region asc,    
    region_manager asc,
    case when `work` ='大区经理' then 0 else 1 end asc,
    part asc,
    case when `work` ='分区经理' then 0 else 1 end asc,
    `work` asc""")
    result_list = execute.fetchall()
    execute.close()
    user_list = []
    for result_one in result_list:
        u = UserTb()
        u.code = result_one[0]
        u.true_name = result_one[1]
        u.region = result_one[2]
        u.region_manager = result_one[3]
        u.part = result_one[4]
        u.company = result_one[5]
        u.work = result_one[6]
        u.dealer_name = result_one[7]
        user_list.append(u)
    return user_list


def convert_bean(res_list):
    user_list = []
    for result_one in res_list:
        u = UserTb()
        u.code = result_one[0]
        u.true_name = result_one[1]
        u.region = result_one[2]
        u.region_manager = result_one[3]
        u.part = result_one[4]
        u.company = result_one[5]
        u.work = result_one[6]
        u.dealer_name = result_one[7]
        user_list.append(u)
    return user_list


def remove_all():
    conn = sqlite3.connect(DB_URL)
    execute = conn.cursor()
    execute.execute("delete from user_tb;")
    conn.commit()
    execute.close()
