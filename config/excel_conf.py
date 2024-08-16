# 边框
import json

from openpyxl.styles import Border, Font, Side

border = Border(left=Side(border_style='thin', color='000000'),
                right=Side(border_style='thin', color='000000'),
                top=Side(border_style='thin', color='000000'),
                bottom=Side(border_style='thin', color='000000'))

# 定义新的小号字体
fontStyle = Font(size="8")

DB_URL = "local.db"
file_url = "./config.json"
default_month = ["2023年5月份", "2023年6月份", "2023年7月份", "2023年8月份", "2023年9月份", "2023年10月份", "2023年11月", "2023年12月"]
json_data = {}


def init():
    with open(file_url, 'r', encoding='utf-8') as f:
        data = json.load(f)
        json_data['data'] = data


def get_title_template():
    return json_data['data']["title_template"]


def get_table_row_find_day():
    return json_data['data']["table_row_find_day"]


class CostFileTitleConf:
    def __init__(self):
        self.file_name = None
        self.row_title = None
        self.read_sheet_name = None

        self.row_table_title = None
        self.sum_row_text = None
        self.true_name = None
        self.company = None
        self.work = None
        self.confirm_lost = None
        self.find_count = None
        self.confirm_after_find = None
        self.avg_cost = None
        self.cost_month = None
        self.month_cost = None
        self.month_start = None


def get_cost_file_title_conf():
    conf_dict = {}
    conf_obj = json_data['data']["cost"]
    for conf_one in conf_obj:
        conf_data = CostFileTitleConf()
        conf_data.file_name = conf_one["file_name"]
        conf_data.row_title = conf_one["row_title"]
        conf_data.row_table_title = conf_one["row_table_title"]
        conf_data.sum_row_text = conf_one["sum_row_text"]
        conf_data.read_sheet_name = conf_one["read_sheet_name"]
        conf_data.true_name = conf_one["true_name"]
        conf_data.company = conf_one["company"]
        conf_data.work = conf_one["work"]
        conf_data.confirm_lost = conf_one["confirm_lost"]
        conf_data.find_count = conf_one["find_count"]
        conf_data.confirm_after_find = conf_one["confirm_after_find"]
        conf_data.avg_cost = conf_one["avg_cost"]
        conf_data.cost_month = conf_one["cost_month"]
        conf_data.month_cost = conf_one["month_cost"]
        conf_data.month_start = conf_one["month_start"]

        conf_dict[conf_one["file_name"]] = conf_data
    return conf_dict


def get_title_bound_from_day():
    return json_data['data']["title_bound_from_day"]


def get_title_other_cost_from_day():
    return json_data['data']["title_other_cost_from_day"]


def get_cost_base_file():
    return json_data['data']["cost_base_file"]


class ABFileConf:
    def __init__(self):
        self.file_name = None
        self.file_month = None
        self.read_pre_sheet_name = None
        self.read_after_sheet_name = None
        self.read_bound_row = None
        self.read_region_row = None
        self.read_part_row = None
        self.read_name_row = None
        self.read_work_row = None
        self.read_start_row = None
        self.data_start_row = 5


def get_ab_file_conf():
    conf_dict = {}
    conf_obj = json_data['data']["ab"]
    for conf_one in conf_obj:
        conf_data = ABFileConf()
        conf_data.file_name = conf_one["file_name"]
        conf_data.file_month = conf_one["file_month"]
        conf_data.read_pre_sheet_name = conf_one["read_pre_sheet_name"]
        conf_data.read_after_sheet_name = conf_one["read_after_sheet_name"]
        conf_data.read_bound_row = conf_one["read_bound_row"]
        conf_data.read_region_row = conf_one["read_region_row"]
        conf_data.read_part_row = conf_one["read_part_row"]
        conf_data.read_name_row = conf_one["read_name_row"]
        conf_data.read_work_row = conf_one["read_work_row"]
        conf_data.read_start_row = conf_one["read_start_row"]
        conf_dict[conf_one["file_name"]] = conf_data
    return conf_dict


class P2pFileConf:
    def __init__(self):
        self.file_name = None
        self.file_month = None
        self.read_sheet_name = None
        self.read_bound_row = None
        # "read_give_row": "实际奖励金额",
        self.read_give_row = None
        # "read_cost_row": "实际罚款金额",
        self.read_cost_row = None
        # "read_cost_month_row": "扣款月份"
        self.read_cost_month_row = None


def get_p2p_file_conf():
    conf_dict = {}
    conf_obj = json_data['data']["p2p"]
    for conf_one in conf_obj:
        conf_data = P2pFileConf()
        conf_data.file_name = conf_one["file_name"]
        conf_data.file_month = conf_one["file_month"]
        conf_data.read_sheet_name = conf_one["read_sheet_name"]
        conf_data.read_give_row = conf_one["read_give_row"]
        conf_data.read_cost_row = conf_one["read_cost_row"]
        conf_data.read_cost_month_row = conf_one["read_cost_month_row"]
        conf_data.sum_row_text = conf_one["sum_row_text"]
        conf_dict[conf_one["file_name"]] = conf_data
    return conf_dict


class SpecFileConf:
    def __init__(self):
        self.file_name = None
        self.file_month = None
        self.read_sheet_name = None
        self.read_bound_row = None
        # "region_row": "实际奖励金额",
        self.region_row = None
        self.name_row = None
        # "read_cost_row": "实际罚款金额",
        self.read_cost_row = None
        # "read_cost_month_row": "扣款月份"
        self.read_cost_month_row = None


def get_spec_file_conf():
    conf_dict = {}
    conf_obj = json_data['data']["spec"]
    for conf_one in conf_obj:
        conf_data = SpecFileConf()
        conf_data.file_name = conf_one["file_name"]
        conf_data.file_month = conf_one["file_month"]
        conf_data.read_sheet_name = conf_one["read_sheet_name"]
        conf_data.region_row = conf_one["region_row"]
        conf_data.read_cost_row = conf_one["read_cost_row"]
        conf_data.read_cost_month_row = conf_one["read_cost_month_row"]
        conf_data.sum_row_text = conf_one["sum_row_text"]
        conf_data.name_row = conf_one["name_row"]
        conf_dict[conf_one["file_name"]] = conf_data
    return conf_dict


class BoundConf:
    def __init__(self):
        self.file_name = None
        self.bound_title = None
        self.read_sheet_name = None
        self.start_column = None
        self.end_column = None
        self.true_name_column = None
        self.company_column = None
        self.work_column = None


def get_bound_file_conf():
    conf_obj = json_data['data']["bound"]
    conf_data = BoundConf()
    conf_data.file_name = conf_obj["file_name"]
    conf_data.bound_title = conf_obj["bound_title"]
    conf_data.read_sheet_name = conf_obj["read_sheet_name"]
    conf_data.start_column = conf_obj["start_column"]
    conf_data.end_column = conf_obj["end_column"]
    conf_data.true_name_column = conf_obj["true_name_column"]
    conf_data.company_column = conf_obj["company_column"]
    conf_data.work_column = conf_obj["work_column"]
    return conf_data


def get_bound_table_title_area():
    conf = get_bound_file_conf()
    return conf.start_column + '2:' + conf.end_column + '3'


#       "file_month": "2024年4月份",
#       "file_name": "1+11月份  2024-04  C类推广员百分之7推广奖励.xlsx",
#      "read_sheet_name": "2024年04月C类推广员7%推广奖励 (2)",
#     "read_bound_row": "业务员应得金额合计",
#     "read_region_row": "大区",
#     "read_part_row": "分区",
#    "read_name_row": "业务员姓名",
#   "read_work_row": "岗位名称",
#   "read_start_row": 2
class C7Conf:
    def __init__(self):
        self.file_month = None
        self.file_name = None
        self.read_sheet_name = None
        self.read_bound_row = None
        self.read_region_row = None
        self.read_part_row = None
        self.read_name_row = None
        self.read_work_row = None
        self.read_start_row = None
        self.data_start_row = 3


def get_c7_file_conf():
    conf_dict = {}
    conf_objs = json_data['data']["c"]

    for conf_obj in conf_objs:
        conf_data = C7Conf()
        conf_data.file_month = conf_obj["file_month"]
        conf_data.file_name = conf_obj["file_name"]
        conf_data.read_sheet_name = conf_obj["read_sheet_name"]
        conf_data.read_bound_row = conf_obj["read_bound_row"]
        conf_data.read_region_row = conf_obj["read_region_row"]
        conf_data.read_part_row = conf_obj["read_part_row"]
        conf_data.read_name_row = conf_obj["read_name_row"]
        conf_data.read_work_row = conf_obj["read_work_row"]
        conf_data.read_start_row = conf_obj["read_start_row"]
        conf_dict[conf_obj["file_name"]] = conf_data
    return conf_dict

