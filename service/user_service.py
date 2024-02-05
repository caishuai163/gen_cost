import openpyxl

from dao import basic_user


def reload_file():
    basic_user.remove_all()
    xlsx_file = openpyxl.load_workbook(filename="./source/user_tb.xlsx",
                                       data_only=True)  # 打开文件
    default_sheet = xlsx_file['sheet1']  # 读取名为Sheet1的表
    row_max = default_sheet.max_row  # 获取最大行
    title_size = 1
    for x in range(row_max - title_size):
        user_name = default_sheet.cell(row=x + title_size + 1, column=1).value
        region = default_sheet.cell(row=x + title_size + 1, column=2).value
        region_manager = default_sheet.cell(row=x + title_size + 1, column=3).value
        part = default_sheet.cell(row=x + title_size + 1, column=4).value
        company = default_sheet.cell(row=x + title_size + 1, column=5).value
        work = default_sheet.cell(row=x + title_size + 1, column=6).value
        dealer_name = default_sheet.cell(row=x + title_size + 1, column=7).value
        platform = default_sheet.cell(row=x + title_size + 1, column=8).value
        need_generated = default_sheet.cell(row=x + title_size + 1, column=9).value

        if user_name is None or user_name == "":
            continue
        basic_user.insert(user_name, region, region_manager, part, company, work, dealer_name, platform, need_generated)

