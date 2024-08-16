
def get_column(sheet, row, row_val):
    column = 0
    for i in range(100):
        val = sheet.cell(row=row, column=i + 1).value
        if val is not None and str(val).replace('\n', '').replace('\r', '').replace('\t', '') \
                .replace(' ', '') == row_val:
            column = i + 1
            break
    if column == 0:
        raise Exception("加载当前文件未找到" + row_val + "列")
    return column
