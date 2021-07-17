# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

import openpyxl
from openpyxl import Workbook
import mysql.connector
import getopt
import sys
import re  # python的正则表达式模块
import copy

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
def update_db(data):
    # mydb = mysql.connector.connect(
    #     host="172.27.16.6",  # 数据库主机地址
    #     user="root",  # 数据库用户名
    #     passwd="Yxkj12345678!!",  # 数据库密码
    #     database="intent"
    # )

    mydb = mysql.connector.connect(
        host="localhost",  # 数据库主机地址
        user="root",  # 数据库用户名
        passwd="123456789",  # 数据库密码
        database="intent"
    )
    mycursor = mydb.cursor()

    sql = "INSERT INTO KA (regist, lastreach, UIN, company, leads, " \
          "team, sale, winchance, exporder, expmoney" \
          "produce, stuckpoint, customerbusiness, ascription, order, " \
          "money, first, remark) " \
          "VALUES (%s, %s, %s, %s, %s, " \
          "%s, %s, %d, %s, %d," \
          "%s, %s, %s, %d, %s," \
          "%d, %d, %s)"

    val = tuple(data)
    print(val)
    mycursor.execute(sql, val)
    mydb.commit()  # 数据表内容有更新，必须使用到该语句


def change_date_format(dt):
    if dt != '':
        return dt.strftime("%Y/%m/%d")


def format_data(lst_value):
    lst_value[0] = change_date_format(lst_value[0])
    lst_value[1] = change_date_format(lst_value[1])
    UIN = lst_value[2]
    if isinstance(UIN, str) is False:
        UIN = str(int(UIN))
    else:
        UIN = UIN.replace('_x000D_', '')
        UIN = UIN.strip()


    lst_value[2] = UIN

    Leads = str(lst_value[4]).strip()
    lst_value[4] = Leads

    pattern = r"[\s]*[+-]?[\d]+"
    match = re.match(pattern, str(lst_value[7]).strip())
    if match:
        lst_value[7] = float(match.group(0))/100

    match = re.match(pattern, str(lst_value[9]).strip())
    if match:
        lst_value[9] = int(match.group(0))
    else:
        lst_value[9] = 9999
    # (expMoney) = [t(s) for t, s in
    #                   zip((int, int), re.search('^(\d+).(\d+)$', str(lst_value[9]).strip()).groups())]
    # lst_value[9] = float(expMoney[0])
    # if len(expMoney) > 1 and expMoney[-1] == '+':
    #     lst_value[9] = float(expMoney[:-1])
    # lst_value[9] = float(lst[expMoney][0])

    lst_value[15] = str(lst_value[15])
    return lst


def check_legal(lst_value):
    try:
        UIN = lst_value[2]
        if len(UIN) > 16 or len(UIN) < 4:
            return False

        # Company = lst_value[3]
        # if len(Company) < 4:
        #     return False

        ExpDate = lst_value[8].strip()
        if isinstance(ExpDate, str) is True:
                (fq, q, fw, w) = [t(s) for t, s in
                                  zip((str, int, str, int), re.search('^(\w)(\d)(\w)(\d+)$', ExpDate).groups())]
    except:
        return False

    return True


def load_excel_file(filename, sheet_name=None):
    mapData = {}
    book = openpyxl.load_workbook(filename)
    if sheet_name is None:
        sheet = book.active
    else:
        sheet = book.get_sheet_by_name(sheet_name)

    sheet.guess_types = True
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=19):
        lst_cell = []
        for cell in row:
            # print(cell.value, end=" ")
            if cell.value is None:
                lst_cell.append("")
            else:
                lst_cell.append(cell.value)

        format_data(lst_cell)
        if lst_cell[9] < 3000.0:
            continue
        if lst_cell[0] is None:
            continue
        lst_UIN = lst_cell[2].split()
        # if len(lst_UIN) > 1:
        #     print(1, lst_UIN)

        for item in lst_UIN:
            # if len(lst_UIN) > 1:
            #     print(item, lst_UIN)
            lst_cell[2] = item
            #print(lst_value)

            if check_legal(lst_cell) is False:
                lst_cell.append(False)
            else:
                lst_cell.append(True)

            # 登记日期 + UIN + 组
            key = lst_cell[0] + str(lst_cell[2]) + lst_cell[5]
            if mapData.get(key) is None:
                value = copy.deepcopy(lst_cell)
                mapData[key] = value
                if len(lst_UIN) > 1:
                    print(value, 'ld', len(mapData))
            lst_cell.pop()

    return mapData


def write_title(sheet):
    title = ["登记时间", "最近触达时间", "UIN", "客户名称", "Leads来源", "主管", "销售", "赢单率", "预计下单时间",
             "商机金额（预估）", "一级产品", "卡点", "客户业务", "状态（电销/SMB）", "实际下单时间", "实际金额", "是否首次采购", "备注"]
    for i in range(1, len(title) + 1):
        sheet.cell(row=1, column=i).value = title[i - 1]


def write_excel_file(file, lstsheet):
    wb = Workbook()
    for i in range(0, len(lstsheet)):
        if i == 0:
            name = '全部'
        elif i == 1:
            name = '更新'
        elif i == 2:
            name = '新增'
        elif i == 3:
            name = '有问题'
        elif i == 4:
            name = '5W'
        sheet = wb.create_sheet(title=name, index=i)
        write_title(sheet)
        nrow = 1
        for k, v in lstsheet[i].items():
            # if v[3] == '广州金嗓音文化发展有限公司':
            #     print(v, 'w')
            nrow = nrow + 1
            for j in range(1, len(v)):
                sheet.cell(row=nrow, column=j).value = v[j - 1]

    wb.save(file)


def cmp_value(s_lst, c_lst):
    return True


def merge(s_file, i_file, o_file):
    src_map = {}#load_excel_file(s_file, '明细')
    src_map5W = load_excel_file(s_file, '5W+')
    cmp_map = load_excel_file(i_file)
    sheet1_map = {}  # all
    sheet2_map = {}  # update
    sheet3_map = {}  # add
    sheet4_map = {}  # error
    sheet5_map = {}  # 5W+

    for ck, cv in cmp_map.items():
        sv = src_map.get(ck)
        if cv[3] == '广州金嗓音文化发展有限公司':
            print(cv, 'm1')
        if sv is None:  # 源文件中没有相同的值
            flag = cv[-1]
            cv.pop()  # 删除标记位
            if cv[9] >= 50000.0:
                sv = src_map5W.get(ck)
                if sv is None:
                    sheet5_map[ck] = cv
            else: # 不足5万
                if flag is False:  # 新文件值有错误
                    sheet4_map[ck] = cv
                else:
                    if cv[3] == '广州金嗓音文化发展有限公司':
                         print(cv, 'm2')
                    sheet3_map[ck] = cv

        else:  # 源文件中有相同的值
            if cmp_value(sv, cv) is False:  # 比较文件与源文件内容不同
                sheet2_map[ck] = cv

    # for ck, cv in cmp_map.items():  # 新增
    #     print(cv)
    #     UIN = cv[2]
    #     if isinstance(UIN, str) is False:
    #         cv[2] = str(int(UIN))
    #     else:
    #         UIN = UIN.strip()
    #         if len(UIN) > 16:
    #             continue
    #         cv[2] = UIN

        # expMoney = cv[9].strip()
        # if len(str(expMoney)) < 4:
        #     continue
        # if expMoney[-1] == '+':
        #     expMoney = expMoney[:-1]

        # if int(float(expMoney)) < 3000:
        #     continue

        # sheet1_map[ck] = cv
        # sheet3_map[ck] = cv
        # flag = cv[len(cv) - 1]
        # cv.pop()
        # if flag is False:
        #     sheet4_map[ck] = cv

    sheet1_map.clear()
    lst_sheet = [sheet1_map, sheet2_map, sheet3_map, sheet4_map, sheet5_map]
    write_excel_file(o_file, lst_sheet)


# Press the green button in the gutter to run the script.
def lst(param):
    pass


if __name__ == '__main__':
    opts, args = getopt.getopt(sys.argv[1:], "hs:i:o:", ["help", "src=", "input=", "output="])

    for opts, arg in opts:
        print(opts)
        if opts == "-h" or opts == "--help":
            print("我只是一个说明文档")
        elif opts == "-s" or opts == "--src":
            src_file = arg
        elif opts == "-i" or opts == "--input":
            input_file = arg
        elif opts == "-o" or opts == "--output":
            output_file = arg

    # print_hi('PyCharm')
    if len(sys.argv) < 4:
        exit(0)

    merge(src_file, input_file, output_file)
