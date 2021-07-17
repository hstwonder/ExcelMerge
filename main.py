# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

import openpyxl
from openpyxl import Workbook
import mysql.connector
import getopt
import sys


def change_format(dt):
    if dt != '':
        return dt.strftime("%Y/%m/%d")


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press ⌘F8 to toggle the breakpoint.


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


def format_data(lst):
    lst[0] = change_format(lst[0])
    lst[1] = change_format(lst[1])
    UIN = lst[2]
    if isinstance(UIN, str) is False:
        lst[2] = str(int(UIN))
    else:
        UIN = UIN.strip()
        if len(UIN) > 16:
            return
        lst[2] = UIN
    lst[9] = str(lst[9])
    lst[15] = str(lst[15])
    return lst


def check_format(str):
    if type(str) != type('a'):
        return False
    str = str.strip()
    str = str.lower()
    if len(str) == 0 or len(str) > 5 or len(str) < 4:
        return False

    fis = str[0]
    sen = str[1]
    thd = str[2]
    if fis != 'q' or \
            isinstance(int(sen), int) is not True or \
            thd != 'w':
        return False
    return True


def load_excel_file(filename):
    mapdata = {}
    book = openpyxl.load_workbook(filename)
    sheet = book.active
    sheet.guess_types = True
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=19):
        lst_cell = []
        for cell in row:
            # print(cell.value, end=" ")
            if cell.value is None:
                lst_cell.append("")
            else:
                lst_cell.append(cell.value)

        # lst_cell[0] = change_format(lst_cell[0])
        # lst_cell[1] = change_format(lst_cell[1])
        # UIN = lst_cell[2]
        # if isinstance(UIN, str) is False:
        #     lst_cell[2] = str(int(UIN))
        # else:
        #     UIN = UIN.strip()
        #     if len(UIN) > 16:
        #         continue
        #     lst_cell[2] = UIN
        # lst_cell[9] = str(lst_cell[10])
        # lst_cell[15] = str(lst_cell[15])
        # # print(lst_cell[2], lst_cell[3], lst_cell[4])

        print(lst_cell)
        if format_data(lst_cell) is None:
            continue

        if check_format(lst_cell[8]) is False:
            lst_cell.append(False)
        else:
            lst_cell.append(True)

        key = str(lst_cell[2]) + lst_cell[3]

        mapdata[key] = lst_cell
    return mapdata


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
        else:
            name = '有问题'
        sheet = wb.create_sheet(title=name, index=i)
        write_title(sheet)
        nrow = 1
        for k, v in lstsheet[i].items():
            nrow = nrow + 1
            for j in range(1, len(v)):
                sheet.cell(row=nrow, column=j).value = v[j - 1]

    wb.save(file)


def cmp_value(s_lst, c_lst):
    return True


def merge(s_file, i_file, o_file):
    src_map = load_excel_file(s_file)
    cmp_map = load_excel_file(i_file)

    sheet1_map = {}  # all
    sheet2_map = {}  # update
    sheet3_map = {}  # add
    sheet4_map = {}  # error

    for sk, sv in src_map.items():
        cv = cmp_map.get(sk)
        if cv is None:
            flag = sv[len(sv) - 1]
            sv.pop()
            if flag is False:
                sheet4_map[sk] = sv
            sheet1_map[sk] = sv  # 未更新
        else:
            cmp_map.pop(sk)  # 找到值
            if cmp_value(sv, cv) is True:
                sheet1_map[sk] = sv
            else:
                sheet1_map[sk] = cv
                flag = cv[len(cv) - 1]
                cv.pop()
                if flag is True:
                    sheet2_map[sk] = cv
                else:
                    sheet4_map[sk] = cv

    for ck, cv in cmp_map.items():  # 新增
        print(cv)
        UIN = cv[2]
        if isinstance(UIN, str) is False:
            cv[2] = str(int(UIN))
        else:
            UIN = UIN.strip()
            if len(UIN) > 16:
                continue
            cv[2] = UIN

        expMoney = cv[9].strip()
        if len(str(expMoney)) < 4:
            continue
        if expMoney[-1] == '+':
            expMoney = expMoney[:-1]

        if int(float(expMoney)) < 3000:
            continue

        sheet1_map[ck] = cv
        sheet3_map[ck] = cv
        flag = cv[len(cv) - 1]
        cv.pop()
        if flag is False:
            sheet4_map[ck] = cv

    sheet1_map.clear()
    lst_sheet = [sheet1_map, sheet2_map, sheet3_map, sheet4_map]
    write_excel_file(o_file, lst_sheet)


# Press the green button in the gutter to run the script.
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
