import itchat
import xlwt
import xlrd
import glob
from xlutils.copy import copy
import random
import time

max_delayed = 10

wb = None
sheet = None
rows = 0

id_index = 0
sex_index = 1
name_index = 2
remark_index = 3
special_index = 4
group_index = 5

boy_blessing = 0
girl_blessing = 1


# 设置表格样式
def set_style(name, height, bold=False):
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = name
    font.bold = bold
    font.color_index = 4
    font.height = height
    style.font = font
    return style


def find_file(filename):
    files = glob.glob('*')
    for file in files:
        if file == filename:
            return True

    return False


def fetch_friend_list():

    global  wb
    global  sheet
    global  rows
    global  name_index
    global  remark_index
    global  special_index
    global  group_index

    filename = "好友列表.xls"
    total_list = []
    has_file = find_file(filename)

    if has_file:
        twb = xlrd.open_workbook(filename)
        t_sheet = twb.sheet_by_name('好友')
        rows = t_sheet.nrows
        for i in range(0, rows):
            row = t_sheet.row_values(i)
            total_list.append(row)

        wb = copy(twb)
        sheet = wb.get_sheet(0)
    else:
        wb = xlwt.Workbook()
        sheet = wb.add_sheet("好友", cell_overwrite_ok=True)
        row0 = ["id","性别","用户名", "备注名", "特殊名", "分组"]
        # 写第一行
        for i in range(0, len(row0)):
            sheet.write(0, i, row0[i], set_style('Times New Roman', 250))

        sheet2 = wb.add_sheet("祝福语", cell_overwrite_ok=True)
        sheet2_row0 = ["男生祝福语", "女生祝福语"]
        sheet2_row1 = ["祝name新的一年里，身体健康，工作顺利。", "祝name新的一年里，天天开心，越来越美。"]
        for i in range(0, len(sheet2_row0)):
            sheet2.write(0, i, sheet2_row0[i], set_style('Times New Roman', 250))
            sheet2.write(1, i, sheet2_row1[i], set_style('Times New Roman', 250))

        rows = 1

    friends_list = get_friends()

    old_names_list = []
    for friend in total_list:
        old_names_list.append([friend[name_index], friend[remark_index]])

    i = rows

    for friend in friends_list:
        row = [friend["UserName"],friend["Sex"],friend["NickName"], friend["RemarkName"],"", ""]
        name_row = [friend["NickName"], friend["RemarkName"]]
        if old_names_list.count(name_row) > 0:
            index = old_names_list.index(name_row)
            old_row = total_list[index]
            new_value = [friend["UserName"], friend["Sex"], friend["NickName"], friend["RemarkName"], old_row[special_index], old_row[group_index]]
            total_list[index] = new_value
            for j in range(0, len(new_value)):
                sheet.write(index, j, new_value[j], set_style('Times New Roman', 250))
        else:
            total_list.append(row)
            for j in range(0, len(row)):
                sheet.write(i, j, row[j], set_style('Times New Roman', 250))
            i += 1

    wb.save(filename)
    return  total_list


def get_friends():
    # 获取微信好友
    itchat.auto_login()
    friends_list = itchat.get_friends(update=True)[1:]
    return friends_list


def get_blessing():
    filename = "好友列表.xls"
    wb = xlrd.open_workbook(filename)
    t_sheet = wb.sheet_by_name('祝福语')
    cols = t_sheet.ncols
    blessing_dic = {}
    for i in range(0, cols):
        col = t_sheet.col_values(i)
        if len(col) >= 2:
            list = []
            for j in range(1, len(col)):
                list.append(col[j])

            blessing_dic[col[0]] = list

    return blessing_dic

def send_or_print(friend_list):
    result = input("确定群发消息？ y发送/n打印:")
    while True:
        if result in ("y", "Y"):
            send_msg(friend_list, True)
            break
        if result in ("n", "N"):
            send_msg(friend_list, False)
            break


def send_msg(list, is_send):

    global  name_index
    global  remark_index
    global  special_index
    global max_delayed

    blessing_dic = get_blessing()

    for i in range(1, len(list)):
        friend = list[i]
        if friend[special_index]:
            username = friend[special_index]
        elif friend[remark_index]:
            username = friend[remark_index]
        else:
            username = friend[name_index]


        if len(friend[group_index]) > 1:
            if friend[group_index] in blessing_dic.keys():
                blessings = blessing_dic[friend[group_index]]
        else:
            if friend[sex_index] == 2:
                blessings = blessing_dic["女生祝福语"]
            else:
                blessings = blessing_dic["男生祝福语"]

        blessing = random.choice(blessings)
        msg = blessing.replace("name", username)

        if is_send:
            time.sleep(random.randint(2, max_delayed))
            itchat.send_msg(msg, toUserName=friend[id_index])
            print("发送: %s" % msg)
        else:
            print("打印: %s" % msg)






if __name__ == '__main__':

    friend_list = fetch_friend_list()

    result = input("已经修好好好友列表？ y修改好/n没修改好:")
    while True:
        if result in ("y", "Y"):
            send_or_print(friend_list)
            break
        if result in ("n", "N"):
            result = input("是否已经修好好好友列表 y/n:")










