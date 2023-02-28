# 查找指定目录下所有Excel表格（精确匹配）
import openpyxl
import xlrd
from tkinter import Tk, filedialog
from os import listdir

print(
    "    ______               _________           __         \n   / ____/  __________  / / ____(_)___  ____/ /__  _____\n  / __/ | |/_/ ___/ _ \/ / /_  / / __ \/ __  / _ \/ ___/\n / /____>  </ /__/  __/ / __/ / / / / / /_/ /  __/ /    \n/_____/_/|_|\___/\___/_/_/   /_/_/ /_/\__,_/\___/_/     \n                                                        ")
print("**V1.1 - Written by Shimin Gao**")


def run():
    excel_name = []
    sheet_name = []
    cell_num = []
    lists = []
    # 获取待查找单元格值
    search_value = str(input('*' * 40 + "\n输入查找值（精确匹配）："))
    # 获取目录路径
    root = Tk()
    root.attributes("-topmost", 1)
    root.withdraw()
    path = filedialog.askdirectory()
    ls = listdir(path)
    print("查找目录：" + path + '\n' + '-' * 40)

    for k in ls:
        if k.endswith("xlsx") and not k.startswith("~$"):
            print("正在查找：" + k)
            excel = path + r'\%s' % k
            data = openpyxl.load_workbook(excel, read_only=True, data_only=True)
            for sheet in data.worksheets:
                for row in sheet.rows:
                    for cell in row:
                        if cell.value == search_value:
                            excel_name.append(k)
                            sheet_name.append(sheet.title)
                            cell_num.append(cell.coordinate)
        elif k.endswith("xls") and not k.startswith("~$"):
            print("正在查找：" + k)
            excel = path + r'\%s' % k
            data = xlrd.open_workbook_xls(excel)
            for sheet in data.sheets():
                for rowidx in range(sheet.nrows):
                    row = sheet.row(rowidx)
                    # print(row)
                    for colidx, cell in enumerate(row):
                        if cell.value == search_value:
                            excel_name.append(k)
                            sheet_name.append(sheet.name)
                            cell_num.append(xlrd.cellname(rowidx, colidx))
        else:
            print("跳过无效文件/目录：" + k)
    for k in range(len(sheet_name)):
        lists.append([excel_name[k], sheet_name[k], cell_num[k]])

    print('-' * 40 + '\n' + "【查询结果】")
    for k in lists:
        print(str(k))


while True:
    try:
        run()
    except IOError:
        print("ERROR: 路径/读取错误")
    except:
        print("ERROR: 未知错误")
