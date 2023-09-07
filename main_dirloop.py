# 查找指定目录下所有Excel表格（精确匹配）
# 使用pyinstaller打包exe可执行文件，命令: pyinstaller -F xxx.py
import openpyxl
import xlrd
from tkinter import Tk, filedialog
from os import listdir, path

print(
    "    ______               _________           __         \n   / ____/  __________  / / ____(_)___  ____/ /__  _____\n  / __/ | |/_/ ___/ _ \/ / /_  / / __ \/ __  / _ \/ ___/\n / /____>  </ /__/  __/ / __/ / / / / / /_/ /  __/ /    \n/_____/_/|_|\___/\___/_/_/   /_/_/ /_/\__,_/\___/_/     \n                                                        ")
print("**V2.1 dirloop ver -- 作者：高世岷**")


def run():
    excel_name = []
    sheet_name = []
    cell_num = []
    lists = []

    def search_path(v, p):
        ls = listdir(p)
        for k in ls:
            if k.endswith("xlsx") and not k.startswith("~$"):
                print("+正在查找：" + k)
                excel = p + r'\%s' % k
                data = openpyxl.load_workbook(excel, read_only=True, data_only=True)
                for sheet in data.worksheets:
                    for row in sheet.rows:
                        for cell in row:
                            if str(cell.value) == v:
                                excel_name.append(k)
                                sheet_name.append(sheet.title)
                                cell_num.append(cell.coordinate)
            elif k.endswith("xls") and not k.startswith("~$"):
                print("+正在查找：" + k)
                excel = p + r'\%s' % k
                data = xlrd.open_workbook_xls(excel)
                for sheet in data.sheets():
                    for rowidx in range(sheet.nrows):
                        row = sheet.row(rowidx)
                        # print(row)
                        for colidx, cell in enumerate(row):
                            if str(cell.value) == v:
                                excel_name.append(k)
                                sheet_name.append(sheet.name)
                                cell_num.append(xlrd.cellname(rowidx, colidx))
            else:
                if path.isdir(p + r'\%s' % k):
                    print("*进入目录：" + k)
                    search_path(v, p + r'\%s' % k)
                else:
                    print("-跳过无效文件：" + k)

    # 获取待查找单元格值
    value = str(input('*' * 40 + "\n输入查找值（精确匹配）："))
    # 获取目录路径
    root = Tk()
    root.attributes("-topmost", 1)
    root.withdraw()
    directory = filedialog.askdirectory()
    print("查找目录：" + directory + '\n' + '-' * 40)
    search_path(value, directory)

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
