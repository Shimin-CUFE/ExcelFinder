import openpyxl
import xlrd
from os import listdir, path
from tkinter.constants import END


class FindMethod:
    def __init__(self):
        pass

    def search_path(self, param_dict):
        p = param_dict.get("p")
        v = param_dict.get("v")
        tx = param_dict.get("text")
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
                                print("mark")  # 保存
                                tx.insert(END, str(k) + str(sheet.title) + str(cell.coordinate))
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
                                print("mark")  # 保存
            else:
                if path.isdir(p + r'\%s' % k):
                    print("进入目录：" + k)
                    new_dict = {
                        'v': v,
                        'p': p + r'\%s' % k,
                        'text': tx
                    }
                    self.search_path(new_dict)
                else:
                    print("-跳过无效文件：" + k)
