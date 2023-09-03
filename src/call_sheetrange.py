# sheet.range('A1').value=[[1, 2], [3, 4]]
# sheet.range('A4').value='Hello World'

import xlwings

# 開啟→excel（工作簿）→sheet（工作表）→填入資料
# app = xlwings.App(visible=True, add_book=False)
# workbook = app.books.open(r'./local/expense_tracker_app_text.xlsx')   # 開啟excel（工作簿）
# sheet = workbook.sheets['工作表1']  # 開啟sheet（工作表）

# sheet.range('A1').value=[[1, 2], [3, 4]]   # 在指定儲存格內填入資料



# 封裝（打包）成class
class MyExcel():
    def __init__(self):
        self.app = xlwings.App(visible=True, add_book=False)
        self.workbook = self.app.books.open(r'../local/expense_tracker_app_text.xlsx')   # 開啟excel（工作簿）
        self.sheet = self.workbook.sheets['工作表1']  # 開啟sheet（工作表）