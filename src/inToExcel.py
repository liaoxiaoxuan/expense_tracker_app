# x=input("請輸入數字: ") # 取得字串形式的使用者輸入
# x=int(x) #將字串轉型態轉換成數字型態
# if x>200:
#     print("大於200")
# elif x>100:
#     print("大於100, 小於等於200")
# else:
#     print("小於等於100")

import pygsheets
import pprint # 使用 pprint 套件，印出較容易閱讀的格式
from call_sheetrange import MyExcel

auth_file = "../static/expense-tracker-app-392308-a0f2bb4f12a1.json"
googleCloud = pygsheets.authorize(service_file = auth_file)

# setting sheet
sheet_url = "https://docs.google.com/spreadsheets/d/1yHXGgmFkMhYJYzsunphNErEkDntMP5MiWYyxJx-fqUs/edit#gid=881554370" 
spreadSheet = googleCloud.open_by_url(sheet_url)

#選取by名稱
sheet_test01 = spreadSheet.worksheet_by_title("表單回應 1")
sheet_data = sheet_test01.get_all_records()


import xlwings

# 開啟→excel（工作簿）→sheet（工作表）→填入資料
app = xlwings.App(visible=True, add_book=False)
workbook = app.books.open(r'../local/expense_tracker_app_text.xlsx')   # 開啟excel（工作簿）
sheet = workbook.sheets['g.s.C']  # 開啟sheet（工作表）

# # 資料獲取範圍
# last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row

# print(range(len(sheet_data)))

# 填入迴圈

for row in range(len(sheet_data)):
    # for row_data in sheet_data:
    # print(row,sheet_data[row]["消費內容"])
    if sheet_data[row]["收入類別"] != '':
        sheet.range('G' + str(row + 2)).value = sheet_data[row]["收入類別"]
    elif sheet_data[row]["消費內容"] != '':
        sheet.range('G' + str(row + 2)).value = sheet_data[row]["消費內容"]
    elif sheet_data[row]["開銷內容"] != '':
        sheet.range('G' + str(row + 2)).value = sheet_data[row]["開銷內容"]

# is not None 要改成「!= ''」（不等於空字串），因為 google sheet 裡面空格是放空字串不是 None
