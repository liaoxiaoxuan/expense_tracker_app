import pygsheets
import pprint # 使用 pprint 套件，印出較容易閱讀的格式

auth_file = "../static/expense-tracker-app-392308-a0f2bb4f12a1.json"
googleCloud = pygsheets.authorize(service_file = auth_file)

# setting sheet
sheet_url = "https://docs.google.com/spreadsheets/d/1yHXGgmFkMhYJYzsunphNErEkDntMP5MiWYyxJx-fqUs/edit#gid=881554370" 
spreadSheet = googleCloud.open_by_url(sheet_url)

#選取by名稱
sheet_test01 = spreadSheet.worksheet_by_title("表單回應 1")
# print(sheet_test01[1])
# for row in sheet_test01:  # 印出worksheet的所有row（橫排） # column（豎排）要再另外查
#     print(row)
# print(type(sheet_test01.cell))
# print(type(sheet_test01[1]))

# read
A1 = sheet_test01.cell('A1')
# print(A1)
# print(A1.value)

#讀取成 df
# df = pd.DataFrame(wks.get_all_records())
#讀取 df 也可以這樣寫
# sheet_test01.get_as_df()

# sheet_test01.get_all_records()
# pprint.pprint(sheet_test01.get_all_records()) # 輸出整頁 google sheet
# pprint.pprint(sheet_test01.get_all_records()[1]) # 輸出整頁 google sheet
sheet_data = sheet_test01.get_all_records()
for row in sheet_data:
    pprint.pprint(row['時間戳記']) # 輸出整頁 google sheet

