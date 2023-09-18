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
# print(sheet_test01[1])
# for row in sheet_test01:  # 印出worksheet的所有row（橫排） # column（豎排）要再另外查
#     print(row)
# print(type(sheet_test01.cell))
# print(type(sheet_test01[1]))

# read
# A1 = sheet_test01.cell('A1')
# print(A1)
# print(A1.value)

#讀取成 df
# df = pd.DataFrame(wks.get_all_records())
#讀取 df 也可以這樣寫
# sheet_test01.get_as_df()

sheet_data = sheet_test01.get_all_records()
pprint.pprint(sheet_data) # 輸出整頁 google sheet
# pprint.pprint(sheet_test01.get_all_records()[1]) # 輸出整頁 google sheet

# sheet_data = sheet_test01.get_all_records()
# for row in sheet_data:
#     pprint.pprint(row['時間戳記']) # 輸出整頁 google sheet

print(sheet_data[0]['帳務時間'])

myExcel = MyExcel()
column_A = myExcel.sheet.range('A1').expand('down').value  # 把右邊裝到左邊（把右邊寫到左邊），把右邊一連串東東回傳的結果放到名為 column_A 的盒子

print(column_A)
print(type(column_A))

if column_A is None:  # 如果 column_A 沒有內容
    length_A = 0  # 把右邊的 0 放到名為 length_A 的盒子
elif not isinstance(column_A, list):  # 如果 column_A 這個盒子的型態是 list
    length_A = 1  # 把右邊的 1 放到名為 length_A 的盒子
else:
    length_A = len(column_A)  # 把右邊計算 column_A 長度（len）的結果放到名為 length_A 的盒子
startIndex_A = length_A+1  # 把右邊計算 length_A+1 的結果放到名為 startIndex_A 的盒子

print(list(sheet_data[0].values())) # [0] 1. 代表取了 index（索引）為零的東西 2. （包含標題列在內的）第一列的內容
googleHeader = ["帳務時間","類型","收入類別","消費內容","開銷內容","收入明細","支出明細","收入金額","入帳戶名","支出金額","付款方式","收入備註","支出備註"] 
# 將欲填入 excel 裡面的 google 表單的 key 順序，命名為 googleHeader
# googleHeader 就是 google sheet 的標題列
# print([sheet_data[0][key] for key in googleHeader]) # list comprehension（列表生成式）

inToExcel = []
for key in googleHeader:
    # 按照 googleHeader 的順序去跑 key
    print(key)
    print(sheet_data[0][key])
    # 印出 key 對應的 value
    if not sheet_data[0][key]:
        # 如果 key 對應的 value 是非空字串，再填入 value
        # 使用時機，如 google sheet 的「收入類別」、「消費內容」、「開銷內容」三欄的內容，都要視時機填入 Excel 的「類別」
        inToExcel.append(sheet_data[0][key])

# 將 value 填入 excel 對應的行
# myExcel.sheet.range('A'+str(startIndex_A)).value = sheet_data[0]  # 把右邊的資料（sheet_data[0]）填入 A 行的 startIndex_A（length_A+1）格子
# myExcel.sheet.range('A'+str(startIndex_A)).options(transpose = True).value = sheet_data[0]['日期'],   # 把右邊的資料（sheet_data[0]）經過轉置（options(transpose = True)）填入 A 行的 startIndex_A（length_A+1）格子
# myExcel.sheet.range('A'+str(startIndex_A)).options(transpose = False).value = list(sheet_data[0].values())  # 取出 sheet_data[0] 這個 dict 當中的 values，但取出的型態（type）會是 dict_values 需要轉換成可以讓 excel 吃到的 list
myExcel.sheet.range('A'+str(startIndex_A)).options(transpose = False).value = inToExcel



 # 之後進度（action item）：將 value 填入 excel 對應的行→用 if else 迴圈，篩選、插入 G →測試一次填入好幾行→設定程式執行的時機→從 google sheet 抓的範圍


