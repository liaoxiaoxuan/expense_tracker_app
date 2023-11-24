import pprint # 使用 pprint 套件，印出較容易閱讀的格式
import time

import pygsheets
import xlwings as xw # 使用 xw 執行 xlwings

from .call_sheetrange import MyExcel
from src import MyLogger
from src import __author__, __version__

# 插入log
def WriteLog(Logger, strMsg):
    try:
        Logger.info(strMsg)
        strMsg = time.strftime("%Y/%m/%d %H:%M:%S ") + str(strMsg)
    except AttributeError as e:
        # need an error log
        pass

def expense_tracker():
    # 建立Logger
    Logger = MyLogger.SetLogger(__name__)
    strMsg = "start expense tracker"
    WriteLog(Logger, strMsg)

    auth_file = "./local/expense-tracker-app3-b8d7f50165a5.json"
    googleCloud = pygsheets.authorize(service_file = auth_file)
    
    # setting sheet
    sheet_url = "https://docs.google.com/spreadsheets/d/1yHXGgmFkMhYJYzsunphNErEkDntMP5MiWYyxJx-fqUs/edit#gid=881554370" 
    spreadSheet = googleCloud.open_by_url(sheet_url)
    
    #選取by名稱
    sheet_test01 = spreadSheet.worksheet_by_title("記帳暫存")
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
    
    # 獲取 google sheet data
    sheet_data = sheet_test01.get_all_records()
    # pprint.pprint(sheet_data) # 輸出整頁 google sheet
    # pprint.pprint(sheet_test01.get_all_records()[1]) # 輸出整頁 google sheet
    
    # sheet_data = sheet_test01.get_all_records()
    # for row in sheet_data:
    #     pprint.pprint(row['時間戳記']) # 輸出整頁 google sheet
    
    # print(sheet_data[0]['帳務時間'])
    
    myExcel = MyExcel()
    column_A = myExcel.sheet.range('A1').expand('down').value  # 把右邊裝到左邊（把右邊寫到左邊），把右邊一連串東東回傳的結果放到名為 column_A 的盒子
    sheet = myExcel.sheet  # 把右邊裝到左邊（把右邊寫到左邊），把右邊一連串東東回傳的結果放到名為 column_A 的盒子
    
    # print(column_A)
    # print(type(column_A))
    
    if column_A is None:  # 如果 column_A 沒有內容
        length_A = 0  # 把右邊的 0 放到名為 length_A 的盒子
    elif not isinstance(column_A, list):  # 如果 column_A 這個盒子的型態是 list
        length_A = 1  # 把右邊的 1 放到名為 length_A 的盒子
    else:
        length_A = len(column_A)  # 把右邊計算 column_A 長度（len）的結果放到名為 length_A 的盒子
    startIndex_A = length_A+1  # 把右邊計算 length_A+1 的結果放到名為 startIndex_A 的盒子
    
    # print(list(sheet_data[0].values())) # [0] 1. 代表取了 index（索引）為零的東西 2. （包含標題列在內的）第一列的內容
    # excelHeader = ["帳務時間","類型","收入類別","消費內容","開銷內容","收入明細","支出明細","收入金額","入帳戶名","支出金額","付款方式","收入備註","支出備註"] 
    excelHeader = ["帳務時間","類型","C","D","收入金額","入帳戶名","G","付款方式","I"]
    # G = {"共同支出","曉仙支出","育玠支出","固定支出"}
    # 英文字母對照：C = 收入類別、消費內容、開銷內容 ; D = 收入明細、支出明細 ; G = 支出金額、支出者 ; I = 收入備註、支出備註
    # 將欲填入 excel 裡面的 google 表單的 key 順序，命名為 excelHeader
    # excelHeader 就是 google sheet 的標題列
    # print([sheet_data[0][key] for key in excelHeader]) # list comprehension（列表生成式）
    
    row = 1
    inToExcel = []

    strMsg = "start add data to excel"
    WriteLog(Logger, strMsg)
    # 一次填入多個 column
    for row_data in sheet_data:
        inToExcel = []  # 以一個新的空列表，儲存迴圈中的數據
    
        for key in excelHeader:
            # 按照 excelHeader 的順序去跑 key
            # print(key)
            # print(sheet_data[0][key])
            # 印出 key 對應的 value
        
            # 支出
            if row_data["類型"] == "支出":
                
                # 支出：日常消費
                if row_data["支出類別"] == "日常消費":
                    
                    # 會有填入excel格子的問題，多餘的空格？
                    # 把 google sheet 的 G 吃進來
                    if key == "C":
                        inToExcel.append(row_data["消費內容"])
                    elif key == "D":
                        inToExcel.append(row_data["支出明細"])
                    elif key == "G":
                        if row_data["支出者"] == "曉仙":
                            inToExcel.append("")
                            inToExcel.append(row_data["支出金額"])
                            inToExcel.append("")
                            inToExcel.append("")
                        elif row_data["支出者"] == "育玠":
                            inToExcel.append("")
                            inToExcel.append("")
                            inToExcel.append(row_data["支出金額"])
                            inToExcel.append("")
                        elif row_data["支出者"] == "共同分擔":
                            inToExcel.append(row_data["支出金額"])
                            inToExcel.append("")
                            inToExcel.append("")
                            inToExcel.append("")
                        elif row_data["支出者"] == "固定開銷":
                            inToExcel.append("")
                            inToExcel.append("")
                            inToExcel.append("")
                            inToExcel.append(row_data["支出金額"])
                    elif key == "I":inToExcel.append(row_data["支出備註"])
                    else:
                        inToExcel.append(row_data[key])
        
                # 支出：固定開銷
                elif row_data["支出類別"] == "固定開銷":
                    
                    # excel 只需要填入「固定支出」即可
                    # 把 google sheet 的 G 吃進來
                    if key == "C":
                        inToExcel.append(row_data["開銷內容"])
                    elif key == "D":
                        inToExcel.append(row_data["支出明細"])
                    elif key == "G":
                        if row_data["支出者"] == "曉仙":
                            inToExcel.append("")
                            inToExcel.append(row_data["支出金額"])
                            inToExcel.append("")
                            inToExcel.append("")
                        elif row_data["支出者"] == "育玠":
                            inToExcel.append("")
                            inToExcel.append("")
                            inToExcel.append(row_data["支出金額"])
                            inToExcel.append("")
                        elif row_data["支出者"] == "共同分擔":
                            inToExcel.append(row_data["支出金額"])
                            inToExcel.append("")
                            inToExcel.append("")
                            inToExcel.append("")
                        elif row_data["支出者"] == "固定開銷":
                            inToExcel.append("")
                            inToExcel.append("")
                            inToExcel.append("")
                            inToExcel.append(row_data["支出金額"])
                    elif key == "I":
                        inToExcel.append(row_data["支出備註"])
                    else:
                        inToExcel.append(row_data[key])
                
            # 收入
            elif row_data["類型"] == "收入":
                if key == "C":
                    inToExcel.append(row_data["收入類別"])
                elif key == "D":
                    inToExcel.append(row_data["收入明細"])
                elif key == "I":
                    inToExcel.append(row_data["收入備註"])
                elif key == "G":
                    inToExcel.append("")
                    inToExcel.append("")
                    inToExcel.append("")
                    inToExcel.append("")
                else:
                    inToExcel.append(row_data[key])
            
            else:
                # 如果 key 對應的 value 是空字串，則填入「空字串」
                print(key)
                inToExcel.append("")
                # elif sheet_data[0]["收入金額"] == "":
                #     inToExcel.append("")
                # elif sheet_data[0]["入帳戶名"] == "":
                #     inToExcel.append("")
            
                # 把 google sheet 的 G 吃進來
                # if sheet_data[0]["支出者"] == "曉仙":
                #     sheet.range('H' + str(row)).value = sheet_data[0]["支出金額"]
                #     sheet.range('I' + str(row)).value = ""
                # elif sheet_data[0]["支出者"] == "育玠":
                #     sheet.range('I' + str(row)).value = sheet_data[0]["支出金額"]
                #     sheet.range('H' + str(row)).value = ""
        
        print(inToExcel)
        
        # 將 value 填入 excel 對應的行
        # myExcel.sheet.range('A'+str(startIndex_A)).value = sheet_data[0]  # 把右邊的資料（sheet_data[0]）填入 A 行的 startIndex_A（length_A+1）格子
        # myExcel.sheet.range('A'+str(startIndex_A)).options(transpose = True).value = sheet_data[0]['日期'],   # 把右邊的資料（sheet_data[0]）經過轉置（options(transpose = True)）填入 A 行的 startIndex_A（length_A+1）格子
        # myExcel.sheet.range('A'+str(startIndex_A)).options(transpose = False).value = list(sheet_data[0].values())  # 取出 sheet_data[0] 這個 dict 當中的 values，但取出的型態（type）會是 dict_values 需要轉換成可以讓 excel 吃到的 list
        myExcel.sheet.range('A'+str(startIndex_A)).options(transpose = False).value = inToExcel
        startIndex_A += 1

    strMsg = "start save and close excel"
    WriteLog(Logger, strMsg)
    # 自動儲存 excel
    myExcel.workbook.save()
    # 自動關閉 excel
    myExcel.workbook.close()
    
    
    
    strMsg = "start delete google sheet data"
    WriteLog(Logger, strMsg)
    # 刪除 google sheet 當天的資料
    # 獲取數據長度
    data_length = len(sheet_data)
    
    # 删除除标题行以外的数据
    if data_length > 1:
        sheet_test01.delete_rows(2, number = data_length)
    
    # 選擇要操作的工作表
    # sheet_name = '記帳暫存'
    # sheet = spreadSheet.sheet1[sheet_name]        
    # # 獲取標題行
    # title_row = sheet.range('0:0')    
    # 獲取最後一行的行號
    # last_row = sheet.cells.last_cell.row
    # print(last_row)    
    # 刪除除標題行以外的資料
    # if last_row > 1:
    #     sheet.range(f"2:{last_row}").clear()  # 清空除標題行以外的資料
    # strMsg = "finish"
    # WriteLog(Logger, strMsg)