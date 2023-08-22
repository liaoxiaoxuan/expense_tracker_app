import pygsheets

auth_file = "../static/expense-tracker-app-392308-a0f2bb4f12a1.json"
googleCloud = pygsheets.authorize(service_file = auth_file)

# setting sheet
sheet_url = "https://docs.google.com/spreadsheets/d/1yHXGgmFkMhYJYzsunphNErEkDntMP5MiWYyxJx-fqUs/edit#gid=881554370" 
spreadSheet = googleCloud.open_by_url(sheet_url)

#選取by名稱
sheet_test01 = spreadSheet.worksheet_by_title("表單回應 1")
print(sheet_test01[1])
for row in sheet_test01:  # 印出worksheet的所有row（橫排） # column（豎排）要再另外查
    print(row)
print(type(sheet_test01.cell))
print(type(sheet_test01[1]))

# read
A1 = sheet_test01.cell('A1')
print(A1)
print(A1.value)
