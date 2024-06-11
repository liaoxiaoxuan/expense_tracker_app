# 雲端整合財務記帳系統  

這款記帳軟體可以讀取 Google Sheets 中的數據，並將之寫入 Excel 中。  

初衷是為了紀錄將來的家戶現金流量，但翻找了幾個現有的 App，總覺得和心目中的記帳形式有落差。  

於是參考了幾篇網路上既有的 Excel 表格設計，加以改編成較適合自身記帳習慣，以及觀察家戶現金流量的樣貌。  

美中不足的是，Excel 較受限於電腦，且不若 App 那般，具有隨時隨地紀錄的便利性。  

幾經思索，倘若我可以依據 Excel 的格式與內容，撰寫一份 Google 表單，再將表單結果代入 Excel，如此也能擁有 App 那般隨時紀錄的特性。  

因此，這個程式便於焉而生。

## 安裝

1. 安裝所需的 Python 庫：
   ```bash
   pip install pygsheets
   pip install xlwings
   ```

2. 創建 Google Cloud API 服務帳戶，並下載驗證文件 `expense-tracker-app3-b8d7f50165a5.json`，將其放置於 `./local/` 目錄中。

## 使用方法  

1. 在 expense_tracker 函數中設置 Google Sheets 的表單 URL 和工作表名稱。

2. 執行程式，將 Google Sheets 中的數據讀取並寫入 Excel 中。

## 程式結構

+ `call_sheetrange.py`：包含 Excel 範圍操作的程式碼。  
+ `MyLogger.py`：用於記錄日誌的程式碼。  
+ `__author__.py` 和 `__version__.py`：作者和版本訊息。  

## 使用範例

```
from expense_tracker import expense_tracker

# 開始費用追蹤
expense_tracker()
```

## 常見問題

1. 如何修改 Google Sheets 表單 URL 和工作表名稱？  
在 `expense_tracker` 函數中修改 `sheet_url` 和 `worksheet_by_title` 函數的參數。

2. 是否支持自定義報告格式？
是的，你可以根據需要修改 `excelHeader` 變數以及填入 Excel 的數據處理邏輯。

## 主要成果  

1. 自動化  
   實現無縫的數據傳輸和處理，減少了手動輸入的錯誤，並藉由自動化以節省時間。

2. 整合  
   成功地整合 Google Sheets 與 Excel，允許隨時更新和便捷的資料檢索。

3. 效率  
   提升財務管理工作的流程，以致更準確和及時的產出現金流量報表。

4. 技術技能  
   展示了使用 Google Cloud API、Pygsheets 和 Xlwings 開發系統，以及增進財務解決方案的能力。