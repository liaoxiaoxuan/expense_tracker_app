# 記帳軟體

這是一個記帳軟體，可以將 Google Sheets 中的數據讀取並寫入 Excel 中。

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