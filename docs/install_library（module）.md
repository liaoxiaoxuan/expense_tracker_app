# 如何安裝套件／函式庫／模組

```bash
pip install "套件／函式庫／模組的名稱"
pip freeze   # 顯示已安裝的套件
pip freeze > requirements.txt   # 將已安裝的套件寫入requirements.txt檔
pip install -r requirements.txt   # 安裝寫在requirements.txt檔裡的所有套件（專案搬家用）
pip install -U "套件／函式庫／模組的名稱"   # 更新版本
```