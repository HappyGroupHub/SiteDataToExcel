# 【2024 角落微宇宙】台灣微軟企業參訪

## SiteDataToExcel - 將網站資料轉成 Excel 檔案

1. 直接[下載](https://github.com/HappyGroupHub/SiteDataToExcel/archive/refs/heads/master.zip), 或是使用 git clone 指令下載此專案
2. 解壓縮後，進入 SiteDataToExcel 資料夾
3. 在終端機輸入 `pip install -r requirements.txt` 安裝所需套件
4. 在終端機輸入 `python app.py` 即可執行程式
5. 輸入網址，按下 Enter 即可開始爬取網站資料
6. 完成！程式會自動產生一個 名為 `enterprise_visited.xlsx` 的檔案

## 注意事項

- 請確保已安裝 Python 3.9 以上版本
- 請確保已安裝 pip 套件管理工具
- 請確保所有套件包含 requests, beautifulsoup4, openpyxl 都已安裝