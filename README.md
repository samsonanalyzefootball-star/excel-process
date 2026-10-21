# Excel Data Cleaner

一個使用 Streamlit 建立的 Excel 資料清洗工具。

## 功能

- 上傳 Excel 檔案
- 自動偵測缺失值、NULL、N/A、unknown
- 檢查日期、日期時間、電話與 Email 格式
- 檢查數值異常值
- 將資料分為 Clean_Data 與 Incomplete_Data
- 產生 Data_Issues 問題明細
- 下載清洗後 Excel

## 本機執行

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 部署到 Streamlit Community Cloud

1. 將 `app.py` 與 `requirements.txt` 上傳到 GitHub repository。
2. 前往 [Streamlit Community Cloud](https://share.streamlit.io/)。
3. 選擇 `New app`。
4. 連結你的 GitHub repository。
5. 設定：
   - Repository：你的 GitHub repo
   - Branch：`main`
   - Main file path：`app.py`
6. 點選 `Deploy`。# excel-process
