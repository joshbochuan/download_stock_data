# 股票爬蟲程式使用說明

這個程式可以幫助您爬取台灣股市每日成交資料並保存到 Excel 檔案中。此外，我們還提供了已經打包成 `main.exe` 執行檔，讓您無需安裝 Python 環境，即可直接運行程式。

## 如何使用

1. **下載程式**

   請從本專案的 GitHub 儲存庫中下載 `main.exe` 執行檔。

2. **設定參數**

   在您的計算機上創建一個名為 `setting.xml` 的設定檔，並根據您的需求配置以下參數：

   ```xml
   <params>
       <url>https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY</url>
       <excelName>微星</excelName>
       <startYear>2022</startYear>
       <startMonth>01</startMonth>
       <endYear>2023</endYear>
       <endMonth>08</endMonth>
       <stockNo>2377</stockNo>
   </params>
   ```
- `<url>`: 資料爬取的網址，預設為台灣證券交易所的股票成交資料頁面。
- `<excelName>`: 要保存的 Excel 檔案名稱，程式會自動在後面加上 `.xlsx` 副檔名。
- `<startYear>`、`<startMonth>`: 資料爬取的起始年月。
- `<endYear>`、`<endMonth>`: 資料爬取的結束年月。
- `<stockNo>`: 欲爬取的股票代號。

## 注意事項

- 程式在爬取資料時會發送 HTTP 請求，請確保您的網絡連接正常。
- 爬取大量資料可能會對伺服器產生壓力，請勿過度使用或濫用程式。
- 在使用 `main.exe` 執行檔時，請確保 `setting.xml` 與執行檔位於同一目錄下。
- 如果您需要更多幫助或有任何問題，請隨時聯繫程式作者。
- 請合法使用此程式，遵守相關法規，並注意保護他人的權益。
