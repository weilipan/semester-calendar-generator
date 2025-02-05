# Semester Calendar Generator for Google Sheets

這是一個利用 Google Apps Script 自動產生學期行事曆的 side project，並整合臺灣假日日曆以標示假日。  
行事曆具有以下功能：

- **動態產生學期行事曆：** 根據使用者輸入的起始及結束日期，自動產生以週為單位的行事曆。
- **臺灣假日整合：** 從臺灣假日日曆（ID: `zh-tw.taiwan#holiday@group.v.calendar.google.com`）取得假日事件，將假日日期標示為紅色文字。
- **格式美化：** 行事曆中包含標題列、資料列，並在最上方插入教師說明列（合併儲存格，預設文字為「教師說明」，背景白色、文字黑色、靠左且靠上對齊，自動換行，行高 200 像素）。
- **互動式輸入：** 執行時會彈出對話框，預設值根據執行當下的日期自動填入（例如起始年份預設為當前年份）。

## 安裝與使用

1. **在 Google Sheets 中設定：**  
   - 開啟 Google Sheets，選擇 **工具 > 指令碼編輯器**。  
   - 將以下 `Code.gs` 程式碼貼入指令碼編輯器並儲存（可命名為 `SemesterCalendarGenerator`）。

2. **執行程式：**  
   - 重新整理試算表後，選單中會出現「自定義功能」選項。  
   - 點選「自定義功能」→「產生學期行事曆」，依照對話框提示輸入起始與結束日期（預設值為執行當下日期），系統將自動產生學期行事曆。

## 程式碼說明

- **onOpen**  
  當試算表開啟時，自動在選單中加入「自定義功能」→「產生學期行事曆」。

- **generateCalendar**  
  主程式，負責：  
  1. 讀取使用者輸入的起始與結束日期（預設值為當下日期）。  
  2. 取得臺灣假日日曆事件，建立日期對應假日的物件。  
  3. 根據起始日期回溯至該週的星期日，依週產生行事曆資料（包含週次、月份、星期日到星期六的日期、課程目標／課程簡介／重要事項欄）。  
  4. 將資料寫入工作表，並設定格式（欄寬、字型、框線、交替底色、假日標示）。  
  5. 最後，在最上方插入教師說明列，並設定合併儲存格、預設文字及格式。

## 開發環境

- **Google Apps Script**：利用 JavaScript 語法編寫，直接在 Google Sheets 的指令碼編輯器中運行。

