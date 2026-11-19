# 學校採購申請單暨支出憑證黏存單 — 桌面自動化系統

以 **Python + Flask + Waitress** 打造的離線桌面應用程式，整合 **PyMuPDF 表格解析**、**RapidOCR / Gemini Vision 影像辨識**、**Google OR-Tools CP-SAT 求解器**與 **Excel COM 真實列印預覽**，協助快速將發票、估價單或明細清單自動填入學校採購單 `.xlsx` 範本。

## 主要功能

- **多來源單據輸入**：
  - 支援上傳 **PDF 報價單／電子發票**（含加密 PDF 密碼解密與 PyMuPDF 表格欄位座標解析）
  - 支援上傳 **發票／估價單照片**（`.jpg` / `.png` / `.webp` / `.bmp`，內建離線 RapidOCR，並可選用 Gemini Vision）
  - 支援 **電子發票文字貼上**（如小北百貨）、**標準明細 CSV**、**財政部電子發票匯出 CSV** 與 **手動編輯表格**
- **Google OR-Tools CP-SAT 雙重最佳化**：
  - **單據金額核對（[ocr_cpsat_parser.py](ocr_cpsat_parser.py)）**：自動過濾付款／找零雜訊列並核對品項加總與總金額。
  - **採購單最佳化分單（[form_splitter.py](form_splitter.py)）**：依單張金額上限與 7 列品項限制分配至最少張請購單，單筆超額品項自動獨立成單。
- **智慧選表與單一工作表輸出（[fill_purchase_forms.py](fill_purchase_forms.py)）**：
  - 依總金額自動切換 `請購單簡化流程-3000元以下適用` 與 `請購單-3000元以上一定要用此表`。
  - 自動偵測各表單列位、保留同表公式、解析跨表 `VLOOKUP` 並僅輸出目標工作表。
- **真實 Excel 列印預覽（[print_preview.py](print_preview.py)）**：呼叫本機 Excel 匯出實際列印排版 PDF 與預覽圖。
- **累積採購紀錄與彙整統計（[history_store.py](history_store.py)）**：自動累積每次產出紀錄、依預算科目統計小計並匯出 CSV 報表。

## 快速啟動

1. 安裝相依套件（建議使用 Python 3.11 或 3.12 以支援 `rapidocr-onnxruntime`）：
   ```powershell
   python -m pip install -r requirements.txt
   ```
2. 雙擊 `啟動桌面程式.bat`，或在終端機執行：
   ```powershell
   python app.py
   ```
3. 執行自動化測試：
   ```powershell
   python -m unittest discover -v
   ```
