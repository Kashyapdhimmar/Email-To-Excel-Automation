# 📥 Gmail to Excel Automation via Power Query

This project fetches `.xlsx` attachments from Gmail emails with a specific subject and loads them directly into Excel using Power Query and Google Apps Script.

## 🔧 Tools Used
- Power Query (Excel)
- Google Apps Script (Web App as JSON API)
- Base64 decoding with `Binary.FromText`
- Data extraction with `Excel.Workbook`

## 💡 Key Features
- No manual downloads of attachments
- Supports filtering by subject and file type
- Auto-decodes and loads Excel files into structured tables
- Cleans data (removes headers, trims spaces)

## 📂 Files
- `EmailToExcel_Pipeline.pq` — Power Query M code
- (Optional) Sample Excel file for reference

## 📈 Use Cases
- Sales report automation
- Daily operational data sync from Gmail
- Dashboard-ready Excel output

---


