# 📬 Automated Gmail-to-Excel Data Pipeline using Power Query & Google Apps Script

This project demonstrates a fully automated solution to fetch `.xlsx` email attachments directly from Gmail, decode and transform the data using Power Query, and load it into Excel for analysis and reporting.

It eliminates the need for manual downloads and file handling, enabling a scalable and refreshable workflow for daily reporting.

---

## 🚀 Project Overview

- 📧 **Email Integration**: Uses Gmail and Google Apps Script to access and filter email attachments.
- 🔗 **Web API Bridge**: Google Apps Script is deployed as a Web App, which serves Gmail attachment data in a format directly consumable by Power Query.
- 🧾 **Base64 to Binary Conversion**: Power Query decodes Base64 strings into binary Excel workbooks using `Binary.FromText`.
- 📊 **Structured Data Loading**: Extracts and loads the Excel attachment data using `Excel.Workbook()` in Power Query.
- 🧼 **Data Cleaning**: Applies transformations such as header removal and whitespace normalization in columns like `Customer Name`, `City`, and `Product`.

---

## 🧠 Skills Demonstrated

- Power Query (M Language)
- Google Apps Script (Web App deployment)
- Data transformation and automation
- Excel-based reporting workflows
- Integration of cloud-based email and desktop Excel tools

---

## 🔧 Technologies Used

- **Google Apps Script**
- **Power Query in Excel**
- **Gmail (IMAP enabled)**
- **Excel.Workbook / Binary.FromText**

---

## 📈 Use Cases

- Automated sales or operations reporting
- Centralized data extraction from multiple email sources
- Zero-manual-effort ETL pipeline for Excel-based analytics

---
