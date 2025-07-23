# Mannanova Google Drive Automation

This repository contains the complete Google Apps Script code used to **automate data transfer from a source spreadsheet to multiple destination spreadsheets**. It is built for Mannanova's internal operations to streamline and error-proof their data handling workflows.


## ✨ Features

* ✅ Adds a **custom menu** to the Google Sheet for user-friendly execution.
* 📅 Automatically triggered when the spreadsheet is opened.
* 📅 Dynamically detects **data blocks** marked with `TRUE` or `FALSE`.
* 🔄 Extracts only **complete rows** with all required data.
* 🔍 Supports **custom column counts** per block based on headers.
* 🔗 Appends rows to **external spreadsheets** using a configured mapping.
* 🖊️ Automatically updates processed blocks to `TRUE` to avoid reprocessing.


## 📁 Folder Structure

```
📄 Code.gs         # Main Google Apps Script logic
📄 README.md       # Project documentation (this file)
```

---

## 🚧 Use Case

This script is designed for teams that:

* Receive bulk data entries into a single Google Sheet ("Pastepad")
* Need to route subsets of that data into multiple target spreadsheets
* Want to avoid duplicate entries and incomplete data


## 🚀 How It Works

1. **Setup**

   * A custom menu titled `Data Automation` is added on sheet open.
   * Clicking `Run Script` triggers the data extraction and push process.

2. **Data Format in Pastepad**

   * Column A = `TRUE` or `FALSE` indicating a new block
   * Column B = File identifier (used to route data)
   * Next row = Headers for that block
   * Following rows = Data rows

3. **Processing Logic**

   * Each block is read and parsed based on its header column count.
   * Only rows with **all non-empty cells** are considered valid.
   * Data is sent to a corresponding spreadsheet and sheet based on `fileMap`.
   * Once copied, the block's flag in column A is updated to `TRUE`.


## 📝 fileMap Configuration

Found in `Code.gs`, this object defines which file and sheet each identifier maps to:

```javascript
const fileMap = {
  "REG-101": {
    fileId: "1Hb-agA100ZwuvPf9lAlHk-s5h5HG5fVWLePO7EYBzbY",
    sheetName: "Inventaire"
  },
  "REG-102 RRIBC": {
    fileId: "19Ew_W50zatGad97FCpfgzAptkSfrJExF3cnDjDY8ZEc",
    sheetName: "IBC Use and Clean"
  },
  // Add more mappings as needed
};
```


## ▶️ How to Use

1. Open the **Pastepad** spreadsheet.
2. Go to the menu: **Data Automation → Run Script**
3. The script will:

   * Detect each new block
   * Extract, validate, and route data
   * Append it to the mapped destination sheet
   * Mark the block as `TRUE` to avoid duplication


## 🔐 Required Permissions

When running the script for the first time, you'll be prompted to authorize:

* Access to manage Google Sheets
* Permission to connect and write to external spreadsheets


## 📊 Example Pastepad Structure

| A (Processed) | B (File Name) | C (Header1) | D (Header2) | E ... |
| ------------- | ------------- | ----------- | ----------- | ----- |
| FALSE         | REG-101       |             |             |       |
|               | Activity      | Product     | Date        | ...   |
|               | Production    | Item A      | 2025-06-01  | ...   |
|               | Production    | Item B      | 2025-06-01  | ...   |
| FALSE         | REG-102 RRIBC |             |             |       |
|               | Use           | IBC ID      | Name        | ...   |


## 📊 Logs & Debugging

Use `console.log()` inside the script to view:

* Number of rows processed
* File names being accessed
* Skipped or invalid rows
* Block start and end points

All logs can be viewed via `Executions` tab in the Apps Script dashboard.


## 📅 Future Enhancements

* [ ] Reprocessing toggle for `TRUE` blocks
* [ ] Email summary after execution
* [ ] Versioning or backup of data before appending
* [ ] Visual feedback in the Pastepad (e.g., color-coded status)

