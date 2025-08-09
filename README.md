# Test-Sys

This repository contains various scripts and utilities for sales, EOD, and reporting.

---

## 📂 Folder Overview

| Folder Name               | Purpose                                                                 | Status        | Notes |
|---------------------------|-------------------------------------------------------------------------|--------------|-------|
| **CSV_TO_XML**            | Converts CSV sales files to XML format for further processing           | 🧪 Testing    | Still being validated; may have edge cases with empty CSVs |
| **generated_palda**       | Generates a report (MS Word format) with a **custom time range**         | ✅ Stable     | Used for daily/weekly reporting |
| **generated_palda_old**   | Generates a report (MS Word format) with a **fixed hardcoded time**      | 🗑 Deprecated | Kept for reference only |
| **generating_excel_presched** | Generates pre-scheduled Excel reports                                | 🧪 Testing    | Output format under review |
| **palda_docx**            | Test MS Word document for PALDA report formatting                       | 🧪 Testing    | Layout experiments only |

---

## 🛠 Development Notes
- Codebase is in **C#** and uses SQLite for local data storage.
- Most report generators rely on the `tblsales`, `tbleod`, and `tblchecker` tables.
- Test cases should be run before moving any "🧪 Testing" modules to production.

---

## 📌 To-Do
- [ ] Finalize `CSV_TO_XML` error handling for missing headers
- [ ] Replace `generated_palda_old` with calls to the new `generated_palda`
- [ ] Write automated tests for `generating_excel_presched`
