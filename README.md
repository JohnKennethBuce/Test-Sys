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
- Codebase is in **python**
- Most report generators is for Daily Reports
- Test cases should be run before moving any "Testing" modules to production.
