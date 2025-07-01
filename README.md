
# 📄 README

## Opus Invoice Robot

**Opus Invoice Robot** is an automation developed for **Aarhus Kommune**. It streamlines the processing, matching, and forwarding of supplier invoices in the Opus system, reducing manual effort and improving traceability.

---

## 🚀 Features

✅ **Automated Login & Credential Management**  
Logs in to Opus via Selenium and updates expired passwords automatically in OpenOrchestrator.

📥 **Download Invoice Lists**  
Fetches Excel reports for Naturafdelingen and Vejafdelingen EAN numbers.

🧾 **Extract and Filter References**  
Parses invoice references and dates, filtering by creditor and date.

🔍 **Match Invoices to Employees**  
Queries AZIdent mappings from SQL Server and uses Levenshtein distance for fuzzy matching.

📡 **Forward Invoices in Opus**  
Navigates to each invoice in Opus and forwards it to the responsible employee.

🗑️ **Cleanup and State Update**  
Deletes downloaded files and updates the `Bilagsdato` constant.

---

## 🧭 Process Flow

1. Fetch credentials and constants from OpenOrchestrator
2. Initialize ChromeDriver (Selenium)
3. Log in to Opus  
   - If password expired, generate and save a new one
4. Download invoice Excel files for each department
5. Extract invoice references and oldest registration dates
6. If references exist:
   - Search Opus and export supplementary Stark data
7. Query SQL Server for AZIdent mappings
8. Match invoices to employees
   - Fuzzy match names if AZ-identifiers are missing
9. Forward invoices to responsible employees
10. Remove temporary files
11. Update `Bilagsdato`

---

## 🔐 Privacy & Security

- All interactions occur over HTTPS
- Credentials are stored securely in OpenOrchestrator
- No sensitive data is persisted beyond processing
- Temporary files are removed after use

---

## ⚙️ Dependencies

- Python 3.10+
- Selenium
- pandas
- pyodbc
- python-Levenshtein

---

## 👷 Maintainer

Gustav Chatterton  
*Digital udvikling, Teknik og Miljø, Aarhus Kommune*
