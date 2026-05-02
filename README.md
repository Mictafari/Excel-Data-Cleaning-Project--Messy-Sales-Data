# Excel-Data-Cleaning-Project--Messy-Sales-Data
This project demonstrates a complete data cleaning workflow using **Microsoft Excel** and **Power Query**. Raw, inconsistent sales data was transformed into a clean, analysis‑ready dataset with error detection and data validation.
## The Problem

The original dataset had multiple quality issues:

| Issue | Example |
|-------|---------|
| Mixed date formats | `01/15/2025`, `2025-01-20`, `Jan 25 2025` |
| Extra spaces in text | `" michael ansah "` |
| Inconsistent capitalization | `"AMA SERWAA"`, `"kofi mensah"` |
| Blank region values | Empty cells in Region column |
| Text stored as numbers | Quantity and Price stored as text |
| Inconsistent status values | `"shipped"`, `"Shipped"`, `" Pending "` |
| No error checking | Wrong Quantity or Price would go unnoticed |

## Tools & Skills Used

- **Power Query** – Data transformation and cleaning
- **Excel Data Validation** – Restricting status column inputs
- **Excel Formulas** – `IFERROR`, `AND`, structured references
- **Excel Tables** – Structured data storage

## Cleaning Steps Performed

### 1. Power Query Transformations

| Step | Action |
|------|--------|
| Trim spaces | Removed leading/trailing spaces from all text columns |
| Capitalization | Applied "Capitalize Each Word" to CustomerName and Product |
| Standardized case | Converted Status column to lowercase |
| Fixed dates | Converted mixed date formats to proper Excel dates |
| Fixed numbers | Converted Quantity, UnitPrice, TotalSales to correct data types |
| Blank handling | Replaced empty Region cells with "Unknown" |

### 2. Data Validation (Excel)

After loading the clean data back to Excel, I added a dropdown list to the **Status** column:

- Allowed values: `shipped`, `delivered`, `pending`, `cancelled`

This prevents data entry errors in the future.

### 3. Error Detection Column

Added a **Check** column with the following simple excel formula:

=IFERROR(IF(AND([@Quantity]>0, [@UnitPrice]>0), "OK", "Check Values"), "Error")

The use of that formula is to make sure the quantity column and unitprice values are above 0. So if its above 0 u see "ok" in the check column else u see "check values" and In all if an error persist the iferror formula takes care of that.
