# 🧹 Excel Data Cleaning & Transformation: Product Dataset

> From messy rows to analysis-ready data: handling missing values, fixing text, checking duplicates, splitting and merging columns, and formatting, all with Excel formulas.

**Author:** Hari Hara Kumar
**Source:** Entri Academy, Data Analytics Module
**Tool:** Microsoft Excel
**Dataset:** `Product_DB` Excel Table (34 records received, 31 after cleaning)

---

## 🎯 Objective
Real-world data is rarely clean. This project prepares a product dataset for analysis by:

1. Handling missing values in `Price` and `Category`
2. Standardising inconsistent text in `Product Name` and `Category`
3. Checking for and removing duplicate records
4. Splitting `Product ID` into usable parts and merging `Brand Name` with `Product Name`
5. Applying currency and date formatting
6. Highlighting patterns with conditional formatting

All cleaning is done with **formulas in new helper columns** (suffix `CL` = cleaned), so the original raw columns stay untouched and every step can be audited.

---

## 🔎 Data Quality Audit (What Was Wrong)

| # | Issue | Where | Count | Fix |
|---|---|---|---|---|
| 1 | Missing price | Sony Headphones, Coleman Camping Tent, Ray-Ban Sunglasses | **3** of 31 | Fill with `MODE.SNGL` |
| 2 | Missing category | Backpack, Sneakers, Coffee Maker, Fitness Tracker | **4** of 31 | Rule-based fill |
| 3 | Inconsistent case in product names | `laptop` (x2), `smartphone`, `headphones` | **4** cells | `PROPER` |
| 4 | Exact duplicate rows | Laptop Bag (21-AUG-CA), Headphones Bose (16-APR-ES), Laptop HP (17-JUN-IN) | **3** | Remove Duplicates |
| 5 | Product ID packs three facts into one text | e.g. `28-JAN-US` | all 31 | `LEFT`, `MID`, `RIGHT` |
| 6 | Brand and product kept in separate columns | e.g. `Dell` + `Laptop` | all 31 | `CONCAT` |

---

## 🛠️ Step-by-Step Cleaning

### 1️⃣ Handling Missing Prices

**Problem:** 3 of 31 products have no price, so any total or average would silently ignore them.

**Formula (`Price ($) CL`):**
```excel
=IF(ISBLANK([@[Price ($)]]), MODE.SNGL([Price ($)]), [Price ($)])
```

**How it works**
- `ISBLANK` tests whether the price cell is empty.
- If empty, `MODE.SNGL` returns the most frequent price in the column (**130**).
- If not empty, the original price is kept.

**Why the mode and not the mean?** Prices are right-skewed: laptops and phones sit near 1,000, while most items cost 30 to 160. The mean (297.86) is pulled up by a few expensive items and would overstate the price of a pair of sunglasses or a tent. The median is also 130, which supports using 130 as a typical value.

| Product | Raw price | Cleaned price |
|---|---|---|
| Sony Headphones | *(blank)* | 130 |
| Coleman Camping Tent | *(blank)* | 130 |
| Ray-Ban Sunglasses | *(blank)* | 130 |

**Impact:** average price moves from **297.86** (28 priced rows) to **281.61** (all 31 rows).

> ⚠️ Prices 130 and 150 each appear 3 times, so the mode is a tie. `MODE.SNGL` returns the value that appears first in the column (130). Imputed prices are estimates, not facts, and should be flagged as such in any report.

---

### 2️⃣ Handling Missing Categories

**Problem:** 4 products have no category.

**Formula (`Category CL`):**
```excel
=IF([@Category]<>"", [@Category], IF([@[Product Name]]="Coffee Maker", "Kitchen", "Fashion"))
```

**How it works**
- If a category already exists, keep it.
- If blank and the product is a Coffee Maker, assign **Kitchen** (other coffee makers in the data are Kitchen).
- Any other blank becomes **Fashion**.

| Product | Raw category | Cleaned category | Method |
|---|---|---|---|
| Backpack (North Face) | *(blank)* | Fashion | Formula fallback |
| Sneakers (Adidas) | *(blank)* | Fashion | Formula fallback (Nike Sneakers is Fashion) |
| Coffee Maker (Nespresso) | *(blank)* | Kitchen | Formula rule |
| Fitness Tracker (Xiaomi) | *(blank)* | Electronics | Typed manually (other fitness trackers are Electronics) |

**Cleaned category counts:** Electronics 13 · Fashion 8 · Kitchen 6 · Outdoor 2 · Accessories 2 (total 31).

> ⚠️ The "everything else is Fashion" fallback only works because the blanks in this dataset happen to be Fashion items. The Fitness Tracker had to be typed by hand for exactly that reason. See the Known Limitations section below for a more robust approach.

---

### 3️⃣ Standardising Text

**Problem:** 4 product names are lowercase (`laptop`, `smartphone`, `headphones`) while the rest are capitalised, so `laptop` and `Laptop` would be counted as different products in a PivotTable.

**Formula (`Product Name CL`):**
```excel
=PROPER([@[Product Name]])
```

`PROPER` capitalises the first letter of each word and lowercases the rest, fixing all rows in one step instead of many Find and Replace passes.

| Before | After |
|---|---|
| laptop | Laptop |
| smartphone | Smartphone |
| headphones | Headphones |
| T-shirt | T-Shirt |

> `PROPER` also capitalises after a hyphen, so `T-shirt` became `T-Shirt`. This is consistent but is a style change, not a correction.

**Category typos:** all five category values (Electronics, Fashion, Kitchen, Outdoor, Accessories) were already spelled consistently in this workbook, so no Find and Replace correction was required for `Category`.

---

### 4️⃣ Removing Duplicates

**Action:** Data → Remove Duplicates, applied across all columns.

**Finding:** **3 exact duplicate rows** were removed, taking the dataset from **34 to 31 records**.

| Product ID | Product | Brand | Price | Qty | Category |
|---|---|---|---|---|---|
| 21-AUG-CA | Laptop Bag | Samsonite | 50 | 35 | Accessories |
| 16-APR-ES | headphones | Bose | 250 | 20 | Electronics |
| 17-JUN-IN | Laptop | HP | 950 | 25 | Electronics |

Repeated rows would double-count prices in totals, averages and the mode, so duplicates should be removed before imputing or summarising.

---

### 5️⃣ Splitting and Merging Columns

**Splitting `Product ID`** (format `DD-MON-CC`, e.g. `28-JAN-US`):

| New column | Formula | Logic | Example |
|---|---|---|---|
| Product ID No (Day) | `=LEFT([@[Product ID]],2)` | First 2 characters | `28` |
| Date Of.Manu (Month) | `=MID([@[Product ID]],4,3)` | 3 characters starting at position 4 | `JAN` |
| Country Code | `=RIGHT([@[Product ID]],2)` | Last 2 characters | `US` |

The hyphens are skipped automatically because `MID` starts at position 4 and takes exactly 3 characters.

**Merging brand and product:**
```excel
=CONCAT([@[Brand Name]], " ", [@[Product Name CL]])
```
The `" "` adds a space between the two parts. Using the cleaned name (`Product Name CL`) means the merged label is also consistently capitalised.

| Product ID | Day | Month | Country | Brand | Product Name CL | Merged |
|---|---|---|---|---|---|---|
| 28-JAN-US | 28 | JAN | US | Dell | Laptop | Dell Laptop |
| 11-APR-US | 11 | APR | US | Samsung | Smartphone | Samsung Smartphone |
| 16-APR-ES | 16 | APR | ES | Bose | Headphones | Bose Headphones |
| 08-DEC-DE | 08 | DEC | DE | Nespresso | Coffee Maker | Nespresso Coffee Maker |

**What the split reveals:** 12 countries across the data. **CA** is the most common (10 of 31 products), followed by **US** and **UK** (6 each). All 12 months appear.

---

### 6️⃣ Number Formatting
- **Price:** `Price ($) CL` is formatted as currency with 2 decimals (e.g. `1,000.00`).
- **Date:** the month column is stored as a three-letter text value because the Product ID contains no year. See the limitations section for how to turn it into a true `DD-MM-YYYY` date.

---

### 7️⃣ Conditional Formatting

| Column | Rule | Purpose |
|---|---|---|
| `Price ($) CL` | **Data bars** (min to max) | Shows relative price size at a glance: laptops and phones stand out, accessories shrink |
| `Category CL` | **Custom rule: text contains "Electronics"** → pink fill | Instantly highlights all 13 Electronics products |
<img width="1842" height="512" alt="image" src="https://github.com/user-attachments/assets/98b062b4-5cab-4082-8eff-4e027ea4efd9" />


---

## 📊 Before vs After

| Measure | Before | After |
|---|---|---|
| Records | 34 | 31 |
| Exact duplicate rows | 3 | 0 |
| Missing prices | 3 | 0 |
| Missing categories | 4 | 0 |
| Lowercase product names | 4 | 0 |
| Average price | 297.86 (28 values) | 281.61 (31 values) |
| Product ID | 1 text column | Day, Month, Country Code |
| Brand + product | 2 columns | 1 merged label |

---

## 🧮 Formula Reference

| Task | Function(s) | Why it was used |
|---|---|---|
| Fill missing price | `IF`, `ISBLANK`, `MODE.SNGL` | Replace blanks with the most common price |
| Fill missing category | `IF` (nested) | Rule-based fill without overwriting existing values |
| Fix text case | `PROPER` | One formula standardises every name |
| Split Product ID | `LEFT`, `MID`, `RIGHT` | Extract day, month and country from fixed positions |
| Merge columns | `CONCAT` | Combine brand and product into one label |
| Remove duplicates | Data → Remove Duplicates | Detect repeated records |
| Highlight patterns | Data Bars, Custom Rule | Visual checks on price and category |

> `CONCAT` and `MODE.SNGL` need Excel 2019 or Microsoft 365. In older versions use `&` and `MODE`.

---
## 📁 Repository Structure
```
excel-data-cleaning-product-dataset/
├── README.md
├── data/
│   └── Data_Cleaning_and_Transformation.xlsx
├── docs/
│   └── Excel_Assignment_2_Documentation.docx
└── screenshots/
    └── cleaned-dataset.png
```

## 🚀 Skills Practised
- Missing value detection and imputation (`IF`, `ISBLANK`, `MODE.SNGL`)
- Text standardisation (`PROPER`)
- Duplicate detection and removal
- Column splitting and merging (`LEFT`, `MID`, `RIGHT`, `CONCAT`)
- Number formatting and conditional formatting
- Documenting cleaning decisions and their trade-offs

## ✅ Conclusion
The product dataset went from 34 rows with 7 missing values, 3 duplicates, 4 inconsistent names and a packed ID column to a structured 31-record table, with the raw columns preserved beside the cleaned ones. Each choice (mode for prices, rule-based categories, `PROPER` for text) was documented with its reasoning and its limits, which is the habit that matters most in real data cleaning.
