# 🔎 Supplier Data Hider Utility

This VBA macro provides a **quick way to clean up large datasets** in the `Analysis` worksheet by hiding **empty columns and rows**.  
It also includes a companion macro to **unhide all rows and columns**, restoring the sheet to its original view.

---

## 🚀 Features
- Automatically **hides supplier-related columns** (grouped in sets of 5) if their data is completely empty.  
- Hides rows in the `Analysis` sheet if no data exists across supplier columns.  
- Provides **Unhide function** to restore all hidden rows and columns.  
- Includes:
  - Execution timing (`Start_Time`, `Time_Taken`)  
  - User tracking (`UserName`)  
  - Status logging (`Success/Failure`)  

---

## 📋 Workflow
### Macro 1: `hidesupplier`
1. Loops through each **5-column block** (starting from column J).  
2. If the sum of values = `0`, the block of 5 columns is hidden.  
3. Loops through rows and hides any row where all supplier data is `0`.  
4. Updates the **Dashboard** with execution details.  

### Macro 2: `UnhideAllSupplier`
1. Restores all hidden rows and columns in the `Analysis` sheet.  
2. Resets the view for full dataset inspection.  

---

## 🛠️ Usage
1. Place this macro in your VBA project.  
2. Ensure your workbook contains the following:
   - A sheet named `Analysis` with supplier data starting at **Row 5** and **Column J**.  
   - A `Dashboard` sheet with status fields: `[Status]`, `[Start_Time]`, `[Time_Taken]`, `[UserName]`.  
3. Run:
   - `hidesupplier` → Hide empty rows & columns.  
   - `UnhideAllSupplier` → Restore full dataset.  

---

## ⚡ Example
### Before
| Supplier A | Supplier B | Supplier C | Supplier D | Supplier E |
|------------|------------|------------|------------|------------|
| 100        | 0          | 50         | 0          | 0          |
| 0          | 0          | 0          | 0          | 0          |
| 200        | 0          | 0          | 0          | 30         |

### After Running `hidesupplier`
- Columns with only zeros → hidden.  
- Rows with only zeros → hidden.  

---

## 📂 Functions
- `hidesupplier` → Hide empty supplier data.  
- `UnhideAllSupplier` → Reset to full view.  

---

## ⚠️ Notes
- Uses `OptimizedMode` for speed (screen updates/events disabled during execution).  
- Works only if the sheet names (`Analysis`, `Dashboard`) exist.  
- Designed for **Excel Desktop (Classic View)**.  

---

