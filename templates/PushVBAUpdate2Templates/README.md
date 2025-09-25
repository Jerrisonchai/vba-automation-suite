# VBA Macro – Control Selected Workbook Modules Updating from Source Workbook

## 📌 Overview
The **`ControlSelectedWorkbookModulesUpdatingFromSourceWorkbook`** macro automates the process of updating VBA modules in a selected workbook by replacing them with the modules from a **source workbook**.  

This is particularly useful for **maintaining version control** across multiple Excel VBA projects, ensuring that all destination workbooks are always aligned with the latest approved VBA codebase.

---

## ✨ Features
- Prompts the user to select a **destination workbook** (`.xlsm`) for updating.
- Reads the **source workbook path** from the `Dashboard!C21` cell.
- Removes all existing VBA modules from the destination workbook.
- Copies all VBA modules from the source workbook to the destination workbook.
- Saves and closes both workbooks automatically.
- Logs metadata:
  - ✅ Start time  
  - ✅ End time  
  - ✅ Time taken  
  - ✅ Username running the update  

---

## ⚙️ Setup Instructions
1. Insert this macro into a standard VBA module in your **controller workbook**.
2. Create a worksheet called **Dashboard** with:
   - Cell `C21` → Path to the **source workbook** containing the latest VBA modules.
3. Ensure you have supporting utility modules:
   - `RemoveModules` → Handles removal of all VBA modules in the destination workbook.
   - `CopyModules` → Handles copying of modules from source to destination.
4. Add named ranges (optional, for logging):
   - `[Status]`, `[Start_Time]`, `[Time_Taken]`, `[UserName]`.

---

## 🚀 Usage
1. In your controller workbook:
   - Enter the **source workbook path** in `Dashboard!C21`.
2. Run the macro:
   ```vba
   Call ControlSelectedWorkbookModulesUpdatingFromSourceWorkbook
   ```
3. When prompted, select the destination workbook (the one to update).
4. The macro will:
- Remove old VBA modules.
- Copy in the latest modules from the source workbook.
- Save and close the destination workbook.
5. A completion message will confirm: "The end of program"

## 📊 Example Workflow
- Source workbook: \\Server\Projects\MasterTemplate.xlsm
- Destination workbook (selected via dialog): FinanceReport_v3.xlsm
- Result:
  - All old modules in FinanceReport_v3.xlsm are removed.
  - Latest modules from MasterTemplate.xlsm are copied in.
  - Updated workbook saved automatically.

## 🛠 Customization
- Change .Filters.Add "Excel", "*.xlsm" if you need support for .xls or .xlsx.
- Modify logging ([Status], [Start_Time], etc.) to fit your existing system.
- Extend RemoveModules or CopyModules utilities if you want selective updates (e.g., only specific modules).

## 📄 Tested Environments
- Excel 2016, 2019, Microsoft 365
- Windows 10 / Windows 11

## ⚠️ Notes
- Destination workbook must be .xlsm to support VBA modules.
- Always backup your files before performing updates.
- Requires Trust access to the VBA project object model enabled in Excel (Options > Trust Center > Macro Settings).
