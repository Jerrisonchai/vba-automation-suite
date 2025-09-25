# VBA Macro – SaveAs (Batch Export to Production)

## 📌 Overview
The **`SaveAs`** macro automates the process of exporting one or more workbooks from a selected folder into a **production-ready format**.  

It performs the following tasks:
1. Prompts the user to select a folder containing `.xlsm` workbooks.  
2. Saves each workbook into a **production folder** (defined in `Dashboard!C21`), renaming files as needed.  
3. Updates and reassigns all external links in formulas.  
4. Logs execution details such as time taken, user, and status.  
5. Notifies the user once the process is complete.  

This macro is designed to **streamline template deployment** (e.g., migrating from *Beta* to *Production* workbooks).

---

## ✨ Features
- Folder picker for selecting multiple files.  
- Automatic **renaming**: replaces `T.xlsm` with `.xlsm` during save.  
- Converts workbooks into **macro-enabled format (`.xlsm`, FileFormat:=52)**.  
- Updates **all Excel links** in formulas to point to the newly saved workbook.  
- Writes log data into named ranges:
  - `[Status]` → "Success"  
  - `[Start_Time]` → Process start time  
  - `[Time_Taken]` → Duration in HH:MM:SS format  
  - `[UserName]` → Current Windows user  

---

## ⚙️ Setup Instructions
1. Insert the macro into a standard VBA module in your **controller workbook**.  
2. Ensure you have a **Dashboard sheet** with:
   - `C20` → Default folder path (for initial folder picker).  
   - `C21` → Path to production/export folder.  
3. Add named ranges `[Status]`, `[Start_Time]`, `[Time_Taken]`, `[UserName]` to capture log results.  
4. Ensure utility functions are available:
   - `capturetime` / `captureendtime` → For logging execution time.  
   - `OptimizedMode` → For toggling Excel performance optimization.  
   - `MyShape_Click` → (Optional) Triggered if UI interaction is required.  

---

## 🚀 Usage
1. In your **Dashboard sheet**:  
   - Set `C20` = Default Beta folder.  
   - Set `C21` = Production/export folder.  

2. Run the macro:
   ```vba
   Call SaveAs
   ```
3. When prompted:
- Select the folder containing the .xlsm Beta templates.
4. The macro will:
- Loop through all .xlsm files in the chosen folder.
- Save each into the Production folder with updated names.
- Update all formula links to point to the new workbook.
- Log execution metadata.
5. A message box will confirm completion: "Beta template exported to Production"

## 📊 Example Workflow
- Dashboard!C20: D:\BetaTemplates
- Dashboard!C21: D:\ProductionTemplates
- Beta file: FinanceReportT.xlsm
- Result file: FinanceReport.xlsm in the Production folder.
- All formula links inside the workbook are updated to match the new production file.

## 🛠 Customization
- Modify file renaming logic:
  ```vba
  wbkSN = WorksheetFunction.Substitute(wbkS.Name, "T.xlsm", ".xlsm")
  ```
  - → Adjust if your Beta/Prod naming pattern differs.
- Add additional file formats by changing: "FileFormat:=52"
  - 51 = .xlsx
  - 52 = .xlsm
- Extend logging to a separate log sheet for auditing.

## 📄 Tested Environments
- Excel 2016, 2019, Microsoft 365
- Windows 10 / Windows 11

## ⚠️ Notes
- Ensure Trust access to the VBA project object model is enabled in Excel (Options > Trust Center > Macro Settings).
- Files in the production folder will be overwritten if names match.
- Always backup your Beta and Production folders before running the macro.
