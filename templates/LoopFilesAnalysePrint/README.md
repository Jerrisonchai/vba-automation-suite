# VBA Macro – Print Analytics of Files

## 📌 Overview
The **`PrintAnalyticsofFiles`** macro automates the process of:
1. Looping through all Excel files (`*.xls`) in a specified folder.
2. Extracting raw data from each file.
3. Feeding the data into a pre-defined **Print** worksheet.
4. Automatically printing the **UsedRange** (configured to print as PDF or physical printer).

This is especially useful for **batch-printing reports** or **generating multiple PDFs** from a collection of Excel files without manual effort.

---

## ✨ Features
- Loops through all `.xls` files in a target folder.
- Clears and reloads raw data dynamically for each file.
- Prints the **UsedRange** from a formatted `Print` sheet.
- Automatically records:
  - ✅ **Start time**
  - ✅ **End time**
  - ✅ **Execution duration**
  - ✅ **User running the macro**
- Integrated with `Dashboard` sheet for file path input and logging.

---

## ⚙️ Setup Instructions
1. Open the VBA editor (`Alt + F11`).
2. Insert the macro into a standard module.
3. Prepare your workbook with the following worksheets:
   - **Dashboard**
     - `C16` → Folder path (where source files are located).
     - `C10` → Displays current file being processed.
   - **RawData**
     - Temporary data dump from each file.
   - **Print**
     - Predefined layout that recalculates and prints results.
   - (Optional) Named ranges `[Status]`, `[Start_Time]`, `[Time_Taken]`, `[UserName]` for logging.
4. Set your default printer to **Microsoft Print to PDF** if you want PDFs.

---

## 🚀 Usage
1. Enter the **folder path** containing `.xls` files in `Dashboard!C16`.
2. Run the macro:
   ```vba
   Call PrintAnalyticsofFiles
   ```
3. The macro will:
- Process each file in sequence.
- Print reports based on the Print sheet.
- Log execution details.

## Example Workflow
- Folder contains: Sales_Jan.xls, Sales_Feb.xls, Sales_Mar.xls
- For each file:
  - Data copied into RawData.
  - Print sheet recalculates.
  - Report is printed.
- Loop continues until all files are processed.
- Final message box confirms completion: "Print done"

## 🛠 Customization
- Change Range("A:AZ") to adjust which columns of data are imported.
- Update .UsedRange.PrintOut if you want:
  - Preview instead of print.
  - Export to PDF via ExportAsFixedFormat.
- Uncomment the Application settings (DisplayAlerts, AskToUpdateLinks, OptimizedMode) to further speed up execution.

## 📄 Tested Environments
- Excel 2016, 2019, Microsoft 365
- Windows 10 / Windows 11

## ⚠️ Notes
- Only supports .xls by default. Change Dir(folderPath & "*.xls") to *.xlsx or *.xlsm if needed.
- Ensure your Print sheet is properly formatted before running the macro.
- Recommended to back up your files before batch processing.
