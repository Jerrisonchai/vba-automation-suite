# 📝 Bulk Signature Utility

This VBA utility automates the process of **applying standardized signatures** across multiple `.xlsm` templates in a folder.  
It ensures consistent attribution, version control, and compliance with **ISO documentation standards**.

---

## 🚀 Features
- Processes **all `.xlsm` files** in a selected folder.  
- Automatically creates a **backup folder** before modifying files.  
- Applies a signature with:
  - **Brand** (customizable via `Dashboard` sheet cell `C6`)  
  - **Repository Name** (`Dashboard!C8`)  
  - **Version Tag** (`Dashboard!C10`)  
  - **Timestamp**  
  - **Unique GUID** per file  
- Signature is written into:
  - **Custom document properties**  
  - **Defined Name** (`_JERR`)  
  - **Print footer** of all sheets  
- Built-in **logging & error handling** for transparency.  

---

## 📋 Example Signature
`
Template produced by Jerrison | Repo: VBA-Utility-Library/vba-file-folder-utils |
v2025-08-18 | SignedOn=2025-09-19 11:35 | GUID={xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}
`

---

---

## 🛠️ Usage
1. Open the **signing workbook** (`Signer.xlsm`) containing this macro.  
2. Make sure the workbook has a **Dashboard sheet** with:
   - `C6` → Brand  
   - `C8` → Repo  
   - `C10` → Version Tag  
3. Run `SignAllXlsmInFolder`.  
4. Select the target folder containing `.xlsm` templates.  
5. Review summary message box for processed, success, and error counts.  

---

## ⚡ Workflow
1. Pick a folder.  
2. Backup all `.xlsm` files to a `backup_yyyymmdd_HHmmss` subfolder.  
3. Open each file, apply signature, save, and close.  
4. Show summary report.  

---

## ⚠️ Notes & Requirements
- Macro requires **VBA project trust** enabled (optional: if extending signature to code modules).  
- Works with **Office Classic View** on Windows only.  
- Avoid running on shared drives with simultaneous access.  
- **Backup folder** is always created as a safeguard.  

---

## 📂 Example Folder Structure
`
C:\Projects\Templates
├── MyTemplate1.xlsm
├── MyTemplate2.xlsm
└── backup_20250919_113500
├── MyTemplate1.xlsm
└── MyTemplate2.xlsm
`
---


