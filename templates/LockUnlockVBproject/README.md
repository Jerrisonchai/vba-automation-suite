# VBA Project Lock/Unlock Utility

## 📌 Overview
This VBA utility automates the process of locking and unlocking the **VBA Project Properties** in Excel workbooks.  
Normally, setting or removing a VBA project password requires manual interaction with dialog boxes.  
This script leverages Windows API calls (`user32.dll`, `kernel32.dll`) to programmatically handle the dialogs, reducing repetitive manual work.

---

## ✨ Features
- **Lock VBA Project** with a password.
- **Unlock VBA Project** using the same password.
- Compatible with both:
  - **VBA7 (64-bit Office)**
  - **Legacy VBA (32-bit Office)**
- Automatically saves the workbook after locking/unlocking.
- Error-handling with cleanup of Windows hooks.

---

## ⚠️ Security Notice
This utility manipulates password dialogs via Windows API calls.  
- Use **only on your own projects**.  
- Do **not** use this to bypass password protection on files you do not own.  
- Store your password securely.  

---

## 🔧 Setup
1. Open the VBA Editor (`Alt + F11`).
2. Insert a **new module** and paste the code.
3. Ensure references to **Excel Object Library** and **VBA Extensibility Library** are enabled:
   - In VBA Editor → `Tools > References`  
   - Check:
     - ✅ **Microsoft Excel XX.0 Object Library**  
     - ✅ **Microsoft Visual Basic for Applications Extensibility 5.3**  
4. Adjust the **password string** (default is `@dm1n`) to your preferred password.

---

## 🚀 Usage

### Lock Project
```vba
Sub Lock_Example()
    'Locks the current workbook VBA Project with the given password
    LockVBProject(WorkbookName:=ThisWorkbook.Name, Password:="@dm1n") = True
End Sub
```
### Unlock Project
```vba
Sub UnLock_Example()
    'Unlocks the current workbook VBA Project with the given password
    LockVBProject(WorkbookName:=ThisWorkbook.Name, Password:="@dm1n") = False
End Sub
```

## 🛠 Example Workflow

- Run Lock_Example() → VBA Project is locked with password.
- Run UnLock_Example() → VBA Project is unlocked.
- Workbook saves automatically after changes.

## 📊 Tested Environments

- Excel 2016, 2019, 2021, and Microsoft 365
- Windows 10 / Windows 11
- Both 32-bit and 64-bit VBA



