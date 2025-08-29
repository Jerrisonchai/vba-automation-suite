# User Guide – VBA Automation Suite

## Introduction
This suite provides ready-to-use macros for common business automation tasks.  
No programming knowledge required — just open the workbook, enable macros, and run.

---

## Prerequisites

- Microsoft Excel (2016/2019/2021/365) with macros enabled.
- Microsoft Outlook (for email automation features).
- Access rights to modify templates (for ISO modules).
- API key (for Gemini chatbot integration).

---

## How to Use

### 1. Download Attachments
- Open `Download_attachments.xlsm`
- Click **Run Macro**
- Select the Outlook subfolder → attachments will be saved into a local folder.

### 2. Hide/Unhide/Delete Hidden Columns
- Open `hide-unhide-delete-hidden-column.xlsm`
- Run the macro → unnecessary columns will hide/unhide to improve performance.

### 3. Loop Files – Analyse & Print
- Place your files in the input folder.
- Open the macro workbook.
- Run → analysis and print/export will run for each file.

### 4. Testing Gemini API
- Add your Gemini API key inside the workbook.
- Run the macro → ask a question and get an AI response in Excel.

### 5. Bulk Signature (ISO)
- Open the `Bulk_signature-ISO.xlsm`.
- Run → signature text `"Template made by..."` is added to all templates in the folder.

### 6. Lock/Unlock VB Project (ISO)
- Open workbook.
- Run → lock/unlock project VBA password automatically.

### 7. Push Updates to Templates (ISO)
- Place your updated VBA project in the source folder.
- Run macro → updates are pushed to all target templates.

### 8. Push Project to Production (ISO)
- Select finalized projects.
- Run macro → copies them to the `Production` folder.

---

## Error Messages

- **“Macros Disabled”** → Enable content under yellow warning bar.
- **“Outlook not running”** → Start Outlook first.
- **“API Key missing”** → Enter your Gemini API key.

---

## Tips

- Always **backup** your templates before running bulk macros.
- Use the **Status sheet** to view start time, runtime, and user name logs.
- For ISO tasks, confirm your company’s official compliance text before pushing signatures.

