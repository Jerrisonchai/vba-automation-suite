# 📂 Download_Attachments – VBA Automation

This module automates the process of downloading email attachments from **Microsoft Outlook Desktop (Classic View)** into a designated folder on your local machine.  
It supports:
- Saving attachments from specific mailbox subfolders
- Handling nested `.msg` attachments
- Extracting and opening links embedded in email body text

---

## 🚀 Features
- Bulk download of attachments from multiple emails in a chosen Outlook subfolder.
- Auto-cleanup of non-Excel files if required.
- Regex-powered extraction of hyperlinks inside email body.
- Configurable mailbox and folder paths via Excel dashboard ranges or named cells.

---

## 📂 Folder Path Structure

The VBA code follows this hierarchy to reach attachments:
```Outlook App → Mailbox → Inbox → Subfolders → Target Folder → Items → Attachments```
Example (as coded in the sample):
```Inbox → Purchasing Project → from Supplier```


---

## ⚙️ Setup & Customization

### 1. Mailbox Name
- Update the `Mailbox_Name` in your Excel sheet or dashboard (linked via `[Mailbox_Name].Text`).  
Example: ```Your.Name@company.com```

### 2. Subfolder Path
- Adjust these folder names in the VBA module if your hierarchy differs:
```vba
Set olFolder = olNS.Folders([Mailbox_Name].Text)
Set olFolder = olFolder.Folders("Inbox")
Set olFolder = olFolder.Folders("Purchasing Project")
Set olFolder = olFolder.Folders("from Supplier")
```
Replace "Purchasing Project" and "from Supplier" with your Outlook subfolder names.

### 3. Export Path
- Attachments are saved to the folder specified in [Export_To].Text.
Example:```C:\Users\<username>\Documents\Attachment_Downloads```

### 4. File Handling
- By default, .xls/.xlsx files are preserved.
- Other file types may be deleted after processing (optional cleanup logic).

## 🖥 Usage Instructions
1. Ensure Outlook Desktop (Classic View) is open and logged in.
2. Configure Mailbox_Name, target subfolders, and Export_To path in the dashboard sheet.
3. Run the macro:
- download_attachments → Downloads all attachments from the defined folder.
- SaveOlAttachments → Handles .msg attachments (attachments inside attachments).
- OpenLinksMessage → Extracts and opens hyperlinks in Internet Explorer.
- clicklinks → Opens specific catalog links in Chrome.

## 🔧 Example Workflow
1. Place supplier emails into Inbox → Purchasing Project → from Supplier.
2. Set export folder: ```C:\Users\Jerr\Documents\Supplier_Files```
3. Run download_attachments.
4. All Excel attachments are now available in your export folder.

## 🧩 Notes
- This works only with Outlook Desktop (Classic View). It does not support web or mobile clients.
- Ensure Outlook security prompts are handled (some environments may require enabling programmatic access).
- Internet Explorer is deprecated — update OpenLinksMessage to use Chrome/Edge if needed.

## ✅ Example Customization
- If your emails are under: ```Inbox → Finance → Monthly Reports```
- Update the VBA path:
```
Set olFolder = olNS.Folders([Mailbox_Name].Text)
Set olFolder = olFolder.Folders("Inbox")
Set olFolder = olFolder.Folders("Finance")
Set olFolder = olFolder.Folders("Monthly Reports")
```
- Export path in Excel [Export_To].Text: ```D:\Shared\Reports\2025```

## 📌 Status Logging
Each run updates:
- [Status] → "Success" or "Failed"
- [Start_Time] → Timestamp of execution
- [Time_Taken] → Duration (HH:MM:SS)
- [User_Name] → Current Windows user

## 🔒 Error Handling
- Missing folder → exits gracefully.
- Corrupt files → skipped but logged.
- Wrong mailbox name → macro stops with a clear message.

