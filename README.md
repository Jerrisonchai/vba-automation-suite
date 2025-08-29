# VBA Automation Suite

## Overview
This repository contains a **collection of reusable VBA modules** designed to speed up daily business workflows in Excel and Outlook.  
The suite addresses common productivity needs such as:

- Automating **email attachment downloads** from Outlook.
- Handling **large Excel datasets** without lag.
- Performing **batch analysis and printing** of files.
- Integrating with **AI APIs** for quick Q&A inside Excel.
- Ensuring ISO compliance with signatures and project locks.
- Enforcing **version control & production deployment** for VBA projects.

> ⚠️ This suite is intended for **Windows Desktop Office** (Excel & Outlook). It may not work on Mac or Excel Online.

---

## Folder Structure

- **Download_attachments/**  
  Automates retrieval of email attachments from all messages inside a chosen Outlook subfolder. Useful for bulk reporting workflows.

- **hide-unhide-delete-hidden-column/**  
  Provides quick filter macros for large Excel workbooks.  
  Helps avoid performance lag by hiding/unhiding or cleaning hidden columns systematically.

- **Loopfiles-Analyse-Print/**  
  Loops through multiple files in a folder, performs predefined analysis, and prints/export results.  
  Ideal for batch processing of branch reports or daily logs.

- **testing-for-Gemini-API/**  
  Sample VBA integration with the **Google Gemini API**.  
  Lets you ask quick questions to the chatbot and return responses directly inside Excel.

- **Bulk_signature-ISO/**  
  - Inserts a standard ISO signature such as:  ```Template made by <Your Company>```
  across multiple templates in bulk.

- **Lock-or-unlock-VB-Project-ISO/**  
Provides a quick way to **lock or unlock VBA projects** for `.xlsm` templates.  
Ensures source code protection and compliance.

- **Push-updates-to-templates-ISO/**  
Implements a **lightweight version control system** for VBA templates.  
Allows pushing updated VBA modules to multiple workbooks/projects automatically.

- **PushProject2Production-ISO/**  
Standardizes deployment by moving all final VBA projects into a **Production folder**.  
Ensures consistent delivery and clean separation from drafts.

---

## Prerequisites

- **Microsoft Excel (2016/2019/2021/365)** – with macros enabled.
- **Microsoft Outlook (Desktop)** – required for email automation.
- **API Key (Gemini)** – needed for `testing-for-Gemini-API`.

---

## Installation & Setup

1. Clone or download this repo.  
2. Open any `.xlsm` file inside the module folders.  
3. Enable macros (`File > Options > Trust Center > Macro Settings`).  
4. For **Outlook automation**:
 - Grant access to programmatic access under Outlook Trust Center.
 - Ensure Outlook is running when scripts are executed.
5. For **Gemini API**:
 - Place your API key in the VBA script section:
   ```vba
   Const API_KEY As String = "YOUR_API_KEY_HERE"
   ```
 - Ensure you have internet access.

---

## Usage Scenarios

- **Finance/Reporting Teams** → Collect weekly sales attachments from Outlook, merge, and push updates to production templates.  
- **Data Analysts** → Hide/unhide columns dynamically to reduce Excel lag during heavy analysis.  
- **IT/ISO Auditors** → Enforce signature policies and lock projects for compliance.  
- **Automation Developers** → Test Gemini API responses inside Excel without switching to browser.  
- **Deployment Managers** → Use the `PushProject2Production` macro to ensure only verified builds move to production.

---

## Security & Compliance

- Macros should be stored in a **trusted folder** to avoid warnings.  
- Outlook macros may trigger security prompts — configure access responsibly.  
- Lock/unlock scripts should be used only by authorized maintainers.  
- For ISO features, always align with your company’s official compliance policy.

---

## Roadmap

- [ ] Add support for other AI APIs (e.g., OpenAI, Claude).  
- [ ] Extend Outlook automation to handle email body parsing.  
- [ ] Add interactive dashboards for template versioning.  
- [ ] Improve error handling and logging across all modules.  

---

## Status
✅ Stable for internal workflows  
⚠️ Some features (API integration, production push) require customization for your environment  

---

  
