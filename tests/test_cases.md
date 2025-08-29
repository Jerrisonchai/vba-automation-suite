# VBA Automation Suite – Test Cases

This document outlines the **functional test cases** for each module in the **VBA Automation Suite**. Test coverage ensures robustness, error handling, and predictable user experience.

---

## 1. Download Attachments
- **Case A**: No attachments in folder → Macro exits gracefully.  
- **Case B**: 100 small attachments → Saved in <15 sec.  
- **Case C**: Large attachments (≥10 MB) → No timeout.  
- **Case D**: Special characters in file names → Saved without corruption.  

---

## 2. Hide/Unhide/Delete Hidden Columns
- **Case A**: Sheet with 1000+ columns → Macro completes in ≤ 2 sec.  
- **Case B**: Mix of hidden + visible → Only intended columns are unhidden/deleted.  
- **Case C**: Empty sheet → Macro exits without error.  

---

## 3. Loop Files – Analyse & Print
- **Case A**: 50 Excel files → All processed and logged.  
- **Case B**: Corrupt file → Skipped, error logged.  
- **Case C**: Very large file (≥10 MB) → Macro completes without freezing.  

---

## 4. Gemini API Testing
- **Case A**: Valid API key → Response <2 sec.  
- **Case B**: Invalid API key → Clear error message.  
- **Case C**: No internet → Retry mechanism or graceful failure.  

---

## 5. Bulk Signature (ISO)
- **Case A**: 50 templates → All updated with signature.  
- **Case B**: Already signed template → No duplication.  
- **Case C**: Read-only file → Error logged, process continues.  

---

## 6. Lock/Unlock VB Project
- **Case A**: Lock 10 projects → All secured.  
- **Case B**: Unlock with correct password → Works correctly.  
- **Case C**: Unlock with wrong password → Error raised, project remains locked.  

---

## 7. Push Updates to Templates (ISO)
- **Case A**: Push to 50 templates → Runtime ≤ 2 min.  
- **Case B**: Missing template → Error logged, process continues.  
- **Case C**: Version mismatch → Logged without halting process.  

---

## 8. Push Project to Production (ISO)
- **Case A**: Deployment folder empty → Macro exits gracefully.  
- **Case B**: Push 20 projects → All appear in Production folder.  
- **Case C**: File locked by another process → Retry or log error.  

---

## 📌 Test Execution Notes
- Each test case must be **repeatable** and **logged**.  
- A **PASS/FAIL log sheet** should be maintained with timestamps.  
- Where applicable, **screenshots or log files** should be attached for evidence.  
