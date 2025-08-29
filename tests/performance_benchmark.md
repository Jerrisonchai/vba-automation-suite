# VBA Automation Suite – Performance Benchmark

This document outlines the performance targets and measurement standards for the automation modules within the **VBA Automation Suite**. Benchmarks ensure scalability, efficiency, and reliability of the solution across different environments.

---

## 📈 Benchmark Summary

| Module                                   | Key Metric                       | Target Performance             | Measurement Method                  |
|------------------------------------------|----------------------------------|--------------------------------|--------------------------------------|
| **Download_attachments**                 | Avg. time per 100 attachments    | ≤ 15 sec                       | Compare `Start_Time` vs `End_Time`   |
| **hide-unhide-delete-hidden-column**     | Column operation speed (1k cols) | ≤ 2 sec                        | `Timer` before/after loop execution |
| **Loopfiles-Analyse-Print**              | Processing 100 files             | ≤ 60 sec                       | Compare runtime against file count   |
| **testing-for-Gemini-API**               | Response round-trip              | ≤ 2 sec (local environment)    | `Timer` around API call              |
| **Bulk_signature-ISO**                   | Templates updated/minute         | ≥ 20 templates/min             | Count updated vs runtime             |
| **Lock/Unlock VB Project**               | Avg. time to lock/unlock project | ≤ 1 sec/project                | Batch lock/unlock test               |
| **Push-updates-to-templates-ISO**        | Sync speed (50 templates)        | ≤ 2 min                        | Runtime logging per batch            |
| **PushProject2Production-ISO**           | Deployment success rate          | 100%                           | Verify checksum of deployed files    |

---

## 🛠 Benchmark Methodology
1. **Timer Functions** – Use VBA `Timer` or `Now()` to measure start vs end runtime.
2. **Batch Testing** – Execute macros on sample sets (50–100 files, 1000+ columns).
3. **Stress Testing** – Include large file sizes (≥10 MB), corrupt files, and read-only files.
4. **Error Logging** – Capture failure points without halting execution.
5. **Scalability Testing** – Run macros on increasing workloads to validate consistency.

---

## 📌 Notes
- Benchmarks may vary depending on **system specifications**, **network speed**, and **Outlook/Excel version**.
- Results should be logged in a central worksheet or `.log` file for transparency.

