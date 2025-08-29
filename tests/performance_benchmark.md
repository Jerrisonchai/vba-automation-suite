| Module                               | Key Metric                          | Benchmark Target   | Measurement Method                             |
| ------------------------------------ | ----------------------------------- | ------------------ | ---------------------------------------------- |
| **Download\_attachments**            | Avg. time per 100 attachments       | ≤ 15 sec           | Measure `Start_Time` vs `End_Time` (log sheet) |
| **hide-unhide-delete-hidden-column** | Column operation speed (1k columns) | ≤ 2 sec            | `Timer` before/after loop                      |
| **Loopfiles-Analyse-Print**          | Processing 100 files                | ≤ 60 sec           | Compare total runtime against file count       |
| **testing-for-Gemini-API**           | Response round-trip                 | ≤ 2 sec (local)    | `Timer` around API call                        |
| **Bulk\_signature-ISO**              | Templates updated/minute            | ≥ 20 templates/min | Count updated vs runtime                       |
| **Lock/Unlock VB Project**           | Avg. time to lock/unlock            | ≤ 1 sec/project    | Batch lock across 10 files                     |
| **Push-updates-to-templates-ISO**    | Sync speed (50 templates)           | ≤ 2 min            | Log completed templates                        |
| **PushProject2Production-ISO**       | Deployment success rate             | 100%               | Verify checksum of deployed files              |
