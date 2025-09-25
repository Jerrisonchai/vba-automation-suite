# Excel Integration with Gemini API

This project integrates **Google Gemini API** with Excel using VBA, allowing users to generate **AI-powered summaries** of their daily tasks directly within a workbook.

---

## 📌 Features
- Connects Excel to **Gemini 1.5 Flash** via HTTP request.
- Reads an API key stored in a worksheet (`Reference!C5`).
- Logs tasks from the `Tasks` sheet and generates AI-based **summaries/insights**.
- Flexible VBA function (`GetAISummary_Gemini`) for re-use in other macros.

---

## ⚡ Quick Setup

1. **Get Gemini API Key**  
   - Sign in to [Google AI Studio](https://ai.google.dev/).  
   - Navigate to **Get API Key**.  
   - Copy the key and paste it into your workbook at:  
     `Reference!C5`

2. **Download JSON Converter for VBA**  
   - Go to the [VBA JSON GitHub Repo](https://github.com/VBA-tools/VBA-JSON).  
   - Download `JsonConverter.bas`.  
   - In Excel VBA Editor:  
     - `File > Import File > JsonConverter.bas`  
     - Enable **Microsoft Scripting Runtime** under `Tools > References`.  

---

## ⚙️ Workbook Preparation
1. Add a worksheet named **`Reference`**:
   - In cell **C5**, paste your Gemini API key.
2. Add a worksheet named **`Tasks`**:
   - Column `G` → Daily tasks text.
   - Column `H` → AI-generated insights (filled by macro).

---

## 🚀 Usage

### 1. Generate AI Summary
Call the function directly in VBA:
```vba
MsgBox GetAISummary_Gemini("Summarize today's top 3 priorities.")
```
### 2. Log Daily Tasks
Run the macro: "Call LogDailyTasks"
- Takes the last task entry from Tasks!G:G.
- Sends to Gemini for summarization.
- Outputs AI summary into Tasks!H:H.

## 📊 Example Workflow
- Enter tasks in Tasks!G2:G10.
- Run LogDailyTasks.
- AI-generated summary will appear in the corresponding row of column H.

## ⚡ Performance Benchmark
| Test Case             | Avg Response Time | Notes                           |
| --------------------- | ----------------- | ------------------------------- |
| 1 short task (1 line) | \~2s              | Near instant                    |
| 5 tasks (1 paragraph) | \~3–4s            | Works smoothly                  |
| 10+ tasks (long list) | \~6–7s            | JSON parsing overhead increases |

## ✅ Test Cases
| Scenario                | Expected Outcome                       |
| ----------------------- | -------------------------------------- |
| Empty `Tasks!G` cell    | AI returns error message in column `H` |
| Valid API key + 3 tasks | Summary written to column `H`          |
| Invalid API key         | Error message logged in column `H`     |
| Multiple task entries   | Only the **latest row** is summarized  |

## 🛠 Troubleshooting
- "Gemini API error" → Check API key in Reference!C5.
- JSON Parse Error → Ensure JsonConverter.bas is properly imported.
- Slow responses → Minimize task text length; Gemini response time scales with input size.
