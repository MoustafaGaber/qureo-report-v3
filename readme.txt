# 📊 QUREO Student Progress Report Tool

A fast, lightweight desktop utility for Windows to extract student progress reports from the **QUREO** educational platform (completed chapters, completion percentage out of 42 chapters, total diamonds, and last activity timestamp in Cairo time) and export the results to an Excel-ready CSV file.

---

## 1️⃣ Preparing the Students File (`students.csv`)

Before running the tool, prepare your student list in **Microsoft Excel** (or Google Sheets) with the following column headers in the first row:

| name | class | student_id | password |
| :--- | :--- | :--- | :--- |
| Ahmed Mohamed Ali | 1 | 475000001 | Pass@1234 |
| Sara Mahmoud Hassan | 2 | 475000002 | Pass!5678 |

*(Note: Arabic headers `الاسم` and `الفصل` are also supported for the first two columns).*

### Raw CSV Format (`students.csv`):
```csv
name,class,student_id,password
Ahmed Mohamed Ali,1,475000001,Pass@1234
Sara Mahmoud Hassan,2,475000002,Pass!5678