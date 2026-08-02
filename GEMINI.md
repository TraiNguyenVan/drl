# Project Rules & Constraints (GEMINI.md)

## 🚫 Critical Constraints
* **Source of Truth**: You will not modify [HV_Mau 2_Tong hop KQRL cua SV.xlsx](HV_Mau%202_Tong%20hop%20KQRL%20cua%20SV.xlsx), scores here always true.

## 📄 Template & Document Management (`master.docx`)
* [master.docx](master.docx) is the active template — it uses paragraph-based info lines with a custom tab stop at 3.8 inches for names and DOBs, and a shape textbox for the student name signature. Read [STRUCTURE.md](STRUCTURE.md) to understand the structure of the template.
* Whenever I need to create more files like `master.docx`, you must make a copy of the active template ([master.docx](master.docx)), edit right on it to preserve intended format.
* `./students` is where raw student papers live, they can be doc or docx, and may not have the right format as the template, but we may extract scores from here whenever we need to generate new files.
* Table and paragraph indices in the generated files (`master.docx`):
  * `paragraphs[4]` = Name/DOB
  * `paragraphs[5]` = MSV/Class (both set with a 3.8" custom tab stop)
  * `tables[0]` = Logo
  * `tables[1]` = Grading criteria & signature headers (56 rows)

## ⚙️ Pipeline Execution
* The pipeline script is [process_and_generate.py](process_and_generate.py) — run it with `python3 process_and_generate.py` to regenerate all 37 student files, the static charts, and the HTML dashboard.

## 🔢 Scoring & Step-Increment Constraints
Rules for scoring and step-increment constraints (to avoid mathematical inconsistencies in the CSV and Word files):
* **1.2 (GPA classification)**: can only take values from {0, 4, 6, 8, 10} corresponding to classification tiers. Intermediate values (e.g. 7 or 9) are invalid.
* **1.4 (Extracurricular activities)**: must be multiples of 0.5 points (max 2).
* **2.3 (Job seminars)**: must be integers (1 point/seminar, max 5).
* **3.1 (Political & social activities)**: must be multiples of 2 points (2 points/activity, max 10). Do not use odd numbers (like 5, 7, 9).
* **3.2 (Social work/charity)**: must be integers (1 point/activity, max 4).
* **3.3 (Social media school promotion)**: must be integers (1 point/activity, max 3).
* **5.1 (Class officer tasks)**: must be integers (usually 0 or 4).
* **5.2 (Club participation)**: must be integers (max 3).
* **TC1-TC5 & TOTAL**: Category summary rows (TC1-TC5) and TOTAL rows must always mathematically equal the sum of their respective sub-criteria and category rows.

## 📓 Lessons Learned & Mistakes Log
* **Instruction**: Whenever user report/question/offend you with your stupid mistake, you must ask if the user want you to take notes here (be short, precise and only with something real important with this project, chores and something like that just ignore it).

### Log: