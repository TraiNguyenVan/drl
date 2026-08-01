Automated pipeline for processing student training evaluation sheets (*Phiếu đánh giá kết quả rèn luyện*) for class **E25CQCE02-N**. Parses student self-evaluations, cross-references official scores, generates final Word documents for all 37 students, and compiles a detailed sub-section Excel report.

quick result can be found here on [Google Drive](https://drive.google.com/drive/folders/16J6rIboK8c9o4W_eTQlUI6b4l9qazNNB?usp=sharing)

## 📋 Table of Contents
* [Overview](#-overview)
* [Codebase Structure](#-codebase-structure)
* [Document Architecture](#-document-architecture-masterdocx)
* [Pipeline Workflow](#-pipeline-workflow)
* [Key Technical Details](#-key-technical-details)
* [Setup & Installation](#-setup--installation)
* [Usage](#-usage)

---

## 🔍 Overview

At the end of each semester, students submit training point evaluation documents (`.doc` or `.docx`). This pipeline automates:

1. **Loads scores & DOB** directly from `ai_studio_code.csv` and the official roster summary.
2. **Generates** final formatted Word documents for all 37 students (present and absent) using custom tab stops for info alignment.
3. **Renders** statistics charts and compiles an interactive performance HTML dashboard.

---

## 📁 Codebase Structure

| File / Folder | Description |
|---|---|
| `master.docx` | **Active template** — Google Docs-compatible, uses paragraph-based info block and shape-wrapped textbox for name |
| `HV_Mau 2_Tong hop KQRL cua SV.xlsx` | **Source of truth** — official final scores (do not modify) |
| `HV_Mau 2_Chi tiet KQRL.xlsx` | *(Legacy)* Detailed sub-section report (no longer generated) |
| `process_and_generate.py` | Main pipeline script |
| `render_charts.py` | Module for generating static DRL chart images |
| `render_html.py` | Module for generating the interactive DRL HTML dashboard |
| `drl_dashboard.html` | *(Generated)* Interactive performance analytics dashboard |
| `charts/` | *(Generated)* Folder containing static DRL statistic charts |
| `STRUCTURE.md` | Document layout map (paragraph/table indices, cell mappings) |
| `GEMINI.md` | AI agent rules and constraints for this project |
| `students/` | Raw student papers (`.doc` / `.docx`) |
| `generated_students/` | *(Generated)* Final output Word files |

---

## 📐 Document Architecture (`master.docx`)

Every generated file is a copy of `master.docx` with student data filled in. The document body contains **12 paragraphs** and **2 tables** in the following order:

```
Para 0   — empty spacer
Para 1   — spaces (layout padding)
Table 0  — Logo & header block [1 row × 2 cols, borderless]
Para 2   — Title: "PHIẾU ĐÁNH GIÁ KẾT QUẢ RÈN LUYỆN"
Para 3   — Term & year: "Học kỳ: II      Năm học: 2025-2026"
Para 4   — Student Name & DOB: "Họ và tên: <Name>\tNgày sinh: <DOB>" (custom tab stop at 3.8")
Para 5   — Student ID & Class: "Mã số sinh viên: <MSV>\tLớp: E25CQCE02-N" (custom tab stop at 3.8")
Para 6-10— empty spacers
Table 1  — Grading criteria matrix & signature headers [56 rows × 12 cols]
Para 11  — Static class monitor name / spacing
```

### Student Info Alignment (Paragraphs 4 & 5)
Aligned using a custom **Left-Aligned Tab Stop** at **3.8 inches (9.65 cm)** and exactly one tab character (`\t`) to ensure alignment on Word/Google Docs/LibreOffice without tables.

### Table 1 — Grading Criteria Matrix & Signature Headers
- 56 rows (indices 0–55)
- Score columns:
  - `cells[7]` — Student self-score
  - `cells[8]` — Class score
  - `cells[11]` — Advisor score
- Key rows: `20, 30, 37, 45, 52` = section totals; `53` = grand total
- Row 55: Signature headers (Cố vấn học tập, Lớp trưởng, Bí thư, Sinh viên)

### Student Name Signature Textbox
Uses a compatibility-wrapped shape textbox in drawing container (`##STUDENT` / `_NAME##` placeholders) to place the student's name neatly below the signature headers.

---

## ⚙️ Pipeline Workflow

### Step 1 — Database and Score Loading
- Reads the official student roster and final totals from `HV_Mau 2_Tong hop KQRL cua SV.xlsx`.
- Loads specific sub-criterion scores and student Date of Birth (DOB) data from `ai_studio_code.csv`.

### Step 2 — Per-student processing
1. Make a copy of `master.docx`.
2. Clear template info runs and insert student name, MSV, and DOB using custom tab stops.
3. Write subscores from the CSV data into the grading criteria table (Table 1) columns (Student, Class, Advisor).
4. Overwrite section totals and grand totals (Rows 20, 30, 37, 45, 52, 53) using values from the master Excel file to ensure 100% mathematical consistency.
5. Split the student name into the signature textbox (`##STUDENT` / `_NAME##`).
6. Save the final document in `generated_students/`.

### Step 3 — Charts & Dashboard Generation
- Unless `--disable-charts` is passed, runs `render_charts.py` to generate statistical charts (score distribution, rating breakdown, criteria averages) in `charts/`.
- Runs `render_html.py` to generate the interactive, highly optimized `drl_dashboard.html`.

---

## ✨ Key Technical Details

### Tab Stop Alignment
Info columns are perfectly aligned at `3.8 inches` from the margin by clearing default paragraphs and programmatically injecting custom tab stop properties to `doc.paragraphs[4]` and `doc.paragraphs[5]`.

### Excel Ground-Truth Total Overwrite
Totals and Grand Totals are overwritten from the master Excel summary file, guaranteeing correct final totals across all files.

---

## 🛠️ Setup & Installation

Requires Python 3.

### Python packages
```bash
pip3 install python-docx openpyxl lxml pillow matplotlib numpy
```

---

## 🚀 Usage

Run the complete pipeline:
```bash
python3 process_and_generate.py
```

### Command Line Flags
You can customize the pipeline run using the following flags:
* `--disable-charts`: Skip generating statistic charts and the interactive HTML dashboard.
* `--disable-gdrive`: Skip syncing the generated student files to Google Drive.
* `--disable-excel`: Deprecated alias for `--disable-charts`.

**Example:**
```bash
python3 process_and_generate.py --disable-charts --disable-gdrive
```
