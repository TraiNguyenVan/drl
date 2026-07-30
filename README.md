Automated pipeline for processing student training evaluation sheets (*Phiếu đánh giá kết quả rèn luyện*) for class **E25CQCE02-N**. Parses student self-evaluations, cross-references official scores, generates final Word documents for all 37 students, and compiles a detailed sub-section Excel report.

quick result can be found here on (Google Drive)[[https://docs.google.com/document/d/1EBcRCFQ5teoBeHXYPzOqpDBr-USL0J2B/edit?usp=sharing&ouid=109372832204939230581&rtpof=true&sd=true](https://drive.google.com/drive/folders/16J6rIboK8c9o4W_eTQlUI6b4l9qazNNB?usp=sharing)]

## 📋 Table of Contents
* [Overview](#-overview)
* [Codebase Structure](#-codebase-structure)
* [Document Architecture](#-document-architecture-master_v2docx)
* [Pipeline Workflow](#-pipeline-workflow)
* [Key Technical Details](#-key-technical-details)
* [Setup & Installation](#-setup--installation)
* [Usage](#-usage)

---

## 🔍 Overview

At the end of each semester, students submit training point evaluation documents (`.doc` or `.docx`). This pipeline automates:

1. **Batch-converts** `.doc` files to `.docx` via headless LibreOffice
2. **Parses & extracts** student name, DOB, and per-criteria scores with fuzzy matching (handles typos)
3. **Cross-references** parsed scores against the official ground-truth Excel summary
4. **Generates** final formatted Word documents for all students (present and absent)
5. **Compiles** a detailed sub-section Excel report with styling, borders, and signature blocks

---

## 📁 Codebase Structure

| File / Folder | Description |
|---|---|
| `master_v2.docx` | **Active template** — Google Docs-compatible, no floating textboxes |
| `master.docx` | Original template (preserved, used by `main` branch) |
| `HV_Mau 2_Tong hop KQRL cua SV.xlsx` | **Source of truth** — official final scores (do not modify) |
| `HV_Mau 2_Chi tiet KQRL.xlsx` | *(Generated)* Detailed sub-section report |
| `process_and_generate.py` | Main pipeline script |
| `STRUCTURE.md` | Document layout map (paragraph/table indices, cell mappings) |
| `GEMINI.md` | AI agent rules and constraints for this project |
| `students/` | Raw student papers (`.doc` / `.docx`) |
| `generated_students/` | *(Generated)* Final output Word files |

---

## 📐 Document Architecture (`master_v2.docx`)

Every generated file is a copy of `master_v2.docx` with student data filled in. The document body contains **5 paragraphs** and **4 tables** in the following order:

```
Para 0   — empty spacer
Para 1   — spaces (layout padding)
Table 0  — Logo & header block (institution name + national motto)
Para 2   — Title: "PHIẾU ĐÁNH GIÁ KẾT QUẢ RÈN LUYỆN"
Para 3   — Term & year: "Học kỳ: II      Năm học: 2025-2026"
Table 1  — Info table [2 rows × 2 cols, borderless]
Para 4   — empty spacer
Table 2  — Grading criteria matrix [56 rows × 12–15 cols]
Table 3  — Signature table [2 rows × 4 cols, borderless]
```

### Table 1 — Info Table (borderless)
| Cell | Content | Alignment |
|---|---|---|
| `rows[0].cells[0]` | `Họ và tên: <name>` | Left |
| `rows[0].cells[1]` | `Ngày sinh: <dob>` | Right |
| `rows[1].cells[0]` | `Mã số sinh viên: <msv>` | Left |
| `rows[1].cells[1]` | `Lớp: E25CQCE02-N` | Right |

### Table 2 — Grading Criteria Matrix
- 56 rows (indices 0–55)
- Score columns (constant regardless of horizontal merges):
  - `cells[7]` — Student self-score
  - `cells[8]` — Class score
  - `cells[11]` — Advisor score
- Key rows: `20, 30, 37, 45, 52` = section totals; `53` = grand total

### Table 3 — Signature Table (borderless)
| Col | Content | Source |
|---|---|---|
| `rows[1].cells[0]` | `Nguyễn Trung Hiếu` | Fixed in template |
| `rows[1].cells[1]` | `Ngô Trí Long` | Fixed in template |
| `rows[1].cells[2]` | `Nguyễn Phương Quốc Vương` | Fixed in template |
| `rows[1].cells[3]` | `<student name>` | Written by pipeline |

---

## ⚙️ Pipeline Workflow

### Step 1 — Conversion
Scans `students/` and batch-converts `.doc` → `.docx` using `soffice --headless`.

### Step 2 — Database load
Reads `HV_Mau 2_Tong hop KQRL cua SV.xlsx`:
- Student roster: TT index, full name, MSV, DOB, section scores (cols E–J), classification (col K)

### Step 3 — Per-student processing

**Present students** (file found):
1. Parse DOB from student's raw file paragraph text
2. Extract scores from `doc_student.tables[1]` (grading table in raw student file, which still uses the old template structure)
3. Fuzzy-match criteria descriptions (ratio ≥ 0.8) to handle typos
4. Copy `master_v2.docx` → `generated_students/<name>_<msv>.docx`
5. Fill **Table 1** (info): name, DOB, MSV
6. Fill **Table 2** (grading): overwrite score cells + section totals + grand total, all centered
7. Fill **Table 3** (signature): write student name to `rows[1].cells[3]`

**Absent students** (no file):
1. Copy `master_v2.docx` → `generated_students/<name>_<msv>.docx`
2. Fill **Table 1**: name and MSV (DOB left as `Ngày sinh:`)
3. Set all score cells in **Table 2** to `0`
4. Fill **Table 3**: write student name to `rows[1].cells[3]`

### Step 4 — Excel report
Compiles all 22 sub-sections into `HV_Mau 2_Chi tiet KQRL.xlsx` with borders, color fills, and footer signatures.

---

## ✨ Key Technical Details

### Centered scores
`write_centered_score(cell, text)` writes text and explicitly sets `alignment = WD_ALIGN_PARAGRAPH.CENTER`. Overwriting `.text` resets alignment to `None` in python-docx, so every score write uses this helper.

### Google Docs compatibility (`google-docs-compat` branch)
The `main` branch uses `master.docx` with floating textboxes (`wp:anchor`) for signatures. These render inconsistently in Google Docs (ghost duplicates, page splits).

`master_v2.docx` (this branch) resolves this by:
- **Removing all floating textboxes** — no `wp:anchor` / VML `w:pict` shapes
- **Info section** — replaced tab-stop paragraphs with a borderless 2×2 table (tabs render differently per app)
- **Signature section** — replaced absolute-position floats with a borderless 2×4 table

### Absent student detection
Students `Nguyễn Trương Minh Triết (N25DECE071)` and `Vương Hà Hải Đăng (N25DECE078)` have no files in `students/`. They are auto-detected and given blank templates with all scores set to `0`.

---

## 🛠️ Setup & Installation

Requires Python 3 and LibreOffice.

### System dependencies (Debian/Ubuntu)
```bash
sudo apt update && sudo apt install libreoffice
```

### Python packages
```bash
pip3 install python-docx openpyxl lxml
```

---

## 🚀 Usage

```bash
python3 process_and_generate.py
```

Expected output:
```
Converting and preparing raw student files...
Preparation completed.

Processing student files...
  TT=22: Nguyễn Trương Minh Triết (N25DECE071) -> MISSING FILE (Absent)
  TT=29: Vương Hà Hải Đăng (N25DECE078) -> MISSING FILE (Absent)
Successfully generated 37 student Word files.

Generating detailed Excel report...
Successfully generated detailed Excel report at /home/trai/workspace/drl/HV_Mau 2_Chi tiet KQRL.xlsx
Temporary files cleaned up.
Pipeline run finished successfully!
```
