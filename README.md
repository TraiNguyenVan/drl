# Student Training Point (Kết quả rèn luyện - KQRL) Processing Pipeline

A robust, automated pipeline for processing student training point papers (Kết quả rèn luyện) for the class **E25CQCE02-N**. The system parses student self-evaluations, computes subsections and totals, matches scores with the official database, generates clean aligned Word documents for all 37 students, and compiles a comprehensive sub-section report spreadsheet.

## 📋 Table of Contents
* [Overview](#-overview)
* [Codebase Structure](#-codebase-structure)
* [Pipeline Workflow (`process_and_generate.py`)](#-pipeline-workflow-process_and_generatepy)
* [Key Features & Technical Implementations](#-key-features--technical-implementations)
* [Setup & Installation](#-setup--installation)
* [Usage](#-usage)

---

## 🔍 Overview

At the end of each semester, students submit their training point evaluation documents (either as `.doc` or `.docx`). The manual compilation of these scores into a unified report and the generation of formatted final papers is tedious and error-prone. 

This pipeline automates the entire process:
1. **Converts & Normalizes**: Automatically batch-converts old `.doc` binary formats to modern `.docx` using LibreOffice.
2. **Parses & Extracts**: Extracts the student's name, Date of Birth (DOB), and individual criteria scores from their paper using fuzzy matching to handle typos.
3. **Cross-References**: Cross-references parsed scores with the official ground-truth scores in the summary sheet.
4. **Generates Final Word Papers**: Creates polished final evaluation sheets for all 37 students (including absent ones) with centered cell alignments, correct totals, and centered plain-text signatures.
5. **Compiles Detailed Excel Report**: Generates a new summary spreadsheet containing all 22 sub-sections, styled with exact borders, color fills, summary counts, and footer signature blocks.

---

## 📁 Codebase Structure

* **`master.docx`**: The base template file used to generate final student Word papers.
* **`HV_Mau 2_Tong hop KQRL cua SV.xlsx`**: The official summary sheet containing final ground-truth scores for the 37 students.
* **`HV_Mau 2_Chi tiet KQRL.xlsx`** *(Generated)*: Compiled report containing all 22 sub-sections with styling, borders, and footer signatures.
* **`process_and_generate.py`**: The main executable Python script containing the pipeline logic.
* **`STRUCTURE.md`**: Layout mapping of Word paragraphs and Excel cells.
* **`students/`**: Directory containing raw student papers in `.doc` and `.docx`.
* **`generated_students/`** *(Generated)*: Output folder containing final Word files for all students.

---

## ⚙️ Pipeline Workflow (`process_and_generate.py`)

1. **Batched Conversion**: Checks the `students/` folder, converting all `.doc` files to `.docx` via headless LibreOffice and copies existing `.docx` files to a temporary workspace.
2. **Database Load**: Reads the student roster, names, IDs, dates of birth, and final scores from the official summary Excel sheet.
3. **Data Extraction**:
   * Extracts date of birth.
   * Parses the 3 columns of scores (Student, Class, Advisor) for each of the 53 criteria rows in Table 1.
   * Uses fuzzy matching (difflib, ratio $\ge 0.8$) to find equivalent criteria rows even when typos exist (e.g. `nghiệm túc` vs `nghiệp túc`).
4. **Sub-section Aggregation**: Sums up criteria points to calculate scores for all 22 sub-sections (e.g., 1.1, 1.2, 1.3... 5.3) using Advisor scores as final, with Class/Student scores as fallback.
5. **Word Document Generation**:
   * **Present Students**: Populates name, ID, and DOB. Overwrites score cells in Table 1, centers all score alignments, computes totals, and prints the student's name at the bottom.
   * **Absent Students**: Generates blank templates with name/ID, fills all score columns with `0`, rating as `Kém`, and prints the name at the bottom.
   * **Signature Alignment**: Splits student names with 3+ words into 2 lines, aligning both lines center under the `SINH VIÊN` header using tab-stops and spaces, while clearing overlapping textbox placeholders.
6. **Detailed Spreadsheet Compilation**: Writes all sub-sections to a styled report, applying summary statistics, fonts, cell borders, fill colors, and merged footer signatures.

---

## ✨ Key Features & Technical Implementations

### 1. Centered Score Alignments
By default, overwriting table text in `python-docx` resets paragraphs to left-aligned. The script overrides this behavior using:
```python
def write_centered_score(cell, text):
    cell.text = text
    if cell.paragraphs:
        cell.paragraphs[0].alignment = 1 # WD_ALIGN_PARAGRAPH.CENTER
```
This applies to every criteria score, subsection total, and grand total.

### 2. Wrapped Two-Line Centered Signatures
To avoid overlapping textboxes and prevent page overflows while keeping names centered under the `SINH VIÊN` header:
* Splits names into a maximum of 2 words per line (e.g., `Nguyễn Phương` and `Quốc Vương`).
* Centers Line 1 on Paragraph 11 using dynamic space padding: `f"\t" + " " * pad1 + line1`.
* Appends a new Paragraph 12 with matching font style (size 13, Calibri/Times New Roman) and pads spaces from the left margin to center Line 2 under Line 1.
* Clears the legacy student signature textbox XML tags (`w:txbxContent`) to make them invisible.

### 3. Absent Student Formatting
The two absent students (`Nguyễn Trương Minh Triết` and `Vương Hà Hải Đăng`) are auto-detected by their missing files. The script fills their score columns with `0` and sets their final rating to `Kém` (Poor) in both the Word files and the Excel sheet.

---

## 🛠️ Setup & Installation

Ensure you have Python 3 and LibreOffice installed on your system.

### 1. System Dependencies (Debian/Ubuntu)
```bash
sudo apt update
sudo apt install libreoffice python3-pip
```

### 2. Python Packages
```bash
pip3 install python-docx openpyxl lxml
```

---

## 🚀 Usage

Simply run the script from the root workspace:
```bash
python3 process_and_generate.py
```

The script will print progress and generate the files:
```text
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
