# Structural Analysis of `master.docx`

`master.docx` acts as the master template for student training behavior evaluation sheets (*Phiếu đánh giá kết quả rèn luyện*). This document defines the layout, metadata positions, paragraph sequences, and the grid cell structure for student training scores.

---

## 1. Document Elements Sequence

The document's main body contains a sequence of **14 top-level elements** (12 paragraphs and 2 tables) in the following order:

| Index | Element Type | Style / Size | Description / Exact Text |
|---|---|---|---|
| **00** | Paragraph | `Normal` | Spacer (empty string `''`) |
| **01** | Paragraph | `Normal` | Spacer containing multiple spaces |
| **02** | **Table 0** | 1 Row, 2 Cols | Institution headers and National motto (Logo Block) |
| **03** | Paragraph | `Normal` | Title: `'PHIẾU ĐÁNH GIÁ KẾT QUẢ RÈN LUYỆN'` |
| **04** | Paragraph | `Normal` | Term & Year: `'Học kỳ: II      Năm học: 2025-2026'` |
| **05** | Paragraph | `Normal` | Name placeholder: `'Họ và tên:\t\t\t\t\t\tNgày sinh:'` (6 tabs `\t`) |
| **06** | Paragraph | `Normal` | Student ID: `'Mã số sinh viên: N25DECE0  \t\t\tLớp: E25CQCE02-N'` (3 tabs `\t`) |
| **07** | Paragraph | `Normal` | Spacer (empty string `''`) |
| **08** | **Table 1** | 56 Rows | The detailed grading and evaluation criteria matrix |
| **09** | Paragraph | `Normal` | Spacer (empty string `''`) |
| **10** | Paragraph | `Normal` | Spacer (empty string `''`) |
| **11** | Paragraph | `Normal` | Spacer (empty string `''`) |
| **12** | Paragraph | `Normal` | Spacer (empty string `''`) |
| **13** | Paragraph | `Normal` | Student signature line: `'........................\t       Ngô Trí Long\t          '` |

---

## 2. Table Structures

### Table 0: Logo & Header Block
- **Dimensions**: 1 row, 2 columns.
- **Cell (0, 0)**: Left header containing the institution details:
  `'HỌC VIỆN CN BƯU CHÍNH VIỄN THÔNG\nHỌC VIỆN CN BCVT CƠ SỞ TẠI TP. HCM\n\n'`
- **Cell (0, 1)**: Right header containing the national motto:
  `'CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc\n'`

### Table 1: Score & Evaluation Criteria Matrix
- **Dimensions**: 56 rows (0-55).
- **Columns**: 12 columns for Rows 0-54, and 15 columns for Row 55 (the signature row).
- **Row-by-Row Layout**:
  - **Rows 00–01**: Table Headers.
  - **Row 02**: Section 1 Header (*Tiêu chí 1: Đánh giá về ý thức tham gia học tập*).
  - **Rows 03–19**: Items for Criterion 1 (including mutual exclusive academic ratings in Rows 6-10 and penalty rows).
  - **Row 20**: Total row for Criterion 1 (*Mức điểm tối đa Tiêu chí 1*).
  - **Row 21**: Section 2 Header (*Tiêu chí 2: Đánh giá về ý thức chấp hành nội quy, quy chế, quy định...*).
  - **Rows 22–29**: Items for Criterion 2 (including penalty rows).
  - **Row 30**: Total row for Criterion 2 (*Mức điểm tối đa Tiêu chí 2*).
  - **Row 31**: Section 3 Header (*Tiêu chí 3: Đánh giá về ý thức và kết quả tham gia hoạt động CT-XH...*).
  - **Rows 32–36**: Items for Criterion 3.
  - **Row 37**: Total row for Criterion 3 (*Mức điểm tối đa Tiêu chí 3*).
  - **Row 38**: Section 4 Header (*Tiêu chí 4: Đánh giá ý thức công dân trong quan hệ cộng đồng*).
  - **Rows 39–44**: Items for Criterion 4 (including penalty rows).
  - **Row 45**: Total row for Criterion 4 (*Mức điểm tối đa Tiêu chí 4*).
  - **Row 46**: Section 5 Header (*Tiêu chí 5: Đánh giá về ý thức và kết quả tham gia phụ trách lớp...*).
  - **Rows 47–51**: Items for Criterion 5.
  - **Row 52**: Total row for Criterion 5 (*Mức điểm tối đa Tiêu chí 5*).
  - **Row 53**: Grand Total row (*TỔNG CỘNG*).
  - **Row 54**: Date line: `Cols 3-11` merged as `'TP. HCM, ngày 1 tháng 8 năm 2026'`.
  - **Row 55**: Signature Blocks:
    - `Cols 0-3`: `'XÁC NHẬN CỦA  CỐ VẤN HỌC TẬP'`
    - `Cols 4-5`: `'TM. BAN CÁN SỰ LỚP TRƯỞNG'`
    - `Cols 6-9`: `'TM. BCH CHI ĐOÀN BÍ THƯ'`
    - `Cols 10-14`: `'SINH VIÊN'`

---

## 3. Score Column Access and Merging Rules

Rows 00–54 are built on a 12-column grid. Because cells are merged horizontally in different ways depending on the section, column spans differ, but the **grid indices** mapped by python-docx remain perfectly consistent.

The table uses two horizontal merge patterns in the score sections:
1. **Tiêu chí 1 and 2**:
   - Student Self-Score: Column 7 (span=1)
   - Class Score: Columns 8-10 (span=3)
   - Advisor Score: Column 11 (span=1)
2. **Tiêu chí 3, 4, 5, and Total Row (53)**:
   - Student Self-Score: Column 7 (span=1)
   - Class Score: Columns 8-9 (span=2)
   - Advisor Score: Columns 10-11 (span=2)

### Universal Score Index Formula
Using python-docx's grid indices `row.cells[col_idx]`, scores for any active row can be accessed or written using these constant indices:
- **Student self-evaluation score**: `row.cells[7]`
- **Class evaluation score**: `row.cells[8]`
- **Academic Advisor (CVHT) evaluation score**: `row.cells[11]`

---

## 4. Mapping with Excel Summary Database

The student scores in the Excel database (`HV_Mau 2_Tong hop KQRL cua SV.xlsx`) map directly to the total score rows for each criterion inside Table 1 of `master.docx`:

| Excel Column | Excel Column Header | Word Table 1 Row | Word Cell text | Description |
|---|---|---|---|---|
| **Col 1 (B)** | Last and Middle Name | Paragraph 04 | `Họ và tên: <Name>` | Student name |
| **Col 2 (C)** | First Name | Paragraph 04 / 11 | Sign off | Student first name |
| **Col 3 (D)** | Student ID | Paragraph 05 | `Mã số sinh viên: <ID>` | Student ID prefix |
| **Col 4 (E)** | `Nội dung 1 (Max=20)` | **Row 20** | `Mức điểm tối đa Tiêu chí 1` | Criterion 1 Total Score |
| **Col 5 (F)** | `Nội dung 2 (Max=25)` | **Row 30** | `Mức điểm tối đa Tiêu chí 2` | Criterion 2 Total Score |
| **Col 6 (G)** | `Nội dung 3 (Max=20)` | **Row 37** | `Mức điểm tối đa Tiêu chí 3` | Criterion 3 Total Score |
| **Col 7 (H)** | `Nội dung 4 (Max=25)` | **Row 45** | `Mức điểm tối đa Tiêu chí 4` | Criterion 4 Total Score |
| **Col 8 (I)** | `Nội dung 5 (Max=10)` | **Row 52** | `Mức điểm tối đa Tiêu chí 5` | Criterion 5 Total Score |
| **Col 9 (J)** | `Tổng điểm` | **Row 53** | `TỔNG CỘNG` | Grand Total Score (out of 100) |
| **Col 10 (K)** | `XẾP LOẠI RÈN LUYỆN` | N/A | Calculated | Student classification |
