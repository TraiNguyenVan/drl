# Structural Analysis of `master_v2.docx` & Pipeline Layout

`master_v2.docx` is the active template for the `google-docs-compat` branch. It utilizes paragraph-based student info lines with custom tab stops for precise layout alignment, a logo table, a single main grading criteria table (which includes the signature headers), and a compatibility-wrapped shape textbox for the student's name signature.

---

## 1. Document Body Order

The body of the template contains **12 paragraphs** and **2 tables** in the following sequence:

| Index | Element Type | Description |
|---|---|---|
| **00** | Paragraph | Empty spacer |
| **01** | Paragraph | Spaces (layout padding) |
| **--** | **Table 0** | Logo & header block (1 row × 2 columns, borderless) |
| **02** | Paragraph | Title: `'PHIẾU ĐÁNH GIÁ KẾT QUẢ RÈN LUYỆN'` |
| **03** | Paragraph | Term & year: `'Học kỳ: II      Năm học: 2025-2026'` |
| **04** | Paragraph | Student Name & DOB line: `Họ và tên: <Name>\tNgày sinh: <DOB>` |
| **05** | Paragraph | Student ID & Class line: `Mã số sinh viên: <MSV>\tLớp: E25CQCE02-N` |
| **06** | Paragraph | Empty spacer |
| **07** | Paragraph | Empty spacer |
| **08** | Paragraph | Empty spacer |
| **09** | Paragraph | Empty spacer |
| **10** | Paragraph | Empty spacer |
| **--** | **Table 1** | Grading criteria matrix (56 rows) |
| **11** | Paragraph | Static class monitor name / spacing: `'........................\t       Ngô Trí Long\t          \t\t\t\t'` |

In **python-docx**:
- `doc.paragraphs[n]` references paragraphs only (skipping tables).
- `doc.tables[n]` references tables only.
  - `doc.tables[0]` = Logo & Header block.
  - `doc.tables[1]` = Grading criteria matrix & signature headers.

---

## 2. Student Info Paragraphs & Alignment

Paragraphs 4 and 5 contain the student's personal information. They are aligned without tables to preserve compatibilities across Word processors:
- A custom **Left-Aligned Tab Stop** is programmatically set at **3.8 inches (9.65 cm)** from the left margin on both paragraphs.
- Exactly one tab character (`\t`) separates the first column from the second column.
- All runs are explicitly styled in **Times New Roman** at **13pt** to ensure consistency.

| Paragraph | Column 1 (Left) | Column 2 (Right, aligned at 3.8") |
|---|---|---|
| `doc.paragraphs[4]` | `Họ và tên: <Name>` | `Ngày sinh: <DOB>` |
| `doc.paragraphs[5]` | `Mã số sinh viên: <MSV>` | `Lớp: E25CQCE02-N` |

---

## 3. Table 0 — Logo & Header Block

**Dimensions**: 1 row × 2 columns. No visible borders.

| Cell | Content | Formatting |
|---|---|---|
| `(0, 0)` | `'HỌC VIỆN CN BƯU CHÍNH VIỄN THÔNG\nHỌC VIỆN CN BCVT CƠ SỞ TẠI TP. HCM'` | Normal |
| `(0, 1)` | `'CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc'` | **Bold / Underline** on mottos |

---

## 4. Table 1 — Score & Evaluation Criteria Matrix

**Dimensions**: 56 rows (0–55), 12 columns.

### Row Layout
- **Rows 00–01**: Table headers
- **Row 02**: Section 1 header (`I. Đánh giá về ý thức học tập...`)
- **Rows 03–19**: Criterion 1 items
- **Row 20**: **Section 1 total** (`Mức điểm tối đa Tiêu chí 1`)
- **Row 21**: Section 2 header (`II. Đánh giá về ý thức chấp hành...`)
- **Rows 22–29**: Criterion 2 items
- **Row 30**: **Section 2 total**
- **Row 31**: Section 3 header (`III. Đánh giá về ý thức tham gia...`)
- **Rows 32–36**: Criterion 3 items
- **Row 37**: **Section 3 total**
- **Row 38**: Section 4 header (`IV. Đánh giá về ý thức công dân...`)
- **Rows 39–44**: Criterion 4 items
- **Row 45**: **Section 4 total**
- **Row 46**: Section 5 header (`V. Đánh giá về ý thức tham gia...`)
- **Rows 47–51**: Criterion 5 items
- **Row 52**: **Section 5 total**
- **Row 53**: **Grand total** (`TỔNG CỘNG`)
- **Row 54**: Date line (merged columns): `'TP. HCM, ngày 1 tháng 8 năm 2026'`
- **Row 55**: Signature headers: `['XÁC NHẬN CỦA  CỐ VẤN HỌC TẬP', 'TM. BAN CÁN SỰ LỚP TRƯỞNG', 'TM. BCH CHI ĐOÀN BÍ THƯ', 'SINH VIÊN']`

### Score Column Indices
For all criteria rows:
- `cells[7]` — Student self-evaluation score (SV)
- `cells[8]` — Class evaluation score (Lớp)
- `cells[11]` — Academic advisor score (CVHT)

---

## 5. Student Name Signature Textbox

To handle absolute layout positioning of the student's name under the signature header without using a table layout:
- The template uses a compatibility-wrapped shape textbox in a drawing container.
- It contains two text boxes with placeholders `##STUDENT` and `_NAME##` (representing the modern `<wps:txbx>` element and the legacy `<v:textbox>` element respectively).
- The pipeline splits the student's name: if the student has 3 or fewer words in their name, the entire name is placed in `##STUDENT` (and `_NAME##` is cleared). If the student has 4 or more words, the first two words are placed in `##STUDENT` and the remaining words in `_NAME##`.

---

## 6. Score Column Mapping (Excel → Word)

| Excel Column | Excel Header | Word Table 1 Row | Description |
|---|---|---|---|
| B | Họ đệm | — | Student last & middle name |
| C | Tên | textbox / `_NAME##` | Student first name (used in signature box) |
| D | Mã sinh viên | `doc.paragraphs[5]` | MSV |
| E | Nội dung 1 (Max=20) | Row 20 | Criterion 1 total |
| F | Nội dung 2 (Max=25) | Row 30 | Criterion 2 total |
| G | Nội dung 3 (Max=20) | Row 37 | Criterion 3 total |
| H | Nội dung 4 (Max=25) | Row 45 | Criterion 4 total |
| I | Nội dung 5 (Max=10) | Row 52 | Criterion 5 total |
| J | Tổng điểm | Row 53 | Grand total |
| K | Xếp loại rèn luyện | — | Classification |
