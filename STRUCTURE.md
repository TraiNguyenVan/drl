# Structural Analysis of `master_v2.docx`

`master_v2.docx` is the active template for the `google-docs-compat` branch. It replaces the original `master.docx` by eliminating all floating textboxes (`wp:anchor` / VML `w:pict`) in favour of borderless tables, making it fully compatible with Google Docs, Microsoft Word, and LibreOffice.

> **Note for AI agents**: Always read `GEMINI.md` first for project-level constraints before editing anything here.

---

## 1. Document Body Order

The body contains **5 paragraphs** and **4 tables** in the following sequence:

| Index | Element Type | Description |
|---|---|---|
| **00** | Paragraph | Empty spacer |
| **01** | Paragraph | Spaces (layout padding) |
| **02** | **Table 0** | Logo & header block |
| **03** | Paragraph | Title: `'PHIẾU ĐÁNH GIÁ KẾT QUẢ RÈN LUYỆN'` |
| **04** | Paragraph | Term & year: `'Học kỳ: II      Năm học: 2025-2026'` |
| **05** | **Table 1** | Info table (name / DOB / MSV / class) — borderless 2×2 |
| **06** | Paragraph | Empty spacer |
| **07** | **Table 2** | Grading criteria matrix — 56 rows |
| **08** | **Table 3** | Signature table — borderless 2×4 |

In **python-docx**:
- `doc.paragraphs[n]` — skips tables, counts paragraphs only (indices 0–4)
- `doc.tables[n]` — counts all tables (indices 0–3)

---

## 2. Table 0 — Logo & Header Block

**Dimensions**: 1 row × 2 columns. No visible borders.

| Cell | Content | Formatting |
|---|---|---|
| `(0, 0)` Para 0 | `'HỌC VIỆN CN BƯU CHÍNH VIỄN THÔNG'` | Normal |
| `(0, 0)` Para 1 | `'HỌC VIỆN CN BCVT CƠ SỞ TẠI TP. HCM'` | **Bold + Underline** |
| `(0, 1)` Para 0 | `'CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM'` | **Bold** |
| `(0, 1)` Para 1 | `'Độc lập - Tự do - Hạnh phúc'` | **Bold + Underline** |

---

## 3. Table 1 — Info Table (borderless, 2×2)

Replaces the old tab-stop paragraphs (`Họ và tên:\t\t\t\t\t\tNgày sinh:`). Borderless with no visible lines.

| Cell | Placeholder in template | Written by pipeline | Alignment |
|---|---|---|---|
| `rows[0].cells[0]` | `Họ và tên: ##NAME##` | `f"Họ và tên: {s['name']}"` | Left |
| `rows[0].cells[1]` | `Ngày sinh: ##DOB##` | `f"Ngày sinh: {dob_val}"` + `alignment=2` | **Right** |
| `rows[1].cells[0]` | `Mã số sinh viên: ##MSV##` | `f"Mã số sinh viên: {s['msv']}"` | Left |
| `rows[1].cells[1]` | `Lớp: E25CQCE02-N` | *(static, not overwritten)* | **Right** |

> **Absent students**: DOB cell is set to `"Ngày sinh:"` (no date) with `alignment=2`.

---

## 4. Table 2 — Score & Evaluation Criteria Matrix

**Dimensions**: 56 rows (0–55), 12–15 columns depending on row.

### Row layout
| Row(s) | Content |
|---|---|
| 00–01 | Table headers |
| 02 | Section 1 header |
| 03–19 | Criterion 1 items |
| **20** | **Section 1 total** (`Mức điểm tối đa Tiêu chí 1`) |
| 21 | Section 2 header |
| 22–29 | Criterion 2 items |
| **30** | **Section 2 total** |
| 31 | Section 3 header |
| 32–36 | Criterion 3 items |
| **37** | **Section 3 total** |
| 38 | Section 4 header |
| 39–44 | Criterion 4 items |
| **45** | **Section 4 total** |
| 46 | Section 5 header |
| 47–51 | Criterion 5 items |
| **52** | **Section 5 total** |
| **53** | **Grand total** (`TỔNG CỘNG`) |
| 54 | Date line (merged cols 3–11): `'TP. HCM, ngày 1 tháng 8 năm 2026'` |
| 55 | Signature headers (see Table 3 for names) |

### Score column indices (constant for all data rows)
| Column | Description |
|---|---|
| `cells[7]` | Student self-evaluation score |
| `cells[8]` | Class evaluation score |
| `cells[11]` | Academic advisor score |

> All score writes use `write_centered_score(cell, text)` to preserve `CENTER` alignment after text replacement.

### Section total rows used by pipeline
`[20, 30, 37, 45, 52, 53]` — written in both the absent and present student paths.

---

## 5. Table 3 — Signature Table (borderless, 2×4)

Replaces the old floating textboxes. Borderless with no visible lines.

**Row 0**: Empty signing space (~1.27 cm tall)
**Row 1**: Signatory names

| Cell | Content | Source |
|---|---|---|
| `rows[1].cells[0]` | `Nguyễn Trung Hiếu` | Fixed in `master_v2.docx` |
| `rows[1].cells[1]` | `Ngô Trí Long` | Fixed in `master_v2.docx` |
| `rows[1].cells[2]` | `Nguyễn Phương Quốc Vương` | Fixed in `master_v2.docx` |
| `rows[1].cells[3]` | `<student full name>` | Written by `write_centered_score()` |

All cells are center-aligned.

---

## 6. Score Column Mapping (Excel → Word)

| Excel Col | Header | Word Table 2 Row | Description |
|---|---|---|---|
| B | Last & Middle Name | — | Student name |
| C | First Name | Table 3 `rows[1].cells[3]` | Used as sign-off |
| D | Student ID | Table 1 `rows[1].cells[0]` | MSV |
| E | Nội dung 1 (Max=20) | Row 20 | Criterion 1 total |
| F | Nội dung 2 (Max=25) | Row 30 | Criterion 2 total |
| G | Nội dung 3 (Max=20) | Row 37 | Criterion 3 total |
| H | Nội dung 4 (Max=25) | Row 45 | Criterion 4 total |
| I | Nội dung 5 (Max=10) | Row 52 | Criterion 5 total |
| J | Tổng điểm | Row 53 | Grand total |
| K | Xếp loại rèn luyện | — | Classification (computed) |

---

## 7. Differences from `master.docx` (original / `main` branch)

| Aspect | `master.docx` (`main`) | `master_v2.docx` (this branch) |
|---|---|---|
| Floating textboxes | 2 × `wp:anchor` shapes | **None** |
| Info section | 2 tab-stop paragraphs | Borderless 2×2 table |
| Signature section | Paragraph 11 with space-padding | Borderless 2×4 table |
| `doc.tables` count | 2 | **4** |
| Google Docs safe | ❌ (ghost duplicates, page split) | ✅ |
| LibreOffice page count | 3 | 3 (LibreOffice quirk, renders 2 in Word/GDocs) |
