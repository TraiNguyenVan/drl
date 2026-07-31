import os
import re
import csv
import subprocess
import shutil
import copy
import difflib
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def write_centered_score(cell, text):
    cell.text = text
    if cell.paragraphs:
        cell.paragraphs[0].alignment = 1  # WD_ALIGN_PARAGRAPH.CENTER

def write_signature_name(cell, name):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = 1  # WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = docx.shared.Pt(0)
    p.paragraph_format.space_after = docx.shared.Pt(0)
    p.paragraph_format.line_spacing = 1.0
    run = p.add_run(name)
    run.font.size = docx.shared.Pt(13)
    run.font.name = "Times New Roman"

def format_signature_name(name):
    words = name.split()
    if len(words) >= 3:
        line1 = " ".join(words[:2])
        line2 = " ".join(words[2:])
        return f"{line1}\n{line2}"
    return name

def fill_info(doc, name, msv, dob_val):
    # Paragraph 4: Name and DOB
    p4 = doc.paragraphs[4]
    p4.text = "" # Clear all text/runs
    
    # Set the custom tab stop on p4 (3.8 inches is 9.65 cm)
    p4.paragraph_format.tab_stops.add_tab_stop(docx.shared.Inches(3.8), docx.enum.text.WD_TAB_ALIGNMENT.LEFT)
    
    r_name_label = p4.add_run("Họ và tên: ")
    r_name_label.font.name = "Times New Roman"
    r_name_label.font.size = docx.shared.Pt(13)
    
    r_name_val = p4.add_run(name)
    r_name_val.font.name = "Times New Roman"
    r_name_val.font.size = docx.shared.Pt(13)
    
    p4.add_run("\t") # Tab to jump to 3.8 inches
    
    r_dob_label = p4.add_run("Ngày sinh: ")
    r_dob_label.font.name = "Times New Roman"
    r_dob_label.font.size = docx.shared.Pt(13)
    
    if dob_val:
        r_dob_val = p4.add_run(dob_val)
        r_dob_val.font.name = "Times New Roman"
        r_dob_val.font.size = docx.shared.Pt(13)
        
    # Paragraph 5: Student ID and Class
    p5 = doc.paragraphs[5]
    p5.text = "" # Clear all text/runs
    
    # Set the custom tab stop on p5
    p5.paragraph_format.tab_stops.add_tab_stop(docx.shared.Inches(3.8), docx.enum.text.WD_TAB_ALIGNMENT.LEFT)
    
    r_msv_label = p5.add_run("Mã số sinh viên: ")
    r_msv_label.font.name = "Times New Roman"
    r_msv_label.font.size = docx.shared.Pt(13)
    
    r_msv_val = p5.add_run(msv)
    r_msv_val.font.name = "Times New Roman"
    r_msv_val.font.size = docx.shared.Pt(13)
    
    p5.add_run("\t") # Tab to jump to 3.8 inches
    
    r_class_label = p5.add_run("Lớp: ")
    r_class_label.font.name = "Times New Roman"
    r_class_label.font.size = docx.shared.Pt(13)
    
    r_class_val = p5.add_run("E25CQCE02-N")
    r_class_val.font.name = "Times New Roman"
    r_class_val.font.size = docx.shared.Pt(13)
            
    # Floating Textbox Student Name replacement (split into 2 lines)
    words = name.split()
    if len(words) >= 3:
        line1 = " ".join(words[:2])
        line2 = " ".join(words[2:])
    else:
        line1 = name
        line2 = ""
        
    p_elms = doc.element.xpath('.//*[local-name()="p"]')
    for p_elm in p_elms:
        p = docx.text.paragraph.Paragraph(p_elm, doc)
        for r in p.runs:
            if "##STUDENT" in r.text:
                r.text = r.text.replace("##STUDENT", line1)
            if "_NAME##" in r.text:
                r.text = r.text.replace("_NAME##", line2)

# Directories
workspace = "/home/trai/workspace/drl"
students_dir = os.path.join(workspace, "students")
output_dir = os.path.join(workspace, "generated_students")
os.makedirs(output_dir, exist_ok=True)

scratch_dir = "/home/trai/.gemini/antigravity-cli/brain/a6bf4666-97c8-4c12-851f-ed2c433368a0/scratch"
temp_dir = os.path.join(scratch_dir, "temp_conv_all")
if os.path.exists(temp_dir):
    shutil.rmtree(temp_dir)
os.makedirs(temp_dir, exist_ok=True)

# 1. Convert doc files to docx and copy existing docx files
print("Converting and preparing raw student files...")
for name in os.listdir(students_dir):
    src_path = os.path.join(students_dir, name)
    if name.endswith(".doc"):
        dest_name = name[:-4] + ".docx"
        dest_path = os.path.join(temp_dir, dest_name)
        subprocess.run([
            "soffice", "--headless", "--convert-to", "docx",
            "--outdir", temp_dir, src_path
        ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    elif name.endswith(".docx"):
        shutil.copy2(src_path, os.path.join(temp_dir, name))
print("Preparation completed.")

# 2. Text Normalization for robust description matching
def normalize_desc(text):
    return re.sub(r'[^a-z0-9àáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ]', '', text.lower())

def is_max_pts_refined(val):
    val_clean = val.strip().lower()
    if val_clean.endswith("điểm") and len(val_clean) < 15:
        return True
    if val_clean == "100":
        return True
    if val_clean.startswith("-") and len(val_clean) < 6 and any(c.isdigit() for c in val_clean):
        return True
    return False

# 3. Load master template active rows
master_doc_path = os.path.join(workspace, "master_v2.docx")
master_doc = docx.Document(master_doc_path)
master_table = master_doc.tables[1]
master_rows_info = {} # normalized_desc -> {row_idx, max_pts_str}

for r_idx, row in enumerate(master_table.rows):
    unique_cells = []
    seen_tcs = set()
    for cell in row.cells:
        tc_id = id(cell._tc)
        if tc_id not in seen_tcs:
            seen_tcs.add(tc_id)
            unique_cells.append(cell)
            
    max_idx = None
    for idx, cell in enumerate(unique_cells):
        val = cell.text.strip().replace('\n', ' ')
        if is_max_pts_refined(val):
            max_idx = idx
            break
            
    if max_idx is not None and max_idx > 0:
        desc = unique_cells[max_idx - 1].text.strip().replace('\n', ' ')
        norm_desc = normalize_desc(desc)
        master_rows_info[norm_desc] = {
            "row_idx": r_idx,
            "desc": desc,
            "max": unique_cells[max_idx].text.strip()
        }

# Similarity helper for description matching with typos
def get_best_master_match(student_norm_desc, master_keys):
    # Try exact match first
    if student_norm_desc in master_keys:
        return student_norm_desc
    # Fuzzy match
    best_key = None
    best_ratio = 0.0
    for key in master_keys:
        ratio = difflib.SequenceMatcher(None, student_norm_desc, key).ratio()
        if ratio > best_ratio:
            best_ratio = ratio
            best_key = key
    if best_ratio >= 0.8:
        return best_key
    return None

# Helper to parse float scores
def parse_score_val(val):
    if not val:
        return 0.0
    val_clean = val.strip().replace(',', '.')
    val_clean = re.sub(r'[^0-9.-]', '', val_clean)
    try:
        return float(val_clean)
    except ValueError:
        return 0.0

# 4. Load Excel database (Ground Truth)
db_path = os.path.join(workspace, "HV_Mau 2_Tong hop KQRL cua SV.xlsx")
wb_db = openpyxl.load_workbook(db_path, data_only=True)
sheet_db = wb_db.active

students_db = []
for r in range(14, sheet_db.max_row + 1):
    tt = sheet_db.cell(r, 1).value
    last_name = sheet_db.cell(r, 2).value
    first_name = sheet_db.cell(r, 3).value
    msv = sheet_db.cell(r, 4).value
    
    is_valid_tt = False
    if isinstance(tt, (int, float)):
        is_valid_tt = True
    elif isinstance(tt, str):
        try:
            float(tt)
            is_valid_tt = True
        except ValueError:
            pass
            
    if is_valid_tt and last_name and first_name and msv:
        msv_str = str(msv).strip()
        if msv_str.startswith("N25"):
            full_name = f"{str(last_name).strip()} {str(first_name).strip()}"
            full_name = " ".join(full_name.split())
            
            # Extract criteria totals from Excel
            tc_scores = [sheet_db.cell(r, c).value for c in range(5, 10)]
            tc_scores = [float(x) if x is not None else 0.0 for x in tc_scores]
            total_score = sheet_db.cell(r, 10).value
            total_score = float(total_score) if total_score is not None else 0.0
            rating = sheet_db.cell(r, 11).value
            notes = sheet_db.cell(r, 12).value
            
            students_db.append({
                "tt": int(float(tt)),
                "last_name": str(last_name).strip(),
                "first_name": str(first_name).strip(),
                "name": full_name,
                "msv": msv_str,
                "excel_tc": tc_scores,
                "excel_total": total_score,
                "excel_rating": rating,
                "excel_notes": notes
            })

# 4b. Load new vertical format ai_studio_code.csv
csv_path = os.path.join(workspace, "ai_studio_code.csv")
csv_scores_by_msv = {}
if os.path.exists(csv_path):
    with open(csv_path, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            msv = row["student_id"].strip()
            crit_id = row["criterion_id"].strip()
            if msv not in csv_scores_by_msv:
                csv_scores_by_msv[msv] = {}
            csv_scores_by_msv[msv][crit_id] = {
                "sv": row["student_score"].strip() if row.get("student_score") is not None else "",
                "lop": row["class_score"].strip() if row.get("class_score") is not None else "",
                "cvht": row["advisor_score"].strip() if row.get("advisor_score") is not None else ""
            }

# Subsection Row Mappings in master.docx Table 1
subsection_mapping = {
    "1.1": [4],
    "1.2": [6, 7, 8, 9, 10, 11],
    "1.3": [12, 14, 15, 16, 17],
    "1.4": [18],
    "1.5": [19],
    "2.1": [22, 24, 25],
    "2.2": [26, 27],
    "2.3": [28, 29],
    "3.1": [32],
    "3.2": [33],
    "3.3": [34],
    "3.4": [35],
    "3.5": [36],
    "4.1": [39],
    "4.2": [40],
    "4.3": [41],
    "4.4": [42],
    "4.5": [43],
    "4.6": [44],
    "5.1": [47],
    "5.2": [48],
    "5.3": [50, 51]
}

# Helper to write score from CSV to table cells
def write_csv_score_to_table(table, criterion_id, role_col, score_str):
    if not score_str:
        return
        
    try:
        val = float(score_str)
        if val.is_integer():
            val_str = str(int(val))
        else:
            val_str = str(val)
    except ValueError:
        val_str = score_str
        val = 0.0

    # Map criterion_id to row index and write
    if criterion_id == "1.1":
        write_centered_score(table.rows[4].cells[role_col], val_str)
    elif criterion_id == "1.2":
        if val == 10:
            write_centered_score(table.rows[6].cells[role_col], "10")
        elif val == 8:
            write_centered_score(table.rows[7].cells[role_col], "8")
        elif val == 6:
            write_centered_score(table.rows[8].cells[role_col], "6")
        elif val == 4:
            write_centered_score(table.rows[9].cells[role_col], "4")
        elif val == 0:
            write_centered_score(table.rows[10].cells[role_col], "0")
    elif criterion_id == "1.3":
        write_centered_score(table.rows[12].cells[role_col], "4")
        if val < 4:
            penalty = int(val - 4)
            write_centered_score(table.rows[14].cells[role_col], str(penalty))
    elif criterion_id == "1.4":
        write_centered_score(table.rows[18].cells[role_col], val_str)
    elif criterion_id == "1.5":
        write_centered_score(table.rows[19].cells[role_col], val_str)
    elif criterion_id == "2.1":
        write_centered_score(table.rows[22].cells[role_col], "15")
        if val < 15:
            penalty = int(val - 15)
            write_centered_score(table.rows[24].cells[role_col], str(penalty))
    elif criterion_id == "2.2":
        write_centered_score(table.rows[26].cells[role_col], "5")
        if val < 5:
            penalty = int(val - 5)
            write_centered_score(table.rows[27].cells[role_col], str(penalty))
    elif criterion_id == "2.3":
        if val > 0:
            write_centered_score(table.rows[28].cells[role_col], val_str)
        else:
            write_centered_score(table.rows[28].cells[role_col], "0")
    elif criterion_id == "3.1":
        write_centered_score(table.rows[32].cells[role_col], val_str)
    elif criterion_id == "3.2":
        write_centered_score(table.rows[33].cells[role_col], val_str)
    elif criterion_id == "3.3":
        write_centered_score(table.rows[34].cells[role_col], val_str)
    elif criterion_id == "3.4":
        write_centered_score(table.rows[35].cells[role_col], val_str)
    elif criterion_id == "4.1":
        write_centered_score(table.rows[39].cells[role_col], val_str)
    elif criterion_id == "4.2":
        write_centered_score(table.rows[40].cells[role_col], val_str)
    elif criterion_id == "4.3":
        write_centered_score(table.rows[41].cells[role_col], val_str)
    elif criterion_id == "4.4":
        write_centered_score(table.rows[42].cells[role_col], val_str)
    elif criterion_id == "4.5":
        write_centered_score(table.rows[43].cells[role_col], val_str)
    elif criterion_id == "5.1":
        write_centered_score(table.rows[47].cells[role_col], val_str)
    elif criterion_id == "5.2":
        write_centered_score(table.rows[48].cells[role_col], val_str)
    elif criterion_id == "5.3":
        if val > 0:
            write_centered_score(table.rows[50].cells[role_col], val_str)
        else:
            write_centered_score(table.rows[50].cells[role_col], "0")
    elif criterion_id == "TC1":
        write_centered_score(table.rows[20].cells[role_col], val_str)
    elif criterion_id == "TC2":
        write_centered_score(table.rows[30].cells[role_col], val_str)
    elif criterion_id == "TC3":
        write_centered_score(table.rows[37].cells[role_col], val_str)
    elif criterion_id == "TC4":
        write_centered_score(table.rows[45].cells[role_col], val_str)
    elif criterion_id == "TC5":
        write_centered_score(table.rows[52].cells[role_col], val_str)
    elif criterion_id == "TOTAL":
        write_centered_score(table.rows[53].cells[role_col], val_str)

# 5. Process each student and generate docx and collect scores
print("\nProcessing student files...")
processed_students = []

for s in students_db:
    # Match student file
    found_file = None
    for filename in os.listdir(temp_dir):
        if s["msv"] in filename:
            found_file = os.path.join(temp_dir, filename)
            break
            
    student_record = copy.deepcopy(s)
    student_record["dob"] = ""
    for sub in subsection_mapping.keys():
        student_record[f"sub_{sub}"] = 0.0
        
    dob_val = ""
    if found_file:
        doc_student = docx.Document(found_file)
        # Extract DOB
        for p in doc_student.paragraphs:
            txt = p.text.strip()
            if "ngày sinh" in txt.lower():
                match = re.search(r"Ngày\s+sinh\s*:\s*([^\t\n\r]+)", txt, re.IGNORECASE)
                if match:
                    dob_val = match.group(1).strip()
                    break
    student_record["dob"] = dob_val
    
    if not found_file:
        print(f"  TT={s['tt']}: {s['name']} ({s['msv']}) -> MISSING FILE (Absent)")
    
    # Generate clean word document copy
    dest_doc_path = os.path.join(output_dir, f"{s['name']}_{s['msv']}.docx")
    shutil.copy2(master_doc_path, dest_doc_path)
    doc_out = docx.Document(dest_doc_path)
    
    # Fill personal info paragraphs and student signature name textbox
    fill_info(doc_out, s['name'], s['msv'], dob_val)
    
    # Set all active scores columns in grading table (tables[1]) to empty
    table_out = doc_out.tables[1]
    for r_idx in range(len(table_out.rows)):
        unique_cells = []
        seen_tcs = set()
        for cell in table_out.rows[r_idx].cells:
            tc_id = id(cell._tc)
            if tc_id not in seen_tcs:
                seen_tcs.add(tc_id)
                unique_cells.append(cell)
        max_idx = None
        for idx, cell in enumerate(unique_cells):
            val = cell.text.strip()
            if is_max_pts_refined(val):
                max_idx = idx
                break
        if max_idx is not None and max_idx > 0:
            if max_idx + 1 < len(unique_cells): unique_cells[max_idx + 1].text = ""
            if max_idx + 2 < len(unique_cells): unique_cells[max_idx + 2].text = ""
            if max_idx + 3 < len(unique_cells): unique_cells[max_idx + 3].text = ""
            
    # Write scores from CSV
    student_csv = csv_scores_by_msv.get(s["msv"])
    if student_csv:
        for crit_id, roles in student_csv.items():
            write_csv_score_to_table(table_out, crit_id, 7, roles["sv"])
            write_csv_score_to_table(table_out, crit_id, 8, roles["lop"])
            write_csv_score_to_table(table_out, crit_id, 11, roles["cvht"])
            
    # Sum up subsection scores from CSV (using Advisor score as final, class/student fallback)
    for sub in subsection_mapping.keys():
        sub_sum = 0.0
        if student_csv and sub in student_csv:
            s_dict = student_csv[sub]
            final_val = s_dict["cvht"] if s_dict["cvht"] else (s_dict["lop"] if s_dict["lop"] else s_dict["sv"])
            sub_sum = parse_score_val(final_val)
        student_record[f"sub_{sub}"] = sub_sum
        
    # Overwrite Class and Advisor totals/grand totals with master Excel values
    totals_rows = [20, 30, 37, 45, 52]
    for idx, r_idx in enumerate(totals_rows):
        val_str = str(s["excel_tc"][idx]).replace(".0", "")
        write_centered_score(table_out.rows[r_idx].cells[8], val_str)   # Class column
        write_centered_score(table_out.rows[r_idx].cells[11], val_str)  # Advisor column
        
    gt_str = str(s["excel_total"]).replace(".0", "")
    write_centered_score(table_out.rows[53].cells[8], gt_str)   # Class column
    write_centered_score(table_out.rows[53].cells[11], gt_str)  # Advisor column
        
    doc_out.save(dest_doc_path)
    processed_students.append(student_record)

print(f"Successfully generated {len(processed_students)} student Word files.")

# 6. Create detailed Excel report
print("\nGenerating detailed Excel report...")
wb_out = openpyxl.Workbook()
sheet_out = wb_out.active
sheet_out.title = "Mau 2"
sheet_out.views.sheetView[0].showGridLines = True

# Copy title rows 1-11
for r in range(1, 12):
    for c in range(1, 13):
        cell_src = sheet_db.cell(r, c)
        cell_dest = sheet_out.cell(r, c)
        cell_dest.value = cell_src.value
        if cell_src.has_style:
            cell_dest.font = copy.copy(cell_src.font)
            cell_dest.alignment = copy.copy(cell_src.alignment)

# Re-merge titles proportionally across columns A to AI (Cols 1 to 35)
title_merges = [
    (1, 2, 4, 2, 4),    # B1:D1
    (2, 1, 6, 1, 18),   # A2:F2 -> A2:R2
    (2, 7, 12, 19, 35),  # G2:L2 -> S2:AI2
    (3, 1, 6, 1, 18),   # A3:F3 -> A3:R3
    (3, 7, 12, 19, 35),  # G3:L3 -> S3:AI3
    (5, 6, 12, 18, 35),  # F5:L5 -> R5:AI5
    (7, 1, 12, 1, 35),   # A7:L7 -> A7:AI7
]
for r_idx, min_c, max_c, new_min_c, new_max_c in title_merges:
    # unmerge original if openpyxl merged them automatically
    # copy value of top-left cell to new top-left
    val = sheet_db.cell(r_idx, min_c).value
    sheet_out.cell(r_idx, new_min_c).value = val
    sheet_out.merge_cells(start_row=r_idx, start_column=new_min_c, end_row=r_idx, end_column=new_max_c)

# Explicit cell copy for other specific title cells
sheet_out.cell(9, 1).value = sheet_db.cell(9, 1).value
sheet_out.merge_cells(start_row=9, start_column=1, end_row=9, end_column=35)

# Styling fonts and borders for header rows 12 & 13
font_h1 = Font(name="Times New Roman", size=11, bold=True)
font_h2 = Font(name="Times New Roman", size=10, bold=True)
align_h = Alignment(horizontal="center", vertical="center", wrap_text=True)

thin_border = Side(style='thin')
medium_border = Side(style='medium')
double_border = Side(style='double')

# Write main headers row 12 and 13 with explicit merges
explicit_merges = [
    (12, 1, 13, 1, "TT"),
    (12, 2, 13, 3, "Họ và tên"),
    (12, 4, 13, 4, "Mã sinh viên"),
    (12, 5, 13, 5, "Ngày sinh"),
    (12, 6, 12, 33, "ĐIỂM ĐÁNH GIÁ"),
    (12, 34, 13, 34, "XẾP LOẠI RÈN LUYỆN"),
    (12, 35, 13, 35, "GHI CHÚ")
]

for start_r, start_c, end_r, end_c, text in explicit_merges:
    cell = sheet_out.cell(start_r, start_c, text)
    cell.font = font_h1
    cell.alignment = align_h
    sheet_out.merge_cells(start_row=start_r, start_column=start_c, end_row=end_r, end_column=end_c)

# Write subheaders row 13
sub_sections_list = [
    # TC1
    ("1.1", 6), ("1.2", 7), ("1.3", 8), ("1.4", 9), ("1.5", 10),
    ("Nội dung 1 (Max=20)", 11),
    # TC2
    ("2.1", 12), ("2.2", 13), ("2.3", 14),
    ("Nội dung 2 (Max=25)", 15),
    # TC3
    ("3.1", 16), ("3.2", 17), ("3.3", 18), ("3.4", 19), ("3.5", 20),
    ("Nội dung 3 (Max=20)", 21),
    # TC4
    ("4.1", 22), ("4.2", 23), ("4.3", 24), ("4.4", 25), ("4.5", 26), ("4.6", 27),
    ("Nội dung 4 (Max=25)", 28),
    # TC5
    ("5.1", 29), ("5.2", 30), ("5.3", 31),
    ("Nội dung 5 (Max=10)", 32),
    # AG13
    ("Tổng điểm", 33)
]

for label, col_idx in sub_sections_list:
    cell = sheet_out.cell(13, col_idx, label)
    cell.font = font_h2
    cell.alignment = align_h

# Format headers cells borders (Rows 12 & 13)
for r in [12, 13]:
    for c in range(1, 36):
        cell = sheet_out.cell(r, c)
        
        # Determine borders
        left = double_border if c == 1 else (medium_border if c in [2, 4, 5, 6, 12, 16, 22, 29, 33, 34, 35] else None)
        right = medium_border if c in [1, 3, 4, 5, 11, 15, 21, 28, 32, 33, 34, 35] else None
        top = double_border if r == 12 else None
        bottom = medium_border if r == 13 else None
        
        cell.border = Border(left=left, right=right, top=top, bottom=bottom)

# Write student data rows (Starting at row 14)
font_data = Font(name="Times New Roman", size=10, bold=False)
font_total_col = Font(name="Times New Roman", size=11, bold=True)
align_center = Alignment(horizontal="center", vertical="center")
align_left = Alignment(horizontal="left", vertical="center")

dotted_side = Side(style='dotted')

current_row = 14
for s in processed_students:
    sheet_out.cell(current_row, 1, float(s["tt"])).alignment = align_center
    sheet_out.cell(current_row, 2, s["last_name"]).alignment = align_center
    sheet_out.cell(current_row, 3, s["first_name"]).alignment = align_center
    sheet_out.cell(current_row, 4, s["msv"]).alignment = align_center
    sheet_out.cell(current_row, 5, s["dob"]).alignment = align_center
    
    # Subsections
    sheet_out.cell(current_row, 6, s["sub_1.1"]).alignment = align_center
    sheet_out.cell(current_row, 7, s["sub_1.2"]).alignment = align_center
    sheet_out.cell(current_row, 8, s["sub_1.3"]).alignment = align_center
    sheet_out.cell(current_row, 9, s["sub_1.4"]).alignment = align_center
    sheet_out.cell(current_row, 10, s["sub_1.5"]).alignment = align_center
    
    # TC1 total
    cell_tc1 = sheet_out.cell(current_row, 11, s["excel_tc"][0])
    cell_tc1.alignment = align_center
    cell_tc1.font = font_total_col
    
    sheet_out.cell(current_row, 12, s["sub_2.1"]).alignment = align_center
    sheet_out.cell(current_row, 13, s["sub_2.2"]).alignment = align_center
    sheet_out.cell(current_row, 14, s["sub_2.3"]).alignment = align_center
    
    # TC2 total
    cell_tc2 = sheet_out.cell(current_row, 15, s["excel_tc"][1])
    cell_tc2.alignment = align_center
    cell_tc2.font = font_total_col
    
    sheet_out.cell(current_row, 16, s["sub_3.1"]).alignment = align_center
    sheet_out.cell(current_row, 17, s["sub_3.2"]).alignment = align_center
    sheet_out.cell(current_row, 18, s["sub_3.3"]).alignment = align_center
    sheet_out.cell(current_row, 19, s["sub_3.4"]).alignment = align_center
    sheet_out.cell(current_row, 20, s["sub_3.5"]).alignment = align_center
    
    # TC3 total
    cell_tc3 = sheet_out.cell(current_row, 21, s["excel_tc"][2])
    cell_tc3.alignment = align_center
    cell_tc3.font = font_total_col
    
    sheet_out.cell(current_row, 22, s["sub_4.1"]).alignment = align_center
    sheet_out.cell(current_row, 23, s["sub_4.2"]).alignment = align_center
    sheet_out.cell(current_row, 24, s["sub_4.3"]).alignment = align_center
    sheet_out.cell(current_row, 25, s["sub_4.4"]).alignment = align_center
    sheet_out.cell(current_row, 26, s["sub_4.5"]).alignment = align_center
    sheet_out.cell(current_row, 27, s["sub_4.6"]).alignment = align_center
    
    # TC4 total
    cell_tc4 = sheet_out.cell(current_row, 28, s["excel_tc"][3])
    cell_tc4.alignment = align_center
    cell_tc4.font = font_total_col
    
    sheet_out.cell(current_row, 29, s["sub_5.1"]).alignment = align_center
    sheet_out.cell(current_row, 30, s["sub_5.2"]).alignment = align_center
    sheet_out.cell(current_row, 31, s["sub_5.3"]).alignment = align_center
    
    # TC5 total
    cell_tc5 = sheet_out.cell(current_row, 32, s["excel_tc"][4])
    cell_tc5.alignment = align_center
    cell_tc5.font = font_total_col
    
    # Grand total
    cell_gt = sheet_out.cell(current_row, 33, s["excel_total"])
    cell_gt.alignment = align_center
    cell_gt.font = font_total_col
    
    # Rating & Notes
    sheet_out.cell(current_row, 34, s["excel_rating"]).alignment = align_center
    sheet_out.cell(current_row, 35, s["excel_notes"]).alignment = align_left
    
    # Style data font and borders
    for c in range(1, 36):
        cell = sheet_out.cell(current_row, c)
        if cell.font != font_total_col:
            cell.font = font_data
            
        left_border = double_border if c == 1 else (medium_border if c in [2, 4, 5, 6, 12, 16, 22, 29, 33, 34, 35] else None)
        right_border = medium_border if c in [1, 3, 4, 5, 11, 15, 21, 28, 32, 33, 34, 35] else dotted_side
        cell.border = Border(left=left_border, right=right_border, bottom=dotted_side)
        
    current_row += 1

# Border line under the last student row (medium border bottom)
for c in range(1, 36):
    cell = sheet_out.cell(current_row - 1, c)
    cell.border = Border(left=cell.border.left, right=cell.border.right, bottom=medium_border)

# Write footer stats and details
# Row 52: summary text
sheet_out.cell(current_row + 1, 2, "Danh sách có").font = Font(name="Times New Roman", size=11, bold=True)
sheet_out.cell(current_row + 1, 3, len(processed_students)).font = Font(name="Times New Roman", size=11, bold=True)
sheet_out.cell(current_row + 1, 3).alignment = align_center
sheet_out.cell(current_row + 1, 4, " sinh viên").font = Font(name="Times New Roman", size=11, bold=True)

# Row 53: Lưu ý
sheet_out.cell(current_row + 2, 1, "Lưu ý: Kết quả điểm rèn luyện được phân thành các loại: Xuất sắc, Tốt, Khá, Trung bình, Yếu, Kém").font = Font(name="Times New Roman", size=10, italic=True)

# Calculate Rating stats
rating_counts = {"Xuất sắc": 0, "Tốt": 0, "Khá": 0, "Trung bình": 0, "Yếu": 0, "Kém": 0}
for s in processed_students:
    r_val = s["excel_rating"]
    if r_val in rating_counts:
        rating_counts[r_val] += 1
    else:
        # Fallback if there is a classification mismatch
        rating_counts["Kém"] += 1

rating_ranges = [
    ("Xuất sắc", "_ Loại Xuất sắc: Từ 90- đến 100 điểm"),
    ("Tốt", "_ Loại Tốt: Từ 80 đến dưới 90 điểm"),
    ("Khá", "_ Loại Khá: Từ 65 đến dưới 80 điểm"),
    ("Trung bình", "_ Loại Trung bình: Từ 50 đến dưới 65 điểm"),
    ("Yếu", "_ Loại Yếu: Từ 35 đến dưới 50 điểm"),
    ("Kém", "_ Loại kém: Dưới 35 điểm"),
]

for idx, (key, label) in enumerate(rating_ranges):
    r_idx = current_row + 3 + idx
    # label in cols 3-7 (merge C:G)
    sheet_out.cell(r_idx, 3, label).font = Font(name="Times New Roman", size=10)
    sheet_out.merge_cells(start_row=r_idx, start_column=3, end_row=r_idx, end_column=7)
    
    # count in col 9
    cnt = rating_counts[key]
    sheet_out.cell(r_idx, 9, cnt).font = Font(name="Times New Roman", size=10)
    sheet_out.cell(r_idx, 9).alignment = align_center
    
    # "Sinh viên" in col 10
    sheet_out.cell(r_idx, 10, "Sinh viên").font = Font(name="Times New Roman", size=10)
    
    # percentage in col 11
    pct = (cnt / len(processed_students)) * 100
    cell_pct = sheet_out.cell(r_idx, 11, pct)
    cell_pct.font = Font(name="Times New Roman", size=10)
    cell_pct.number_format = '0.00'
    cell_pct.alignment = align_center
    
    # "%" in col 12
    sheet_out.cell(r_idx, 12, "%").font = Font(name="Times New Roman", size=10)

# Signatures Block (Rows 61, 62, 67 relative to student rows)
sig_row1 = current_row + 10
sig_row2 = current_row + 11
sig_row3 = current_row + 16

# Left signature
sheet_out.cell(sig_row1, 1, "Khoa đào tạo").font = Font(name="Times New Roman", size=11, bold=True)
sheet_out.cell(sig_row1, 1).alignment = align_center
sheet_out.merge_cells(start_row=sig_row1, start_column=1, end_row=sig_row1, end_column=3)

sheet_out.cell(sig_row3, 1, "Công nghệ thông tin 2").font = Font(name="Times New Roman", size=11, bold=True)
sheet_out.cell(sig_row3, 1).alignment = align_center
sheet_out.merge_cells(start_row=sig_row3, start_column=1, end_row=sig_row3, end_column=3)

# Middle signature
sheet_out.cell(sig_row1, 5, "Cố vấn học tập").font = Font(name="Times New Roman", size=11, bold=True)
sheet_out.cell(sig_row1, 5).alignment = align_center
sheet_out.merge_cells(start_row=sig_row1, start_column=5, end_row=sig_row1, end_column=8)

sheet_out.cell(sig_row2, 5, "(Ký và ghi rõ họ tên)").font = Font(name="Times New Roman", size=10, italic=True)
sheet_out.cell(sig_row2, 5).alignment = align_center
sheet_out.merge_cells(start_row=sig_row2, start_column=5, end_row=sig_row2, end_column=8)

sheet_out.cell(sig_row3, 5, "Nguyễn Trung Hiếu").font = Font(name="Times New Roman", size=11, bold=True)
sheet_out.cell(sig_row3, 5).alignment = align_center
sheet_out.merge_cells(start_row=sig_row3, start_column=5, end_row=sig_row3, end_column=8)

# Right signature
sheet_out.cell(sig_row1, 9, "Lớp trưởng").font = Font(name="Times New Roman", size=11, bold=True)
sheet_out.cell(sig_row1, 9).alignment = align_center
sheet_out.merge_cells(start_row=sig_row1, start_column=9, end_row=sig_row1, end_column=12)

sheet_out.cell(sig_row2, 9, "(Ký và ghi rõ họ tên)").font = Font(name="Times New Roman", size=10, italic=True)
sheet_out.cell(sig_row2, 9).alignment = align_center
sheet_out.merge_cells(start_row=sig_row2, start_column=9, end_row=sig_row2, end_column=12)

sheet_out.cell(sig_row3, 9, "Ngô Trí Long").font = Font(name="Times New Roman", size=11, bold=True)
sheet_out.cell(sig_row3, 9).alignment = align_center
sheet_out.merge_cells(start_row=sig_row3, start_column=9, end_row=sig_row3, end_column=12)

# Set Column Widths
sheet_out.column_dimensions['A'].width = 5   # TT
sheet_out.column_dimensions['B'].width = 18  # Họ đệm
sheet_out.column_dimensions['C'].width = 10  # Tên
sheet_out.column_dimensions['D'].width = 15  # Mã sinh viên
sheet_out.column_dimensions['E'].width = 14  # Ngày sinh

# Subsections cols F to AF
for col_char in ['F','G','H','I','J','L','M','N','P','Q','R','S','T','V','W','X','Y','Z','AA','AC','AD','AE']:
    sheet_out.column_dimensions[col_char].width = 6

# Total cols and others
for col_char in ['K', 'O', 'U', 'AB', 'AF']:
    sheet_out.column_dimensions[col_char].width = 16
sheet_out.column_dimensions['AG'].width = 12  # Tổng điểm
sheet_out.column_dimensions['AH'].width = 16  # Xếp loại
sheet_out.column_dimensions['AI'].width = 15  # Ghi chú

dest_excel_path = os.path.join(workspace, "HV_Mau 2_Chi tiet KQRL.xlsx")
wb_out.save(dest_excel_path)
print(f"Successfully generated detailed Excel report at {dest_excel_path}")

# 7. Cleanup temp files
shutil.rmtree(temp_dir)
print("Temporary files cleaned up.")

# 8. Sync generated files to Google Drive if mounted
gdrive_dir = "/mnt/googledrive/generated_students"
if os.path.exists(gdrive_dir):
    print("\nSyncing generated student files to Google Drive...")
    try:
        # Clear existing files in the GDrive directory to prevent duplicates/obsolete entries
        for filename in os.listdir(gdrive_dir):
            file_path = os.path.join(gdrive_dir, filename)
            if os.path.isfile(file_path):
                os.remove(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        
        # Copy newly generated files
        for filename in os.listdir(output_dir):
            src_file = os.path.join(output_dir, filename)
            dest_file = os.path.join(gdrive_dir, filename)
            if os.path.isfile(src_file):
                shutil.copy2(src_file, dest_file)
        print("Sync to Google Drive completed successfully.")
    except Exception as e:
        print(f"Warning: Failed to sync with Google Drive: {e}")
else:
    print("\nNote: Google Drive mount point at /mnt/googledrive/generated_students not found. Skipping sync.")

print("Pipeline run finished successfully!")
