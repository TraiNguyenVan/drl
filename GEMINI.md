things you must know:
1. you will not modify HV_Mau 2_Tong hop KQRL cua SV.xlsx , scores here always true
2. master.docx is the original template (main branch). master_v2.docx is the active template on the google-docs-compat branch — it has no floating textboxes and uses borderless tables for the info and signature sections. read STRUCTURE.md to understand the structure of whichever template is active
3. whenever i need to create more files like master.docx, you must make a copy of the active template (master_v2.docx on this branch), edit right on it to preserve intended format
4. ./students is where raw student papers live, they can be doc or docx, and may not have the right format as the template, but we may extract scores from here whenever we need to generate new files
5. the pipeline script is process_and_generate.py — run it with python3 process_and_generate.py to regenerate all 37 student files and the Excel report
6. table indices in the generated files (master_v2.docx): tables[0]=Logo, tables[1]=Info (name/DOB/MSV), tables[2]=Grading criteria, tables[3]=Signature