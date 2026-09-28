# tools – how the plans are built

Python + openpyxl build the `.xlsx`; Excel (via pywin32) prints it to PDF. Run from this folder:

```
python annual_plan.py                 # school copy (signed) -> Desktop
python annual_plan.py --public        # public copy -> repo folder
python lesson_plan_week03.py          # same for the week-3 lesson plan
python lesson_plan_week03.py --public
```

- **School copy:** signature block (teacher / approval). The names come from `local_signatures.json`, which is in `.gitignore` and never uploaded. Keys: `teacher_name`, `teacher_title`, `approval_word`, `principal_name`, `principal_title`.
- **Public copy:** no signatures; a credit line and a page footer: "Prepared by Ogün Aydın · CC BY 4.0 …" (`common.py`, `CREDIT`).
- **New week:** copy the latest `lesson_plan_weekNN.py`, change the folder/file names in `out_path(...)`, the dates, lesson times and content, then build both copies. Check the PDF is one page.
- Sources for content: the annual plan (week → theme) and the MoNE curriculum PDF (outcomes, vocabulary, structures).
- Never put student names, photos or class lists in anything here.
