# Builds the compact 6th grade English annual plan (2026-2027) from the TYMM framework sheet.
import math

from common import PUBLIC, export_pdf, finish, out_path
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

OUT = out_path("Grade-6_Annual-Plan_2026-2027_English",
               "Grade-6_English_Annual-Plan_2026-2027.xlsx",
               "2026-2027_6AB_Ingilizce_Yillik_Plan.xlsx")

CODES = "ENG.6.{n}. L1–L4, P1, R1–R4, V1, G1, W1–W6, S1–S6 (23)"
THEMES = {
    1: ("THEME 1: SCHOOL LIFE", "Roles and responsibilities at school; school routines; national days and celebrations",
        "Poster about their daily routines at school (sentences + pictures), presented and displayed in class."),
    2: ("THEME 2: CLASSROOM LIFE", "Daily and study routines; learning activities in the classroom; cardinal numbers 100–500; ordinal numbers 1–50",
        "A numbers game (board/card game or quiz) with 100–500 and 1st–50th and classroom-life questions; presented and played in class."),
    3: ("THEME 3: PERSONAL LIFE", "Body parts, physical appearance and clothes; personality and character",
        "Personal profile poster: labelled body parts, appearance, favourite clothes, personality."),
    4: ("THEME 4: FAMILY LIFE", "Family members’ jobs, working places and job routines; different types of family homes and houses",
        "Short booklet about a family member’s job, workplace, daily routine and home, with a drawing of the home."),
    5: ("THEME 5: LIFE IN THE NEIGHBOURHOOD & CITY", "Festivals and events (sports, music, arts) in the neighbourhood and city; transportation",
        "Poster or short presentation about a local festival/event and two ways to travel there."),
    6: ("THEME 6: LIFE IN THE WORLD & CULTURE", "Countries, nationalities and languages; food types and events from different parts of the world",
        "Country poster: nationality, language, a traditional food and a festival where it is served."),
    7: ("THEME 7: LIFE IN NATURE & GLOBAL PROBLEMS", "Activities in nature; environmental problems in the world and solutions",
        "Short report on an environmental problem with solutions, read aloud in class."),
    8: ("THEME 8: LIFE IN THE UNIVERSE & FUTURE", "Planets and the Earth as a planet; life on Earth in the future",
        "6–8 panel comic strip about the planets and life on Earth in the future (Web 2.0 tools optional)."),
}
SBP1 = ("SCHOOL-BASED PLANNING: “Our school in English”: English signs for school areas (Themes 1–2 vocabulary); "
        "10 November: short English text on Atatürk and education + class poster; numbers game (100–500, 1st–50th).")
SBP2 = ("SCHOOL-BASED PLANNING: “Saimbeyli in English” local study: a local food, place or festival in simple English (Theme 6); "
        "“Countries corner” with the Theme 6 posters; a short English greeting/song for 23 April.")
SOC = "SOCIAL ACTIVITIES: English games day; exhibition of the year’s performance tasks (posters, booklets, comic strips)."

# (month, week, dates, kind, value, special days)  kind: T=theme n, X=free text row spanning E:G, H=holiday row
ROWS = [
    ("SEPTEMBER", 1, "14–18 Sep", "X", "ORIENTATION", "15 July Democracy and National Unity Day; Primary Education Week"),
    ("", 2, "21–25 Sep", "X", "REVISION", "Turkish Language Day (26 Sep)"),
    ("", 3, "28 Sep–2 Oct", "X", "REVISION", ""),
    ("OCTOBER", 4, "5–9 Oct", "T", 1, ""),
    ("", 5, "12–16 Oct", "T", 1, ""),
    ("", 6, "19–23 Oct", "T", 1, ""),
    ("", 7, "26–30 Oct", "T", 1, "29 October Republic Day"),
    ("NOVEMBER", 8, "2–6 Nov", "T", 2, ""),
    ("", 9, "9–13 Nov", "X", SBP1, "10 November Atatürk Memorial Day"),
    ("H", 0, "", "H", "FIRST MIDTERM BREAK (16–20 November 2026)", ""),
    ("", 10, "23–27 Nov", "T", 2, "20 November Children’s Rights Day; 24 November Teachers’ Day"),
    ("", 11, "30 Nov–4 Dec", "T", 2, "3 December International Day of Persons with Disabilities"),
    ("DECEMBER", 12, "7–11 Dec", "T", 2, ""),
    ("", 13, "14–18 Dec", "T", 3, ""),
    ("", 14, "21–25 Dec", "T", 3, ""),
    ("", 15, "28 Dec–1 Jan", "T", 3, "1 January New Year"),
    ("JANUARY", 16, "4–8 Jan", "T", 3, ""),
    ("", 17, "11–15 Jan", "T", 4, ""),
    ("", 18, "18–22 Jan", "T", 4, ""),
    ("H", 0, "", "H", "SEMESTER HOLIDAY (25 January–5 February 2027)", ""),
    ("FEBRUARY", 19, "8–12 Feb", "T", 4, ""),
    ("", 20, "15–19 Feb", "T", 4, ""),
    ("", 21, "22–26 Feb", "T", 5, ""),
    ("MARCH", 22, "1–5 Mar", "T", 5, ""),
    ("H", 0, "", "H", "SECOND MIDTERM BREAK (8–12 March 2027) · EID AL-FITR (9–11 March)", ""),
    ("", 23, "15–19 Mar", "T", 5, "12 March Adoption of the National Anthem and Mehmet Akif Ersoy Commemoration Day; 18 March Çanakkale Victory and Martyrs’ Day; Turkic World and Communities Week"),
    ("", 24, "22–26 Mar", "T", 5, ""),
    ("", 25, "29 Mar–2 Apr", "T", 6, "Library Week"),
    ("APRIL", 26, "5–9 Apr", "T", 6, ""),
    ("", 27, "12–16 Apr", "X", SBP2, ""),
    ("", 28, "19–23 Apr", "T", 6, "23 April National Sovereignty and Children’s Day"),
    ("", 29, "26–30 Apr", "T", 7, "29 April Kût’ül Amâre Victory; 1 May Labour Day"),
    ("MAY", 30, "3–7 May", "T", 7, ""),
    ("", 31, "10–14 May", "T", 7, ""),
    ("H", 0, "", "H", "EID AL-ADHA (16–19 May 2027)", ""),
    ("", 32, "17–21 May", "T", 7, "19 May Commemoration of Atatürk, Youth and Sports Day"),
    ("", 33, "24–28 May", "T", 8, "29 May Conquest of İstanbul"),
    ("JUNE", 34, "31 May–4 Jun", "T", 8, ""),
    ("", 35, "7–11 Jun", "T", 8, ""),
    ("", 36, "14–18 Jun", "T", 8, ""),
    ("", 37, "21–25 Jun", "X", SOC, ""),
]

HEAD = ["MONTH", "WEEK", "DATES", "HOURS", "THEME AND SUB-THEMES", "LEARNING OUTCOMES", "PERFORMANCE TASK (rubric / rating scale)", "NATIONAL CELEBRATIONS AND IMPORTANT DAYS"]
WIDTH = [10, 5.5, 12, 6, 36, 22, 40, 26]
FONT = 7.5
LINE_PT = 9.6       # height of one text line at FONT
CHARS_PER_UNIT = 1.25  # characters that fit per column width unit at FONT

thin = Side(style="thin", color="808080")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
GREY = PatternFill("solid", fgColor="D9D9D9")
LIGHT = PatternFill("solid", fgColor="F2F2F2")
F = Font(name="Calibri", size=FONT)
FB = Font(name="Calibri", size=FONT, bold=True)
WRAP = Alignment(wrap_text=True, vertical="center")
CWRAP = Alignment(wrap_text=True, vertical="center", horizontal="center")


def lines(text, width):
    if not text:
        return 1
    cap = max(1, int(width * CHARS_PER_UNIT))
    return sum(max(1, math.ceil(len(part) / cap)) for part in str(text).split("\n"))


wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Annual Plan"
for i, w in enumerate(WIDTH, 1):
    ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

ws.merge_cells("A1:H1")
ws["A1"] = "SAİMBEYLİ ŞEHİT BAYRAM ALİ YARDIMCI YATILI BÖLGE ORTAOKULU · 2026-2027 ACADEMIC YEAR"
ws.merge_cells("A2:H2")
ws["A2"] = "YEAR 6 ENGLISH ANNUAL PLAN  (6/A – 6/B · 3 hours a week · 108 hours · Level: A2.2)"
for c in ("A1", "A2"):
    ws[c].font = Font(name="Calibri", size=10, bold=True)
    ws[c].alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[1].height = 14
ws.row_dimensions[2].height = 14
for i, h in enumerate(HEAD, 1):
    c = ws.cell(3, i, h)
    c.font, c.alignment, c.border, c.fill = FB, CWRAP, BOX, GREY
ws.row_dimensions[3].height = 22
ws.print_title_rows = "3:3"

r0 = 4
rows_at = []
for k, (month, wk, dates, kind, val, special) in enumerate(ROWS):
    r = r0 + k
    rows_at.append(r)
    if kind == "H":
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        c = ws.cell(r, 1, val)
        c.font, c.alignment, c.fill = FB, CWRAP, LIGHT
        for col in range(1, 9):
            ws.cell(r, col).border = BOX
        ws.row_dimensions[r].height = 11
        continue
    vals = [month, f"{wk}.", dates, 3, None, None, None, special]
    for col, v in enumerate(vals, 1):
        c = ws.cell(r, col, v)
        c.font, c.border = F, BOX
        c.alignment = CWRAP if col in (1, 2, 4) else WRAP
    ws.cell(r, 1).font = FB
    if kind == "X":
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=7)
        ws.cell(r, 5, val).font = FB if val in ("ORIENTATION", "REVISION") else F
        need = lines(val, WIDTH[4] + WIDTH[5] + WIDTH[6])
        ws.row_dimensions[r].height = max(11, need * LINE_PT + 2, lines(special, WIDTH[7]) * LINE_PT + 2)

# theme blocks: consecutive rows with the same theme number
blocks = []
for k, row in enumerate(ROWS):
    if row[3] != "T":
        continue
    if blocks and blocks[-1][0] == row[4] and blocks[-1][2] == k - 1:
        blocks[-1][2] = k
    else:
        blocks.append([row[4], k, k])
seen = set()
for n, a, b in blocks:
    ra, rb = r0 + a, r0 + b
    name, subs, task = THEMES[n]
    first = n not in seen
    seen.add(n)
    texts = {
        5: f"{name}\n{subs}" if first else f"{name} (continued)",
        6: CODES.format(n=n) if first else f"ENG.6.{n}. (continued)",
        7: task if first else "(continued)",
    }
    need = 1
    for col, t in texts.items():
        if ra != rb:
            ws.merge_cells(start_row=ra, start_column=col, end_row=rb, end_column=col)
        c = ws.cell(ra, col, t)
        c.font, c.alignment = F, WRAP
        need = max(need, lines(t, WIDTH[col - 1]))
    span = rb - ra + 1
    per_row = max(11, (need * LINE_PT + 3) / span)
    for r in range(ra, rb + 1):
        per_row_r = max(per_row, lines(ws.cell(r, 8).value, WIDTH[7]) * LINE_PT + 2)
        ws.row_dimensions[r].height = per_row_r
        for col in range(1, 9):
            ws.cell(r, col).border = BOX

end = r0 + len(ROWS)
ws.merge_cells(start_row=end, start_column=1, end_row=end, end_column=8)
note = ws.cell(end, 1, (
    "This plan is based on the English Language Curriculum (Years 2–8) annexed to Board of Education Decision No. 45 of 24.07.2025 "
    "(Year 6: A2.2), the TYMM Common Text of Curricula, the MoNE 2026-2027 framework plan and academic calendar, and Directive "
    "No. 58168473 of 19.09.2022. Performance tasks are summarised; their full texts are in the framework plan."))
note.font = Font(name="Calibri", size=7, italic=True)
note.alignment = WRAP
ws.row_dimensions[end].height = 20

s = end + 2
last = finish(ws, s, (1, 4), (7, 8), "28.09.2026")

ps = ws.page_setup
ps.paperSize = ws.PAPERSIZE_A4
ps.orientation = "landscape"
ws.sheet_properties.pageSetUpPr.fitToPage = True
ps.fitToWidth = 1
ps.fitToHeight = 1
ws.page_margins.left = ws.page_margins.right = 0.35
ws.page_margins.top = ws.page_margins.bottom = 0.3
ws.print_options.horizontalCentered = True
ws.print_area = f"A1:H{last}"
wb.save(OUT)
print("saved", OUT)
export_pdf(OUT)
