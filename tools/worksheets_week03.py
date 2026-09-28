# Week 3 printable worksheets: tools/worksheets/*.html -> PDF (A4) in the week-3 folder.
# Needs: pip install playwright && playwright install chromium   (Chromium prints the HTML exactly like the browser does)
#   python worksheets_week03.py
import os
import re

from common import CREDIT, REPO
from playwright.sync_api import sync_playwright

SRC = os.path.join(REPO, "tools", "worksheets")
OUT = os.path.join(REPO, "Grade-6_Week-03_2026-09-28--2026-10-02_English_Revision-Part-2")

# (source html, pdf name, expected pages)
SHEETS = [
    ("week03_ws1_neighbourhood.html", "Grade-6_Week-03_English_Worksheet-1-Neighbourhood-City.pdf", 2),
    ("week03_ws2_food.html", "Grade-6_Week-03_English_Worksheet-2-Food.pdf", 2),
    ("week03_ws3_nature_future.html", "Grade-6_Week-03_English_Worksheet-3-Nature-Future.pdf", 2),
    ("week03_quiz.html", "Grade-6_Week-03_English_Worksheet-Revision-Quiz.pdf", 1),
    ("week03_answer_key.html", "Grade-6_Week-03_English_Worksheets-Answer-Key.pdf", 1),
]

FOOTER = ('<div style="width:100%;font-size:7px;text-align:center;color:#444;font-family:sans-serif">'
          + CREDIT + ' · page <span class="pageNumber"></span>/<span class="totalPages"></span></div>')

with sync_playwright() as p:
    browser = p.chromium.launch()
    page = browser.new_page()
    for src, name, expected in SHEETS:
        page.goto("file:///" + os.path.join(SRC, src).replace("\\", "/"))
        page.wait_for_load_state("networkidle")
        pdf = os.path.join(OUT, name)
        page.pdf(path=pdf, format="A4", prefer_css_page_size=True, print_background=True,
                 display_header_footer=True, header_template="<div></div>", footer_template=FOOTER)
        with open(pdf, "rb") as f:
            pages = len(re.findall(rb"/Type\s*/Page[^s]", f.read()))
        print(("ok   " if pages == expected else "CHECK") + f" {name}: {pages} page(s), expected {expected}")
    browser.close()
