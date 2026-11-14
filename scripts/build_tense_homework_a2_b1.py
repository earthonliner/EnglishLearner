#!/usr/bin/env python3
"""Generate a compact two-page A2–B1 English tense worksheet and answer key."""

from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "worksheets"
WORKSHEET_PATH = OUT_DIR / "tense-practice-a2-b1.docx"
KEY_PATH = OUT_DIR / "tense-practice-a2-b1-answer-key.docx"

NAVY = "000000"
TEAL = "000000"
PALE_TEAL = "FFFFFF"
CORAL = "000000"
PALE_CORAL = "FFFFFF"
YELLOW = "FFFFFF"
INK = "000000"
MID = "000000"
WHITE = "FFFFFF"
LIGHT = "000000"


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=55, start=80, bottom=55, end=80):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for side, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{side}"))
        if node is None:
            node = OxmlElement(f"w:{side}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_cell_border(cell, **edges):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge, attrs in edges.items():
        tag = f"w:{edge}"
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        for key, value in attrs.items():
            element.set(qn(f"w:{key}"), str(value))


def set_repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def keep_paragraph(paragraph, keep_next=False):
    p_pr = paragraph._p.get_or_add_pPr()
    keep_lines = OxmlElement("w:keepLines")
    p_pr.append(keep_lines)
    if keep_next:
        node = OxmlElement("w:keepNext")
        p_pr.append(node)


def set_repeat_header(section):
    header = section.header
    header.is_linked_to_previous = False
    p = header.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)


def add_page_number(section):
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    run = p.add_run("Page ")
    run.font.name = "Times New Roman"
    run.font.size = Pt(9)
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    p._p.append(field)


def setup_document(doc):
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.25)
    section.bottom_margin = Cm(1.2)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)
    section.header_distance = Cm(0.45)
    section.footer_distance = Cm(0.45)
    set_repeat_header(section)
    add_page_number(section)

    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(10)
    normal.font.color.rgb = RGBColor.from_string(INK)
    normal.paragraph_format.space_after = Pt(2.5)
    normal.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

    for style_name in ("Title", "Subtitle"):
        style = doc.styles[style_name]
        style.font.name = "Times New Roman"


def add_title(doc, title, subtitle, score=True, identity=True):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(title)
    run.font.name = "Times New Roman"
    run.font.size = Pt(18)
    run.font.bold = True
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(subtitle)
    run.font.name = "Times New Roman"
    run.font.size = Pt(10.5)

    if identity:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(7)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run("Name 姓名：__________________    Date 日期：______________")
        r.font.size = Pt(10)
        if score:
            r = p.add_run("    Score 得分：______ / 40")


def add_section_heading(doc, label, title, meta):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(f"Part {label}.  {title}")
    r.bold = True
    r.font.size = Pt(11)
    if meta:
        r = p.add_run(f"  （{meta}）")
        r.font.size = Pt(9)
    keep_paragraph(p, keep_next=True)
    return p


def add_tense_bank(doc):
    table = doc.add_table(rows=2, cols=4)
    table.autofit = False
    width = Cm(4.55)
    items = [
        ("PRESENT SIMPLE", "usually / every day"),
        ("PRESENT CONT.", "now / Look!"),
        ("PAST SIMPLE", "yesterday / ago"),
        ("PAST CONT.", "at 8 p.m. yesterday"),
        ("FUTURE SIMPLE", "will + verb"),
        ("PRESENT PERFECT", "since / yet / already"),
        ("PAST PERFECT", "had + past participle"),
        ("PRESENT PERF. CONT.", "has/have been + -ing"),
    ]
    for i, (name, clue) in enumerate(items):
        cell = table.cell(i // 4, i % 4)
        cell.width = width
        set_cell_shading(cell, WHITE)
        set_cell_margins(cell, 50, 70, 50, 70)
        set_cell_border(cell, top={"val": "single", "sz": "3", "color": LIGHT},
                        bottom={"val": "single", "sz": "3", "color": LIGHT},
                        start={"val": "single", "sz": "3", "color": LIGHT},
                        end={"val": "single", "sz": "3", "color": LIGHT})
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(name + "\n")
        r.bold = True
        r.font.size = Pt(8)
        r = p.add_run(clue)
        r.font.size = Pt(8)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def add_numbered_items(doc, items, font_size=8.7, after=2.1):
    for num, text in enumerate(items, 1):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.15)
        p.paragraph_format.first_line_indent = Cm(-0.15)
        p.paragraph_format.space_after = Pt(after)
        r = p.add_run(f"{num}. ")
        r.bold = True
        r.font.color.rgb = RGBColor.from_string(CORAL)
        r = p.add_run(text)
        r.font.size = Pt(font_size)
        keep_paragraph(p)


def add_mc_grid(doc, items):
    table = doc.add_table(rows=4, cols=2)
    table.autofit = False
    for row in table.rows:
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        row.height = Cm(1.65)
    for i, text in enumerate(items):
        cell = table.cell(i // 2, i % 2)
        cell.width = Cm(9.1)
        set_cell_margins(cell, 55, 80, 45, 80)
        set_cell_shading(cell, WHITE)
        set_cell_border(cell, top={"val": "single", "sz": "3", "color": LIGHT},
                        bottom={"val": "single", "sz": "3", "color": LIGHT},
                        start={"val": "single", "sz": "3", "color": LIGHT},
                        end={"val": "single", "sz": "3", "color": LIGHT})
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(f"{i + 1}. ")
        r.bold = True
        r.font.color.rgb = RGBColor.from_string(TEAL)
        r = p.add_run(text)
        r.font.size = Pt(9.4)
        keep_paragraph(p)


def add_tip(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run("提示：")
    r.bold = True
    r = p.add_run(text)
    r.font.size = Pt(9)


def add_correction_grid(doc, items):
    table = doc.add_table(rows=4, cols=2)
    table.autofit = False
    for i, text in enumerate(items):
        cell = table.cell(i // 2, i % 2)
        cell.width = Cm(9.1)
        set_cell_margins(cell, 55, 75, 55, 75)
        set_cell_border(cell, bottom={"val": "single", "sz": "3", "color": LIGHT})
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(f"{i + 1}. ")
        r.bold = True
        r.font.color.rgb = RGBColor.from_string(CORAL)
        r = p.add_run(text)
        r.font.size = Pt(9.3)
        p = cell.add_paragraph("→ __________________________________________")
        p.paragraph_format.space_after = Pt(0)
        for run in p.runs:
            run.font.size = Pt(8.8)
            run.font.color.rgb = RGBColor.from_string(MID)


def add_passage(doc):
    table = doc.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    set_cell_shading(cell, WHITE)
    set_cell_margins(cell, 90, 110, 85, 110)
    set_cell_border(cell, top={"val": "single", "sz": "4", "color": TEAL},
                    bottom={"val": "single", "sz": "4", "color": TEAL},
                    start={"val": "single", "sz": "4", "color": TEAL},
                    end={"val": "single", "sz": "4", "color": TEAL})
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run("THE SCIENCE FAIR")
    r.bold = True
    r.font.size = Pt(10)
    r.font.color.rgb = RGBColor.from_string(NAVY)
    passage = (
        "Our science club usually 1. __________ (meet) on Friday afternoons. This term, we "
        "2. __________ (prepare) for the city science fair. We 3. __________ (work) on a small "
        "robot for six weeks, and we 4. __________ (already / complete) its body.\n\n"
        "Last Tuesday, while Mia 5. __________ (test) the motor, it suddenly 6. __________ (stop). "
        "We discovered that a wire 7. __________ (come) loose before the test began. Our teacher "
        "8. __________ (show) us how to repair it. Since then, the robot 9. __________ (pass) "
        "every test.\n\n"
        "Yesterday, two judges 10. __________ (visit) our school and asked several difficult "
        "questions. We 11. __________ (answer) all of them. The final is next Friday, and I think "
        "our robot 12. __________ (surprise) everyone."
    )
    p = cell.add_paragraph(passage)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.1
    for run in p.runs:
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor.from_string(INK)


def build_worksheet():
    doc = Document()
    setup_document(doc)
    add_title(doc, "MIXED TENSE CHALLENGE", "A2–B1 · 8 Tenses")
    add_tense_bank(doc)

    add_section_heading(
        doc, "A", "GRAMMAR IN CONTEXT",
        "根据完整语境填写正确时态。每空 1 分，共 12 分",
    )
    items_a = [
        "Evan normally __________ (work) from home, but this month he __________ (work) at the branch office.",
        "I __________ (wait) for the bus when I __________ (receive) a message saying it was delayed.",
        "Maya __________ (not finish) her report yet, although she started it yesterday.",
        "By the time the firefighters arrived, the neighbours __________ (lead) everyone outside.",
        "I __________ (lend) you my notes if you are absent tomorrow.",
        "We __________ (know) each other for years, but we only became classmates last term.",
        "How long __________ (you / prepare) for the speaking competition?",
        "During our holiday last year, we __________ (take) a night train to Beijing.",
        "This is the third time I __________ (read) this article, and I still notice something new.",
        "The days __________ (get) shorter now that winter is approaching.",
    ]
    add_numbered_items(doc, items_a, font_size=10, after=4.0)

    add_section_heading(
        doc, "B", "CHOOSE THE BEST MEANING",
        "结合动作状态与先后关系选择最佳答案。每题 1 分，共 8 分",
    )
    items_b = [
        "Lena moved here in 2021 and ___ here ever since.\nA lived  B has lived  C is living",
        "I can’t return the book; I ___ it and I’m only halfway through.\nA read  B have read  C have been reading",
        "When the principal entered, we ___ the problem.\nA discussed  B were discussing  C have discussed",
        "When the office closed, every student ___ the form.\nA submitted  B had submitted  C has submitted",
        "Whenever the alarm rings, the dog ___ under the table.\nA hides  B is hiding  C hid",
        "The sky is clear. I don’t think it ___ tonight.\nA rains  B rained  C will rain",
        "You can check your inbox; I ___ the file already.\nA sent  B have sent  C had sent",
        "Tom didn’t hear me because he ___ to music.\nA listened  B was listening  C has listened",
    ]
    add_mc_grid(doc, items_b)

    doc.add_page_break()

    add_title(doc, "MIXED TENSE CHALLENGE", "Page 2 · Apply & Check", score=False)
    add_tip(doc, "不要只看单个时间词；先画出动作时间线，再判断是结果、持续、背景动作还是先发生的动作。")
    doc.add_paragraph().paragraph_format.space_after = Pt(0)

    add_section_heading(
        doc, "C", "ERROR HUNT",
        "每句有 1 处时态或动词形式错误，改写正确。每题 1 分，共 8 分",
    )
    items_c = [
        "I am usually taking the underground to school.",
        "We have finished the experiment two hours ago.",
        "Ella has been knowing her best friend since Grade 1.",
        "By the time we arrived, the guests already went home.",
        "While I did my homework, my cousin called me.",
        "Leo has been completing three tasks today.",
        "Nora waits for the doctor since 4 p.m. and is still there.",
        "If the last bus has gone, I take a taxi home tonight.",
    ]
    add_correction_grid(doc, items_c)

    add_section_heading(
        doc, "D", "MIXED-TENSE CLOZE",
        "根据整篇语境判断时态。每空 1 分，共 12 分",
    )
    add_passage(doc)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(0)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("SELF-CHECK  □ I circled time clues.   □ I checked irregular verbs.   □ I checked has / have / had.")
    r.bold = True
    r.font.size = Pt(8)
    r.font.color.rgb = RGBColor.from_string(TEAL)

    doc.core_properties.title = "A2–B1 Tense Training: 8 Tenses"
    doc.core_properties.subject = "English grammar homework worksheet"
    doc.core_properties.author = "EnglishLearner"
    doc.save(WORKSHEET_PATH)


def add_key_section(doc, label, title, body):
    add_section_heading(doc, label, title, "")
    p = doc.add_paragraph(body)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.line_spacing = 1.08
    for run in p.runs:
        run.font.size = Pt(9)


def build_answer_key():
    doc = Document()
    setup_document(doc)
    add_title(doc, "ANSWER KEY", "Mixed Tense Challenge · Teacher Copy", score=False, identity=False)

    add_key_section(
        doc, "A", "GRAMMAR IN CONTEXT",
        "1. works; is working   2. was waiting; received   3. has not / hasn’t finished\n"
        "4. had led   5. will lend   6. have known   7. have you been preparing\n"
        "8. took   9. have read   10. are getting",
    )
    add_key_section(
        doc, "B", "CHOOSE THE BEST MEANING",
        "1. B   2. C   3. B   4. B   5. A   6. C   7. B   8. B",
    )
    add_key_section(
        doc, "C", "ERROR HUNT",
        "1. I usually take the underground to school.\n"
        "2. We finished the experiment two hours ago.\n"
        "3. Ella has known her best friend since Grade 1.\n"
        "4. By the time we arrived, the guests had already gone home.\n"
        "5. While I was doing my homework, my cousin called me.\n"
        "6. Leo has completed three tasks today.\n"
        "7. Nora has been waiting for the doctor since 4 p.m. and is still there.\n"
        "8. If the last bus has gone, I will take a taxi home tonight.",
    )
    add_key_section(
        doc, "D", "MIXED-TENSE CLOZE",
        "1. meets   2. are preparing   3. have been working   4. have already completed\n"
        "5. was testing   6. stopped   7. had come   8. showed\n"
        "9. has passed   10. visited   11. answered   12. will surprise",
    )

    table = doc.add_table(rows=1, cols=1)
    cell = table.cell(0, 0)
    set_cell_shading(cell, YELLOW)
    set_cell_margins(cell, 80, 100, 80, 100)
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("Teacher note")
    r.bold = True
    r.font.color.rgb = RGBColor.from_string(CORAL)
    p = cell.add_paragraph(
        "Key contrasts: present perfect shows a result or completed number; present perfect continuous "
        "emphasises duration or an unfinished activity. Stative verbs such as know normally do not use "
        "continuous forms. Past perfect marks the earlier of two past actions."
    )
    p.paragraph_format.space_after = Pt(0)
    for run in p.runs:
        run.font.size = Pt(8.3)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(
        "Quick diagnosis: A/B errors → compare the meaning of both actions. "
        "C errors → review tense contrasts and stative verbs. "
        "D errors → draw one timeline for the whole passage before choosing forms."
    )
    r.font.size = Pt(8.2)
    r.font.color.rgb = RGBColor.from_string(MID)

    doc.core_properties.title = "Answer Key: A2–B1 Tense Training"
    doc.core_properties.subject = "Teacher answer key"
    doc.core_properties.author = "EnglishLearner"
    doc.save(KEY_PATH)


if __name__ == "__main__":
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    build_worksheet()
    build_answer_key()
    print(WORKSHEET_PATH)
    print(KEY_PATH)
