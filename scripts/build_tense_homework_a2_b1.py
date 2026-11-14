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
    r = p.add_run("OUR CHANGING CITY")
    r.bold = True
    r.font.size = Pt(10)
    r.font.color.rgb = RGBColor.from_string(NAVY)
    passage = (
        "Our class usually 1. __________ (do) one big project each term. Last Monday, "
        "our teacher 2. __________ (announce) a new project called “Our Changing City.” "
        "We 3. __________ (work) on it this week.\n\n"
        "Before our first group meeting yesterday, Lily 4. __________ already __________ "
        "(collect) ten old photos. At 4 p.m., while I 5. __________ (interview) my grandfather, "
        "my teammates 6. __________ (look) for information online. Grandpa 7. __________ (live) "
        "in Shenzhen since 1985, so he knows how the city has changed. He 8. __________ (tell) "
        "me some wonderful stories.\n\n"
        "Since Tuesday, our group 9. __________ (prepare) a poster. We 10. __________ not "
        "__________ (finish) it yet, but it already looks great. I think we 11. __________ "
        "(complete) it tomorrow. Our teacher 12. __________ (help) us with the final part now."
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
    add_title(doc, "TENSE TRAINING", "8 Tenses · Mixed Practice")
    add_tense_bank(doc)

    add_section_heading(
        doc, "A", "ONE SENTENCE, ONE CLUE",
        "用括号内动词的正确形式填空。每空 1 分，共 12 分",
    )
    items_a = [
        "Leo usually __________ (practice) basketball after school.",
        "Listen! Someone __________ (knock) at the door.",
        "We __________ (visit) the science museum last Saturday.",
        "I think our team __________ (win) the match tomorrow.",
        "At 8 p.m. yesterday, I __________ (do) my homework.",
        "Amy __________ (know) her best friend since primary school.",
        "By the time we reached the station, the train __________ (leave).",
        "Ben is tired because he __________ (run) for forty minutes.",
        "Mum __________ not __________ (finish) dinner yet.",
        "While I __________ (walk) home, it suddenly began to rain.",
        "My brother __________ (read) this book twice.",
        "Sara normally takes the bus, but today she __________ (cycle) to school.",
    ]
    add_numbered_items(doc, items_a, font_size=10, after=4.0)

    add_section_heading(
        doc, "B", "CHOOSE THE BEST TIMELINE",
        "选择最佳答案。每题 1 分，共 8 分",
    )
    items_b = [
        "When we arrived, the film ___.\nA starts  B had started  C has started",
        "Nina ___ for the bus since 9:00.\nA waits  B waited  C has been waiting",
        "I ___ my keys yesterday.\nA lose  B lost  C have lost",
        "Look! The dog ___ after the ball.\nA runs  B ran  C is running",
        "Dad ___ coffee every morning.\nA drinks  B is drinking  C drank",
        "At midnight, the baby ___ .\nA slept  B was sleeping  C has slept",
        "We ___ that museum three times.\nA visit  B visited  C have visited",
        "Don’t worry. I ___ you tonight.\nA call  B called  C will call",
    ]
    add_mc_grid(doc, items_b)

    doc.add_page_break()

    add_title(doc, "TENSE TRAINING", "Page 2 · Apply & Check", score=False)
    add_tip(doc, "先圈出时间线索，再判断动作是习惯、进行中、已完成，还是先于另一个过去动作。")
    doc.add_paragraph().paragraph_format.space_after = Pt(0)

    add_section_heading(
        doc, "C", "ERROR HUNT",
        "每句有 1 处时态或动词形式错误，改写正确。每题 1 分，共 8 分",
    )
    items_c = [
        "She go to the library every Friday.",
        "Be quiet! The baby sleeps now.",
        "I have seen Leo yesterday afternoon.",
        "When the lights went out, we are watching TV.",
        "By the time I got home, Dad has cooked dinner.",
        "Sam has practiced the piano for two hours and he is still playing.",
        "I think people live on Mars one day.",
        "We didn’t went to the park because it rained.",
    ]
    add_correction_grid(doc, items_c)

    add_section_heading(
        doc, "D", "STORY MODE: MIX THE TENSES",
        "根据时间线与上下文填空。每空 1 分，共 12 分",
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
    add_title(doc, "ANSWER KEY", "Tense Training · Teacher Copy", score=False, identity=False)

    add_key_section(
        doc, "A", "ONE SENTENCE, ONE CLUE",
        "1. practices   2. is knocking   3. visited   4. will win\n"
        "5. was doing   6. has known   7. had left   8. has been running\n"
        "9. has not / hasn’t finished   10. was walking   11. has read   12. is cycling",
    )
    add_key_section(
        doc, "B", "CHOOSE THE BEST TIMELINE",
        "1. B   2. C   3. B   4. C   5. A   6. B   7. C   8. C",
    )
    add_key_section(
        doc, "C", "ERROR HUNT",
        "1. She goes to the library every Friday.\n"
        "2. Be quiet! The baby is sleeping now.\n"
        "3. I saw Leo yesterday afternoon.\n"
        "4. When the lights went out, we were watching TV.\n"
        "5. By the time I got home, Dad had cooked dinner.\n"
        "6. Sam has been practicing the piano for two hours and he is still playing.\n"
        "7. I think people will live on Mars one day.\n"
        "8. We didn’t go to the park because it rained.",
    )
    add_key_section(
        doc, "D", "STORY MODE: MIX THE TENSES",
        "1. does   2. announced   3. are working   4. had, collected\n"
        "5. was interviewing   6. were looking   7. has lived   8. told\n"
        "9. has been preparing   10. have not / haven’t finished   11. will complete\n"
        "12. is helping",
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
        "Present perfect focuses on a result or life experience; present perfect continuous focuses "
        "on duration or an activity still in progress. In British English, “practising” is also an "
        "acceptable spelling of “practicing.”"
    )
    p.paragraph_format.space_after = Pt(0)
    for run in p.runs:
        run.font.size = Pt(8.3)

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(
        "Quick diagnosis:  A/B errors → identify time clues first.  "
        "C errors → review form rules (third-person -s, did + base verb).  "
        "D errors → draw a timeline and compare actions."
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
