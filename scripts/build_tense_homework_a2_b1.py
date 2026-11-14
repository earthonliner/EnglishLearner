#!/usr/bin/env python3
"""Generate a clean, black-and-white school-exam style A2-B1 tense homework worksheet (<= 2 pages) and teacher answer key."""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "worksheets"
WORKSHEET_PATH = OUT_DIR / "tense-practice-a2-b1.docx"
KEY_PATH = OUT_DIR / "tense-practice-a2-b1-answer-key.docx"

BLACK = RGBColor(0, 0, 0)
LATIN_FONT = "Times New Roman"
CJK_FONT = "Noto Serif CJK SC"


def apply_run_fonts(run, size_pt: float, bold: bool = False, italic: bool = False) -> None:
    run.font.name = LATIN_FONT
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = BLACK
    r_pr = run._r.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, r_fonts)
    r_fonts.set(qn("w:ascii"), LATIN_FONT)
    r_fonts.set(qn("w:hAnsi"), LATIN_FONT)
    r_fonts.set(qn("w:cs"), LATIN_FONT)
    r_fonts.set(qn("w:eastAsia"), CJK_FONT)


def set_paragraph_bottom_border(paragraph, sz: int = 6) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(sz))
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), "000000")
    p_bdr.append(bottom)
    p_pr.append(p_bdr)


def keep_together(paragraph, keep_with_next: bool = False) -> None:
    paragraph.paragraph_format.keep_together = True
    if keep_with_next:
        paragraph.paragraph_format.keep_with_next = True


def append_field(paragraph, field_name: str, size_pt: float = 9.0) -> None:
    run_begin = paragraph.add_run()
    apply_run_fonts(run_begin, size_pt)
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    run_begin._r.append(fld_begin)

    run_instr = paragraph.add_run()
    apply_run_fonts(run_instr, size_pt)
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {field_name} "
    run_instr._r.append(instr)

    run_sep = paragraph.add_run()
    apply_run_fonts(run_sep, size_pt)
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    run_sep._r.append(fld_sep)

    run_val = paragraph.add_run("1")
    apply_run_fonts(run_val, size_pt)

    run_end = paragraph.add_run()
    apply_run_fonts(run_end, size_pt)
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run_end._r.append(fld_end)


def configure_page(doc: Document, show_total_pages: bool = True) -> None:
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.45)
    section.left_margin = Cm(1.7)
    section.right_margin = Cm(1.7)
    section.header_distance = Cm(0.8)
    section.footer_distance = Cm(0.8)

    header_p = section.header.paragraphs[0]
    header_p.text = ""

    footer_p = section.footer.paragraphs[0]
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_p.paragraph_format.space_before = Pt(0)
    footer_p.paragraph_format.space_after = Pt(0)
    r1 = footer_p.add_run("第 ")
    apply_run_fonts(r1, 9.0)
    append_field(footer_p, "PAGE", 9.0)
    if show_total_pages:
        r2 = footer_p.add_run(" 页（共 ")
        apply_run_fonts(r2, 9.0)
        append_field(footer_p, "NUMPAGES", 9.0)
        r3 = footer_p.add_run(" 页）")
        apply_run_fonts(r3, 9.0)
    else:
        r2 = footer_p.add_run(" 页")
        apply_run_fonts(r2, 9.0)

    normal = doc.styles["Normal"]
    normal.font.name = LATIN_FONT
    normal.font.size = Pt(10)
    normal.font.color.rgb = BLACK
    r_pr = normal.element.get_or_add_rPr()
    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.insert(0, r_fonts)
    r_fonts.set(qn("w:ascii"), LATIN_FONT)
    r_fonts.set(qn("w:hAnsi"), LATIN_FONT)
    r_fonts.set(qn("w:cs"), LATIN_FONT)
    r_fonts.set(qn("w:eastAsia"), CJK_FONT)


def add_exam_header(doc: Document, title: str, subtitle: str | None = None, student_fields: bool = True) -> None:
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(3 if subtitle or student_fields else 6)
    r_title = p_title.add_run(title)
    apply_run_fonts(r_title, 14.5, bold=True)
    keep_together(p_title, keep_with_next=True)

    if subtitle:
        p_sub = doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_sub.paragraph_format.space_before = Pt(0)
        p_sub.paragraph_format.space_after = Pt(4)
        r_sub = p_sub.add_run(subtitle)
        apply_run_fonts(r_sub, 9.5)
        if not student_fields:
            set_paragraph_bottom_border(p_sub, sz=6)
        keep_together(p_sub, keep_with_next=True)

    if student_fields:
        p_info = doc.add_paragraph()
        p_info.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_info.paragraph_format.space_before = Pt(0)
        p_info.paragraph_format.space_after = Pt(6)
        r_info = p_info.add_run(
            "姓名：____________　　班级：____________　　日期：____________　　得分：______ / 50"
        )
        apply_run_fonts(r_info, 9.5)
        set_paragraph_bottom_border(p_info, sz=6)
        keep_together(p_info, keep_with_next=True)


def add_section_title(doc: Document, heading: str, instruction: str | None = None, space_before: float = 5.0) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(1.5 if instruction else 3.0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    p.paragraph_format.line_spacing = 1.15
    r = p.add_run(heading)
    apply_run_fonts(r, 10.5, bold=True)
    keep_together(p, keep_with_next=True)

    if instruction:
        p_inst = doc.add_paragraph()
        p_inst.paragraph_format.space_before = Pt(0)
        p_inst.paragraph_format.space_after = Pt(3.0)
        p_inst.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        p_inst.paragraph_format.line_spacing = 1.15
        r_inst = p_inst.add_run(instruction)
        apply_run_fonts(r_inst, 9.5)
        keep_together(p_inst, keep_with_next=True)


def add_mcq_item(doc: Document, idx: int, stem_lines: list[str], options: tuple[str, str, str, str]) -> None:
    for line_idx, line in enumerate(stem_lines):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.75)
        p.paragraph_format.first_line_indent = Cm(-0.75 if line_idx == 0 else 0.0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1.0)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        p.paragraph_format.line_spacing = 1.18
        prefix = f"(    ) {idx}. " if line_idx == 0 else ""
        r = p.add_run(prefix + line)
        apply_run_fonts(r, 9.8)
        keep_together(p, keep_with_next=True)

    p_opt = doc.add_paragraph()
    p_opt.paragraph_format.left_indent = Cm(0.75)
    p_opt.paragraph_format.space_before = Pt(0)
    p_opt.paragraph_format.space_after = Pt(2.8)
    p_opt.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    p_opt.paragraph_format.line_spacing = 1.15
    tabs = p_opt.paragraph_format.tab_stops
    tabs.add_tab_stop(Cm(4.95), WD_TAB_ALIGNMENT.LEFT)
    tabs.add_tab_stop(Cm(9.25), WD_TAB_ALIGNMENT.LEFT)
    tabs.add_tab_stop(Cm(13.55), WD_TAB_ALIGNMENT.LEFT)

    labels = ("A", "B", "C", "D")
    opt_text = "\t".join(f"{label}. {opt}" for label, opt in zip(labels, options))
    r_opt = p_opt.add_run(opt_text)
    apply_run_fonts(r_opt, 9.8)
    keep_together(p_opt)


def add_fill_item(doc: Document, idx: int, text: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.55)
    p.paragraph_format.first_line_indent = Cm(-0.55)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2.8)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    p.paragraph_format.line_spacing = 1.25
    r = p.add_run(f"{idx}. {text}")
    apply_run_fonts(r, 9.8)
    keep_together(p)


def add_error_correction_item(doc: Document, idx: int, sentence: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.55)
    p.paragraph_format.first_line_indent = Cm(-0.55)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3.5)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    p.paragraph_format.line_spacing = 1.25
    p.paragraph_format.tab_stops.add_tab_stop(Cm(17.6), WD_TAB_ALIGNMENT.RIGHT)
    r = p.add_run(f"{idx}. {sentence}\t____________________")
    apply_run_fonts(r, 9.8)
    keep_together(p)


def add_cloze_paragraph(doc: Document, segments: list[tuple[str, bool]]) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0.74)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4.5)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    p.paragraph_format.line_spacing = 1.55
    for text, is_bold in segments:
        r = p.add_run(text)
        apply_run_fonts(r, 10.0, bold=is_bold)


def build_student_worksheet() -> None:
    doc = Document()
    configure_page(doc, show_total_pages=True)

    add_exam_header(
        doc,
        title="英语动词时态综合练习（A2–B1）",
        student_fields=True,
    )

    # Section 1: Multiple Choice (10 items, 10 pts)
    add_section_title(
        doc,
        "一、单项选择（每小题 1 分，共 10 分）",
        space_before=2.0,
    )

    mcqs = [
        (
            [
                "— Your eyes look quite red, Kate. Are you alright?",
                "— I'm fine. I ________ onions in the kitchen for the last ten minutes.",
            ],
            ("chop", "chopped", "have been chopping", "had chopped"),
        ),
        (
            [
                "When we reached the campsite at dusk, we found that another group ________ the flattest spot by the river.",
            ],
            ("takes", "took", "has taken", "had taken"),
        ),
        (
            [
                "The museum guide says the tour won't begin until every visitor ________ a safety headset.",
            ],
            ("wears", "will wear", "wore", "has been wearing"),
        ),
        (
            [
                "My aunt ________ as a nurse in Nanjing for eight years before moving to Suzhou, so she knows the hospital well.",
            ],
            ("works", "worked", "has worked", "has been working"),
        ),
        (
            [
                "Mr Green usually takes Line 2 to work, but he ________ a shared bike these days while the station is closed.",
            ],
            ("rides", "rode", "has ridden", "is riding"),
        ),
        (
            [
                "— Did you see who knocked over the flowerpot in the corridor just now?",
                "— Sorry, no. I ________ my shoelaces with my back to the door.",
            ],
            ("tie", "was tying", "have tied", "had tied"),
        ),
        (
            [
                "Our volunteer club ________ nearly two hundred storybooks this month, and we plan to post them all on Friday.",
            ],
            ("collects", "collected", "has collected", "has been collecting"),
        ),
        (
            [
                "Take a warm coat with you—the weather report says the temperature ________ sharply after sunset.",
            ],
            ("will drop", "drops", "dropped", "has dropped"),
        ),
        (
            [
                "Although this pocket dictionary looks worn out, I ________ it since I started learning English in Grade Three.",
            ],
            ("keep", "have had", "had", "have been having"),
        ),
        (
            [
                "Daily life in our town has become much more convenient since the new bridge ________ two years ago.",
            ],
            ("opens", "has opened", "had opened", "opened"),
        ),
    ]

    for idx, (stem_lines, options) in enumerate(mcqs, start=1):
        add_mcq_item(doc, idx, stem_lines, options)

    # Section 2: Verb Form Fill-in (10 sentences, 16 blanks, 16 pts)
    add_section_title(
        doc,
        "二、用括号内所给动词的适当形式填空（每空 1 分，共 16 分）",
        space_before=4.0,
    )

    fill_items = [
        "My father seldom ____________ (watch) TV on weekdays, but tonight he ________________ (watch) the live match.",
        "While the students ________________ (discuss) their group poster, the bell suddenly ____________ (ring).",
        "Ben felt a little nervous at the boarding gate because he ____________________ (never\u00a0/\u00a0travel) by plane before.",
        "Look at the deep puddles on the playground! How long ____________________ (it\u00a0/\u00a0rain) like this?",
        "Unless the traffic ____________ (improve) soon, I'm afraid we ________________ (miss) the opening speech.",
        "Clara ____________________ (try) to phone the clinic all morning, but nobody ________________ (answer) yet.",
        "By the time I ____________ (find) my seat in the hall, the first song ____________________ (already\u00a0/\u00a0end).",
        "According to the winter timetable, the first ferry ____________ (leave) the pier at 6:30 every morning.",
        "I ____________ (visit) Xi'an once as a child, so this is the second time I ________________ (be) here.",
        "At nine o'clock last night, my parents ____________________ (not\u00a0watch) TV; they were packing their suitcases.",
    ]

    for idx, item in enumerate(fill_items, start=1):
        add_fill_item(doc, idx, item)

    doc.add_page_break()

    # Section 3: Single-Sentence Error Correction (10 items, 10 pts)
    add_section_title(
        doc,
        "三、单句时态改错（每小题 1 分，共 10 分）",
        instruction="下列每句中均有一处动词时态错误，请画出错误部分，并在右侧横线上写出正确的动词形式。",
        space_before=0.0,
    )

    error_items = [
        "This soup is tasting a bit too salty; maybe we should add some hot water.",
        "Our school choir has won first prize in the city competition last November.",
        "If you will leave your wet shoes in the bag overnight, they will smell bad.",
        "When I opened the fridge at midnight, someone already ate the last peach.",
        "My brother repairs his old bike since 2 p.m., and his hands are black with oil.",
        "That wooden violin has been belonging to Mr Lin's family for eighty years.",
        "When the power suddenly went out last night, I typed my essay on the computer.",
        "Mum is still in a meeting, but she promises she calls us back right after six.",
        "Owls are sleeping during the daytime and come out to hunt for food at night.",
        "Jason knows every detail because he has been watching that film three times.",
    ]

    for idx, sentence in enumerate(error_items, start=1):
        add_error_correction_item(doc, idx, sentence)

    # Section 4: Passage Grammar Cloze (14 blanks, 14 pts)
    add_section_title(
        doc,
        "四、语篇语法填空（每空 1 分，共 14 分）",
        instruction="阅读下面短文，结合上下文时间逻辑，用括号内所给动词的正确时态填空。",
        space_before=7.0,
    )

    para1 = [
        ("Two autumns ago, my cousin Leo took me to Red Bay Wetland Park for the first time. Before Leo invited me that morning, I ", False),
        ("1.", True),
        ("\u00a0____________________\u00a0(never\u00a0/\u00a0notice)", False),
        (" how many wild birds lived near our city. When we arrived shortly after sunrise, a thin mist ", False),
        ("2.", True),
        ("\u00a0________________\u00a0(rise)", False),
        (" from the water, and dozens of white birds were standing quietly among the reeds. While Leo ", False),
        ("3.", True),
        ("\u00a0________________\u00a0(adjust)", False),
        (" his camera, a small blue bird suddenly ", False),
        ("4.", True),
        ("\u00a0____________\u00a0(fly)", False),
        (" down to the stream and caught a silver fish. Leo told me that local volunteers ", False),
        ("5.", True),
        ("\u00a0________________\u00a0(plant)", False),
        (" thousands of trees there years before our visit. That single morning completely ", False),
        ("6.", True),
        ("\u00a0____________\u00a0(change)", False),
        (" my weekend habits.", False),
    ]

    para2 = [
        ("Today, Leo and I belong to a student nature club. Every second Saturday, Leo ", False),
        ("7.", True),
        ("\u00a0____________\u00a0(lead)", False),
        (" a small team of students to the wetland to record bird numbers. Over the past two years, we ", False),
        ("8.", True),
        ("\u00a0________________\u00a0(fill)", False),
        (" three thick notebooks with sketches and field notes. Since four o'clock this afternoon, I ", False),
        ("9.", True),
        ("\u00a0____________________\u00a0(sort)", False),
        (" our latest bird photos on my laptop, and I am only halfway through the folder. Right now, my desk is covered with printouts because my sister and I ", False),
        ("10.", True),
        ("\u00a0________________\u00a0(design)", False),
        (" a birdwatching map for our school open day. My sister usually ", False),
        ("11.", True),
        ("\u00a0____________\u00a0(prefer)", False),
        (" indoor hobbies, so everyone is surprised by her enthusiasm.", False),
    ]

    para3 = [
        ("So far, more than forty classmates ", False),
        ("12.", True),
        ("\u00a0________________\u00a0(sign)", False),
        (" up for next weekend's guided walk. The weather forecast looks promising, and I am sure our visitors ", False),
        ("13.", True),
        ("\u00a0________________\u00a0(enjoy)", False),
        (" the peaceful morning views. As soon as the open day ", False),
        ("14.", True),
        ("\u00a0____________\u00a0(end)", False),
        (" next Sunday, we will upload our digital bird guide to the school website for everyone to read.", False),
    ]

    for para in (para1, para2, para3):
        add_cloze_paragraph(doc, para)

    doc.core_properties.title = "英语动词时态综合练习（A2–B1）"
    doc.core_properties.subject = "A2-B1 Mixed English Tenses Homework"
    doc.core_properties.author = "EnglishLearner"
    doc.save(WORKSHEET_PATH)


def add_key_block(doc: Document, heading: str, lines: list[str]) -> None:
    add_section_title(doc, heading, space_before=3.5)
    for line in lines:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.4)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1.2)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        p.paragraph_format.line_spacing = 1.13
        r = p.add_run(line)
        apply_run_fonts(r, 9.2)


def build_answer_key() -> None:
    doc = Document()
    configure_page(doc, show_total_pages=True)

    add_exam_header(
        doc,
        title="英语动词时态综合练习 · 参考答案与解析",
        subtitle="适用难度：CEFR A2–B1　　满分：50 分",
        student_fields=False,
    )

    add_key_block(
        doc,
        "一、单项选择（每小题 1 分，共 10 分）",
        [
            "1–5：C  D  A  B  D　　　　6–10：B  C  A  B  D",
            "1. C（have been chopping）：眼眶发红为当前可见结果，结合 for the last ten minutes，用现在完成进行时。",
            "2. D（had taken）：到达营地（reached/found）为过去时间点，“占好位置”发生在此之前，用过去完成时。",
            "3. A（wears）：until 引导时间状语从句，主句为将来时（won't begin），从句用一般现在时表将来。",
            "4. B（worked）：before moving to Suzhou 表明在南京工作的八年已结束，不延续到现在，用一般过去时。",
            "5. D（is riding）：usually takes 表常态习惯，these days 表近期暂时性的变化，用现在进行时。",
            "6. B（was tying）：花盆被撞倒的那一刻“正在系鞋带”，用过去进行时作背景动作。",
            "7. C（has collected）：this month 尚未结束，且有具体完成数量 nearly two hundred，用现在完成时而非完成进行时。",
            "8. A（will drop）：天气预报对今晚日落后的预测，用一般将来时。",
            "9. B（have had）：since 引导过去起点需用完成时；have 表“拥有”为静态动词，不可用进行时态。",
            "10. D（opened）：since 从句中表示过去起点的动作（two years ago）用一般过去时，主句用现在完成时。",
        ],
    )

    add_key_block(
        doc,
        "二、用括号内所给动词的适当形式填空（每空 1 分，共 16 分）",
        [
            "1. watches；is watching（平日习惯用一般现在时 vs 今晚正在进行用现在进行时）",
            "2. were discussing；rang（过去持续背景动作用过去进行时 vs 突然打断用一般过去时）",
            "3. had never travelled / had never traveled（登机前“从未坐过飞机”，过去的过去用过去完成时）",
            "4. has it been raining（由操场积水推断雨从过去一直下到现在，How long 疑问句用现在完成进行时）",
            "5. improves；will miss（unless 条件从句用一般现在时表将来，主句用一般将来时）",
            "6. has been trying；has answered（一上午反复尝试用现在完成进行时；yet 搭配 nobody 用现在完成时单数）",
            "7. found；had already ended（by the time + 一般过去时，主句“已结束”用过去完成时）",
            "8. leaves（渡轮、航班等固定时刻表用一般现在时）",
            "9. visited；have been（童年去过一次用一般过去时；This is the second time + 现在完成时）",
            "10. were not watching / weren't watching（昨晚九点正在进行的否定状态，与后句 were packing 呼应）",
        ],
    )

    add_key_block(
        doc,
        "三、单句时态改错（每小题 1 分，共 10 分）",
        [
            "1. is tasting → tastes（taste 表“尝起来”系状态系动词，不用进行时）",
            "2. has won → won（句末有具体过去时间 last November，用一般过去时）",
            "3. will leave → leave（if 条件状语从句中用一般现在时表将来）",
            "4. already ate → had already eaten / ate → had eaten（半夜打开冰箱前已被吃掉，用过去完成时）",
            "5. repairs → has been repairing（since 2 p.m. 且现在仍满手油污未修完，用现在完成进行时）",
            "6. has been belonging → has belonged（belong 为静态动词，不可用于完成进行时）",
            "7. typed → was typing（停电发生时“正在台式电脑上打字”，断电后无法继续打字，需用过去进行时）",
            "8. calls → will call（promises 后的宾语从句表将来的承诺，用一般将来时）",
            "9. are sleeping → sleep（客观习性且与后半句 come out 并列，用一般现在时）",
            "10. has been watching → has watched（与具体次数 three times 连用表完成结果，用现在完成时）",
        ],
    )

    add_key_block(
        doc,
        "四、语篇语法填空（每空 1 分，共 14 分）",
        [
            "1. had never noticed（过去完成时）　　2. was rising（过去进行时）　　3. was adjusting（过去进行时）",
            "4. flew（一般过去时）　　5. had planted（过去完成时）　　6. changed（一般过去时）",
            "7. leads（一般现在时）　　8. have filled（现在完成时）　　9. have been sorting（现在完成进行时）",
            "10. are designing（现在进行时）　　11. prefers（一般现在时）　　12. have signed（现在完成时）",
            "13. will enjoy（一般将来时）　　14. ends（一般现在时，as soon as 时间状语从句表将来）",
        ],
    )

    doc.core_properties.title = "英语动词时态综合练习 · 参考答案与解析"
    doc.core_properties.subject = "Teacher Answer Key"
    doc.core_properties.author = "EnglishLearner"
    doc.save(KEY_PATH)


if __name__ == "__main__":
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    build_student_worksheet()
    build_answer_key()
    print(WORKSHEET_PATH)
    print(KEY_PATH)
