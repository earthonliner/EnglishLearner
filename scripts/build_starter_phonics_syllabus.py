#!/usr/bin/env python3
"""Build the revised 24-lesson Our World Starter syllabus."""

from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "lesson-plans" / "our-world-starter-phonics-24-lessons.docx"

TEAL = "16A6A1"
NAVY = "214761"
CORAL = "F26B5E"
YELLOW = "FFD166"
MINT = "EAF8F5"
SKY = "EAF3FA"
CREAM = "FFF9EC"
LIGHT = "F5F7F9"
WHITE = "FFFFFF"
GRAY = "536471"


def set_cell_fill(cell, color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), color)


def set_cell_margin(cell, top=90, start=90, bottom=90, end=90):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for tag, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{tag}"))
        if node is None:
            node = OxmlElement(f"w:{tag}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def prevent_row_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    cant_split = OxmlElement("w:cantSplit")
    tr_pr.append(cant_split)


def set_cell_text(cell, text, *, bold=False, color=GRAY, size=8.5, align=None):
    cell.text = ""
    parts = str(text).split("\n")
    for index, part in enumerate(parts):
        p = cell.paragraphs[0] if index == 0 else cell.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.05
        if align is not None:
            p.alignment = align
        run = p.add_run(part)
        run.bold = bold
        run.font.name = "Aptos"
        run._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
        run.font.size = Pt(size)
        run.font.color.rgb = RGBColor.from_string(color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    set_cell_margin(cell)


def set_repeat_header(row):
    set_repeat_table_header(row)


def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run("课程大纲  ·  ")
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor.from_string(GRAY)
    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = "PAGE"
    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")
    run._r.append(fld_char1)
    run._r.append(instr_text)
    run._r.append(fld_char2)


def add_title(doc, text, subtitle=None):
    p = doc.add_paragraph()
    p.style = "Title"
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    run.font.color.rgb = RGBColor.from_string(NAVY)
    if subtitle:
        p2 = doc.add_paragraph()
        p2.style = "Subtitle"
        r2 = p2.add_run(subtitle)
        r2.font.color.rgb = RGBColor.from_string(TEAL)


def add_section_heading(doc, number, title, note=None):
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    table.columns[0].width = Cm(1.25)
    table.columns[1].width = Cm(15.8)
    left, right = table.rows[0].cells
    set_cell_fill(left, TEAL)
    set_cell_fill(right, MINT)
    set_cell_text(left, number, bold=True, color=WHITE, size=13, align=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_text(right, title, bold=True, color=NAVY, size=13)
    if note:
        p = right.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(note)
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor.from_string(GRAY)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)


def add_callout(doc, label, text, color=CREAM):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = table.cell(0, 0)
    set_cell_fill(cell, color)
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(label + "  ")
    r.bold = True
    r.font.color.rgb = RGBColor.from_string(CORAL)
    r.font.size = Pt(10)
    r2 = p.add_run(text)
    r2.font.color.rgb = RGBColor.from_string(NAVY)
    r2.font.size = Pt(10)
    set_cell_margin(cell, 140, 170, 140, 170)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)


def add_bullets(doc, items, level=0, size=10):
    for item in items:
        p = doc.add_paragraph(style="List Bullet" if level == 0 else "List Bullet 2")
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.08
        r = p.add_run(item)
        r.font.size = Pt(size)
        r.font.color.rgb = RGBColor.from_string(GRAY)


def add_grid_table(doc, headers, rows, widths, font_size=8.5):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.style = "Table Grid"
    for i, width in enumerate(widths):
        table.columns[i].width = width
    hdr = table.rows[0]
    set_repeat_header(hdr)
    for i, header in enumerate(headers):
        set_cell_fill(hdr.cells[i], NAVY)
        set_cell_text(
            hdr.cells[i], header, bold=True, color=WHITE, size=8.5,
            align=WD_ALIGN_PARAGRAPH.CENTER
        )
    for r_index, row_data in enumerate(rows):
        row = table.add_row()
        prevent_row_split(row)
        for i, value in enumerate(row_data):
            set_cell_fill(row.cells[i], WHITE if r_index % 2 == 0 else LIGHT)
            set_cell_text(row.cells[i], value, size=font_size)
    return table


def configure_document(doc):
    section = doc.sections[0]
    section.top_margin = Cm(1.65)
    section.bottom_margin = Cm(1.55)
    section.left_margin = Cm(1.65)
    section.right_margin = Cm(1.65)
    section.header_distance = Cm(0.7)
    section.footer_distance = Cm(0.65)

    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Aptos"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    normal.font.size = Pt(10)
    normal.font.color.rgb = RGBColor.from_string(GRAY)
    normal.paragraph_format.space_after = Pt(5)
    normal.paragraph_format.line_spacing = 1.12

    for style_name, size, color in (
        ("Title", 30, NAVY),
        ("Subtitle", 13, TEAL),
        ("Heading 1", 19, NAVY),
        ("Heading 2", 13, TEAL),
        ("Heading 3", 11, CORAL),
    ):
        style = styles[style_name]
        style.font.name = "Aptos Display"
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
        style.font.size = Pt(size)
        style.font.color.rgb = RGBColor.from_string(color)
        style.font.bold = True
        style.paragraph_format.keep_with_next = True
        style.paragraph_format.space_before = Pt(8)
        style.paragraph_format.space_after = Pt(5)

    if "Small Note" not in styles:
        style = styles.add_style("Small Note", WD_STYLE_TYPE.PARAGRAPH)
        style.font.name = "Aptos"
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
        style.font.size = Pt(8.5)
        style.font.color.rgb = RGBColor.from_string(GRAY)

    header = section.header.paragraphs[0]
    header.text = "OUR WORLD STARTER  ·  PHONICS  ·  STORY-BASED READING"
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for run in header.runs:
        run.font.size = Pt(7.5)
        run.font.bold = True
        run.font.color.rgb = RGBColor.from_string(TEAL)
    add_page_number(section.footer.paragraphs[0])


def add_cover(doc):
    doc.add_paragraph().paragraph_format.space_after = Pt(40)
    tag = doc.add_paragraph()
    tag.paragraph_format.space_after = Pt(8)
    r = tag.add_run("REVISED EDITION  ·  24 LESSONS")
    r.bold = True
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor.from_string(CORAL)
    add_title(
        doc,
        "Our World Starter × Phonics 1\n故事化启蒙课程大纲",
        "从 Unit 1 开始｜短元音前置｜第 2 课进入 CVC｜绘本改为真正的故事",
    )
    doc.add_paragraph().paragraph_format.space_after = Pt(12)
    add_callout(
        doc,
        "课程承诺",
        "第一节课就能说出校园物品、使用完整句型，并听辨第一个辅音与短元音；"
        "不再用两节课重复 ABC。",
        MINT,
    )
    meta = [
        ("适用对象", "深圳宝安区幼儿园大班 / 6 岁；大多数孩子认识字母名，字母音与拼读基础不一"),
        ("课时结构", "24 课 × 90 分钟；Our World Starter 45′ + Phonics 25′ + Story Lab 20′"),
        ("主线教材", "Our World Starter + Phonics 1；Unit 0 不再独立授课"),
        ("阅读定位", "课堂共读负责兴趣与理解；可拼读小书负责“我真的会读”的成就感"),
        ("版本说明", "依据试课反馈重构：保留八个主题单元，重排语音与阅读，不增加机械抄写负担"),
    ]
    add_grid_table(doc, ["项目", "说明"], meta, [Cm(3.2), Cm(13.5)], 9.5)
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Teacher-ready syllabus  ·  2026")
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor.from_string(TEAL)
    doc.add_page_break()


def add_revision_section(doc):
    add_section_heading(doc, "01", "为什么这样改", "回应家长期待，也回应试课中孩子的真实起点")
    doc.add_heading("关键判断", level=2)
    add_bullets(doc, [
        "孩子“认识字母名”不等于“会字母音”，但也不需要从 A–Z 重新教一遍。字母薄弱点改为课堂中的即时诊断与随用随补。",
        "第一节课必须有可复述的学习成果：词汇、句型和声音意识同时出现，让孩子回家能具体地说“今天我会了什么”。",
        "短元音若集中在中段，会让前期只有孤立辅音，CVC 拼读启动过晚；因此五个短元音在 L1–L9 逐步出现。",
        "有情节的绘本可以比孩子当前解码水平更丰富，但不能把“老师可读”误当成“孩子必须独立读”。两条阅读线需要分开。",
    ])
    doc.add_heading("五项结构性调整", level=2)
    rows = [
        ("1", "Unit 0", "取消两课字母复习", "问候、姓名和课堂指令嵌入每节课前 5 分钟"),
        ("2", "Unit 1", "L1 直接开始 My School", "真实文具 + 完整句型，首课就有内容获得感"),
        ("3", "Phonics", "辅音与短元音螺旋推进", "L2 拼读 pat / bat / tap；L9 前见齐五个短元音"),
        ("4", "Reading", "名词清单式读物退出主线", "7 本有角色、问题或变化的故事绘本；另配可拼读小书"),
        ("5", "Homework", "取消统一整本朗读", "20–30 秒最低任务 + 自选挑战任务，不以页数制造压力"),
    ]
    add_grid_table(doc, ["#", "板块", "原安排", "新安排"], rows, [Cm(1), Cm(2.1), Cm(5.2), Cm(8.4)], 8.8)
    doc.add_heading("学期结束时希望孩子做到", level=2)
    add_bullets(doc, [
        "听懂并使用 8 个主题单元的核心词汇与 8 组高频句型。",
        "说出已学字母音，能把 2–3 个音合成为简单 CVC 词，也能把 CVC 词拆成音。",
        "理解 7 个短故事的主要情节，借助图片复述 beginning → problem/change → ending。",
        "形成正确阅读习惯：从左到右指读、遇到可拼读词先看音、不靠死背整页。",
    ])
    add_callout(
        doc,
        "边界",
        "本学期不追求学完 Phonics 1 全册。软 c/g、y、th、j、wh、mb 留给下学期，"
        "优先保证短元音与 CVC 真正会用。",
        SKY,
    )


def add_framework_section(doc):
    doc.add_page_break()
    add_section_heading(doc, "02", "每节课的固定节奏", "90 分钟内每 8–15 分钟切换活动形式")
    rows = [
        ("0–5′", "Hello Routine", "问候、姓名 / 情绪、1 条课堂指令；顺带复现易忘字母，不单开 ABC 课"),
        ("5–20′", "Explore", "实物、图片、动作或故事场景导入 6–9 个主题词"),
        ("20–35′", "Speak", "用 1 个核心句型完成信息差、猜测、问答或角色扮演"),
        ("35–45′", "Use It", "小游戏 / 教材故事 / 小作品；完成 Today I can… 口头检查"),
        ("45–55′", "Hear It", "先听音、做动作和口型，不急着看字母"),
        ("55–65′", "Map It", "声音 ↔ 字母；单音辨别、首尾音分类、少量描写"),
        ("65–70′", "Blend It", "从已学音中拼 CVC；基础档跟拼，挑战档独立拼 / 换首音"),
        ("70–82′", "Story Lab", "预测—老师共读—找线索；关注角色与变化，不逐句翻译"),
        ("82–87′", "Respond", "表演、排序、画一页、用重复句回应或 20 秒图片复述"),
        ("87–90′", "Exit Ticket", "每人完成一个词 / 一句 / 一次拼读，明确回家最低任务"),
    ]
    add_grid_table(doc, ["时间", "环节", "做什么"], rows, [Cm(2), Cm(3), Cm(11.7)], 9)
    doc.add_heading("分层不分组贴标签", level=2)
    rows = [
        ("基础任务 MUST DO", "指图说词；跟读完整句；在老师拉音帮助下拼出 CVC；复述 1 个画面"),
        ("挑战任务 CAN DO", "替换句中 1 个词；独立拼读并造句；补充 because；按三幅图复述故事"),
        ("相同反馈", "完成各自目标都获得同样的贴纸 / 印章；不公开比较录音长度或识字量"),
    ]
    add_grid_table(doc, ["层级", "表现"], rows, [Cm(4), Cm(12.7)], 9)

    doc.add_heading("L1：让孩子第一天就“有东西可带回家”", level=2)
    rows = [
        ("Starter", "实物 school bag 寻宝：book, pencil, pen, crayon, paper, eraser, desk, chair"),
        ("Sentence", "What is it? — It’s a ___.  挑战：It’s a red / blue ___."),
        ("Phonics", "先听 /p/，再认识短 /a/；做口型、听辨和 p–a sound slide，不要求背字母表"),
        ("Story", "Spot Goes to School：预测 Spot 带了什么、比较 Spot 和自己的上学准备"),
        ("Exit", "每人抽 1 件文具，说完整句；指出 p 或 a 的声音卡"),
        ("带回家", "“Today I can name school things and say It’s a… / 我会说文具，还会用一个完整句。”"),
    ]
    add_grid_table(doc, ["板块", "可执行内容"], rows, [Cm(3.2), Cm(13.5)], 9)


def add_phonics_section(doc):
    doc.add_page_break()
    add_section_heading(doc, "03", "Phonics 螺旋路线", "元音提前、CVC 提前；教材页码用于取材，不等于按页顺序硬推")
    rows = [
        ("L1", "/p/ + short /a/", "p.12 + p.32", "p–a sound slide；听辨 p / a", "pen, paper"),
        ("L2", "/b/, /t/", "p.13–14", "pat, bat, tap（第一次 CVC）", "book, table"),
        ("L3", "/d/ + a-family", "p.15", "dad, bad, pad；a 词族小游戏", "desk"),
        ("L4", "short /i/ + /f/", "p.34 + p.24", "fit, fib；a / i 最小对比", "five, favorite"),
        ("L5", "/v/, /g/", "p.25–26", "dig, pig, big, gap, bag", "van, big"),
        ("L6", "short /e/ + /l/", "p.33 + p.27", "leg, let, beg, peg", "leg, balloon"),
        ("L7", "/s/, /z/", "p.18–19", "sat, sit, zip；清浊触喉", "sister, zero"),
        ("L8", "short /o/ + /m/", "p.35 + p.20", "mom, mop, log, dog", "mom, bedroom"),
        ("L9", "/n/ + short /u/", "p.21 + p.38", "sun, nut, bun；五元音到齐", "nose, hands"),
        ("L10", "hard c / k", "p.44, 46", "cat, cap, kid, kit；同音不同字母", "coat / classroom"),
        ("L11", "/x/ + CVC review", "p.47 + Review", "box, fox, six；五元音分类", "six, box"),
        ("L12", "Midterm check", "—", "听音选图、字母音、CVC 合音 / 分音", "不考背 A–Z"),
        ("L13", "/h/", "p.53", "hat, hot, hit；换元音", "happy"),
        ("L14", "/r/", "p.52", "rat, red, run；三元音对比", "red, room"),
        ("L15", "ch / sh", "p.60–61", "chat, shop, fish；二合字母作一个音", "chicken, fish"),
        ("L16", "digraph blending", "p.60–61", "chip, chop, shut；句中找词", "lunch, milkshake"),
        ("L17", "U5–6 review", "Review 4–6", "元音中音听辨 + 拼读接力", "food / story words"),
        ("L18", "/w/", "p.64", "win, wet, wag", "wear, white"),
        ("L19", "/ng/", "p.66", "sing, wing, ring；关注词尾", "wearing"),
        ("L20", "CVC fluency", "已学页复现", "每分钟不计数；按准确性分层", "animal words"),
        ("L21", "sentence decoding", "教师配卡", "The cat can run. / The dog is big.", "动作句"),
        ("L22", "story sound hunt", "全册复现", "在熟悉绘本中找已学音", "自主选择"),
        ("L23", "final review", "Review 1–6", "听—拼—读—用 四站轮换", "全部复现"),
        ("L24", "Final check", "—", "音形对应、CVC、短句与故事回应", "成长展示"),
    ]
    add_grid_table(
        doc,
        ["课次", "目标音", "教材取材", "拼读应用", "主题呼应"],
        rows,
        [Cm(1.2), Cm(3.3), Cm(2.7), Cm(5.4), Cm(4.1)],
        7.9,
    )
    add_callout(
        doc,
        "教学口令",
        "Say it slowly → Touch each sound → Push the sounds together → Read it fast. "
        "拉音时用 /m/ /s/ /f/ 等连续音，/b/ /t/ /d/ 不额外加“呃”。",
        CREAM,
    )


def add_reading_section(doc):
    doc.add_page_break()
    add_section_heading(doc, "04", "Story Path：从“词汇清单”换成“真正的故事”")
    add_bullets(doc, [
        "主线绘本允许含少量未学词，由老师借助图片、动作与重复句支撑；目标是预测、理解、参与和复述。",
        "可拼读读物必须与已学音匹配。只有当孩子能解码约 80% 的正文时，才适合作为独立阅读任务。",
        "同一本书通常用 2–3 次：第一次看懂故事，第二次参与重复句，第三次表演 / 重组 / 迁移；不为凑课次机械重复。",
    ])
    rows = [
        ("L1–3 · U1", "Spot Goes to School", "上学、教室、担心与适应", "预测 Spot 会带什么；比较自己的 school bag；表演第一天上学"),
        ("L4–6 · U2", "Kipper’s Toybox", "玩具箱破洞，玩具不断少掉", "数玩具、找问题、猜是谁拿走玩具；用 Is it a…?"),
        ("L7–8 · U3", "Where’s Spot?", "妈妈在家中寻找 Spot", "房间 + 方位；每次翻页前用 Where’s Spot? 预测"),
        ("L9–11 · U4", "Go Away, Big Green Monster!", "怪物逐步出现又被赶走", "颜色 + 五官；按顺序重建怪物；用祈使句回应"),
        ("L13–17 · U5–6", "The Very Hungry Caterpillar", "进食、肚子痛、蜕变", "故事顺序、食物与数量；判断 healthy / unhealthy；三图复述"),
        ("L18–19 · U7", "Pete the Cat: I Love My White Shoes", "鞋子不断变色，Pete 仍继续走", "衣服 + 颜色 + 情绪韧性；重复唱句可做最低输出"),
        ("L20–23 · U8", "Dear Zoo", "动物园寄来不合适的宠物", "动物 + 形容词；猜下一只动物；用 too big / too tall 说明理由"),
    ]
    add_grid_table(
        doc,
        ["课次 / 主题", "主线绘本", "故事钩子", "课堂可做"],
        rows,
        [Cm(2.6), Cm(4), Cm(4), Cm(6.1)],
        8.3,
    )
    doc.add_heading("第二条阅读线：真正可拼读", level=2)
    rows = [
        ("L2–5", "教师自制 1–2 句 mini strips", "仅用已学 p b t d f + a/i；如 A bat. / Pat the bat. 教师检查每个字"),
        ("L6–11", "Bob Books Set 1（选配）", "按实际字母覆盖率选册，不按书号赶进度；先课堂共拼，再带回 1–2 页"),
        ("L13–21", "Jolly Phonics Readers Level 1 或 ORT Floppy’s Phonics Stage 1+（选配）", "挑选短元音、ch/sh 与已学高频词覆盖度高的故事"),
        ("全学期", "班级 Sound Book", "每学一个音，每个孩子贡献 1 张图 / 1 个词，期末合成全班可读小书"),
    ]
    add_grid_table(doc, ["阶段", "读物", "使用原则"], rows, [Cm(2.5), Cm(5), Cm(9.2)], 8.7)
    doc.add_heading("绘本替换备选", level=2)
    p = doc.add_paragraph(style="Small Note")
    p.add_run(
        "衣服主题可用 Froggy Gets Dressed 替换 Pete the Cat；身体主题可增加 From Head to Toe "
        "作为动作共读；若班级语言基础明显较弱，Dear Zoo 每次只读 3–4 个动物，仍保留完整情节。"
    )


LESSONS = [
    ("L1\nU1A", "My School\nbook, pencil, pen, crayon, paper, eraser, desk, chair\nWhat is it? — It’s a ___.",
     "/p/ + short /a/\n听辨与 p–a sound slide", "Spot Goes to School ①\n封面预测；找 school things",
     "20 秒：展示 3 件文具，说 It’s a…"),
    ("L2\nU1B", "green / yellow；3 / 4\nTime for School\nIt’s a green ___.",
     "/b/ /t/\npat, bat, tap（首批 CVC）", "Spot ②\n图片排序 + 重复句回应",
     "拼 2 个 CVC + 说 1 个完整句"),
    ("L3\nWorkshop", "Unit 1 任务：Pack my school bag\n问答、分类、听指令；不新增大组词汇",
     "/d/ + a-family\ndad, bad, pad", "Spot ③\n表演第一天上学；三图复述",
     "Today I can… 学习卡；自选读 1 页"),
    ("L4\nU2A", "My Toys\nballoon, car, doll, teddy bear, kite, truck, robot, ball\nIs it a ___?",
     "short /i/ + /f/\nfit, fib；a / i 对比", "Kipper’s Toybox ①\n数玩具，发现“少了什么”",
     "拍 1 个玩具：Is it a…? Yes, it is."),
    ("L5\nU2B", "orange / brown；5 / 6\nBirthday Boy\nThis is my favorite toy.",
     "/v/ /g/\ndig, pig, big, gap, bag", "Kipper ②\n预测洞从哪里来；角色声音",
     "读 1 条 mini strip；介绍 favorite toy"),
    ("L6\nWorkshop", "Unit 2 玩具商店：问、答、数、分类\n复用 Is it…? / It’s…",
     "short /e/ + /l/\nleg, let, beg, peg", "Kipper ③\nproblem–solution 排序与表演",
     "最低 20 秒；挑战：Bob Books 匹配页"),
    ("L7\nU3A", "My Family\nfamily members + rooms\nWho’s this? / Where’s ___?",
     "/s/ /z/\nsat, sit, zip；清浊对比", "Where’s Spot? ①\n逐页预测可能藏在哪里",
     "用家庭照说 This is my ___."),
    ("L8\nU3B", "black / white；7 / 8\nEight Is Great\nIn the kitchen / bedroom.",
     "short /o/ + /m/\nmom, mop, log, dog", "Where’s Spot? ②\n房间地图 + 翻翻卡复述",
     "Where’s…? In the… 两轮亲子问答"),
    ("L9\nU4A", "My Body\nnose, feet, legs, arms, hands, mouth, eyes, ears, hair\nI have two ___.",
     "/n/ + short /u/\nsun, nut, bun；五元音到齐", "Big Green Monster ①\n怪物逐步出现；顺序卡",
     "画怪物并说 I have…"),
    ("L10\nU4B", "pink / purple；9 / 10\nIt’s Me!\nTouch / Show me…",
     "hard c / k\ncat, cap, kid, kit", "Monster ②\nGo away! 参与读；反向拆除",
     "拼 3 个词；指图说 2 个 body words"),
    ("L11\nReview", "U1–4 stations\n文具 / 玩具 / 家庭 / 身体；四组句型轮换",
     "/x/ + five-vowel review\nbox, fox, six", "Monster ③\n自制新怪物；完整表演",
     "期中口试模拟；不布置整本录音"),
    ("L12\nMidterm", "听力 + 口试 + 成长反馈\n重点：听懂、说整句、主题应用",
     "音形对应 + CVC 合音 / 分音\n不考背字母表", "轻松共读：From Head to Toe\n动作回应",
     "与家长共看成长卡，选 1 项下阶段目标"),
    ("L13\nU5A", "Stories\nking, queen, prince, princess, crown, bird, happy, sad\nI want a ___.",
     "/h/\nhat, hot, hit", "Hungry Caterpillar ①\n封面预测；故事开端",
     "用 I want… 说 1 个愿望"),
    ("L14\nU5B", "gold / silver\nThe Crown\n故事角色与情绪",
     "/r/\nrat, red, run", "Caterpillar ②\n食物顺序 + 数量",
     "排序 3 张图，口述 first / then"),
    ("L15\nU6A", "I Like Food\ncookies, rice, noodles, chicken, milk, juice, bread, bananas, water\nI like / don’t like…",
     "ch / sh\nchat, shop, fish", "Caterpillar ③\n食物分类；找重复结构",
     "家庭投票：I like… / I don’t like…"),
    ("L16\nU6B", "circle / square\nA Picnic\nDo you like ___?",
     "digraph blending\nchip, chop, shut", "Caterpillar ④\n肚子痛原因；healthy choice",
     "读 1–2 页或 1 句；挑战三图复述"),
    ("L17\nWorkshop", "U5–6 故事野餐任务\n选角色、列食物、用两组句型完成表演",
     "U5–6 review\n元音中音听辨 + 拼读接力", "Caterpillar ⑤\nbeginning–middle–ending 完整复述",
     "20–30 秒故事复述，不要求背原文"),
    ("L18\nU7A", "Clothes\nshirt, shorts, socks, pants, shoes, skirt, dress, hat, coat\nI’m wearing ___.",
     "/w/\nwin, wet, wag", "Pete the Cat ①\n白鞋变色；预测下一种颜色",
     "穿搭播报：I’m wearing…"),
    ("L19\nU7B", "triangle / rectangle\nWash Day\nIt’s hot / cold.",
     "/ng/\nsing, wing, ring", "Pete ②\n重复唱句；变化—反应排序",
     "最低：重复句 1 次；挑战：复述 3 次变化"),
    ("L20\nU8A", "Animals\nhippo, crocodile, lion, giraffe, zebra, monkey + actions\nIs it ___ing?",
     "CVC fluency\n按准确性分层，不做速度排名", "Dear Zoo ①\n猜箱中动物；big / tall / fierce",
     "选 1 只动物：It’s too ___."),
    ("L21\nU8B", "star / heart\nAt the Animal Park\nIt’s running / eating / drinking.",
     "sentence decoding\nThe cat can run. / The dog is big.", "Dear Zoo ②\n动物—理由配对；参与重复句",
     "读 1 条短句 + 做动作"),
    ("L22\nProject", "My Little Story\n每人做 1 页：角色 + 地点 + 一个变化；小组排成故事",
     "story sound hunt\n在熟悉文本中找已学音", "Dear Zoo ③\n替换动物，创作班级新版本",
     "练习自己的 1 页，不练全班整本"),
    ("L23\nFinal Review", "U1–8 四站轮换\n词汇、句型、听指令、故事排序；口试同伴模拟",
     "听—拼—读—用 四站\n集中补个人薄弱音", "7 本书自由选择\n最喜欢的角色 / 情节投票",
     "准备展示：1 句 + 1 个拼读词"),
    ("L24\nFinal", "听力 + 个别口试 + 课程展示\n成长评价替代公开排名",
     "音形、CVC、短句综合检查\n记录“独立 / 提示后 / 需帮助”", "Story Showcase\n小组表演 + 班级 Sound Book",
     "发送成长报告与下阶段阅读建议"),
]


def add_lesson_tables(doc):
    phases = [
        ("05", "Phase 1 · Start Strong", "L1–6｜直接进入 U1；L2 开始 CVC；两本有情节的校园 / 玩具故事", 0, 6),
        ("06", "Phase 2 · Five Vowels in Place", "L7–12｜在期中前见齐五个短元音，并完成 U1–4 综合应用", 6, 12),
        ("07", "Phase 3 · Read for Meaning", "L13–17｜故事与食物主题；从 digraph 进入短句拼读", 12, 17),
        ("08", "Phase 4 · Use & Showcase", "L18–24｜衣服、动物、班级创作与期末成长展示", 17, 24),
    ]
    for number, title, note, start, end in phases:
        doc.add_page_break()
        add_section_heading(doc, number, title, note)
        add_grid_table(
            doc,
            ["课次", "Starter 45′", "Phonics 25′", "Story Lab 20′", "轻量作业"],
            LESSONS[start:end],
            [Cm(1.4), Cm(4.6), Cm(3.2), Cm(4.2), Cm(3.3)],
            7.75,
        )
        if start == 0:
            add_callout(
                doc,
                "家长可见的 L1 成果",
                "会说 8 个校园词；会问答 What is it? — It’s a…；能听出 /p/ 与短 /a/；"
                "能说出 Spot 第一天上学发生了什么。",
                MINT,
            )
        elif start == 6:
            add_callout(
                doc,
                "期中不是字母考试",
                "只检查已学音形对应、CVC 合音 / 分音、主题词句与故事理解；不再要求从 A 背到 Z。",
                SKY,
            )
        elif start == 12:
            add_callout(
                doc,
                "共读 ≠ 背书",
                "The Very Hungry Caterpillar 由老师承担复杂文本，孩子只需参与重复结构、排序并用已学句型复述。",
                CREAM,
            )
        else:
            add_callout(
                doc,
                "结课展示",
                "每个孩子只准备“1 个完整句 + 1 个可拼读词 + 1 次故事回应”，保证真实、短小、可成功。",
                MINT,
            )


def add_homework_assessment(doc):
    doc.add_page_break()
    add_section_heading(doc, "09", "作业与评价：减少压力，但保留可见成长")
    doc.add_heading("每课只发一个最低任务", level=2)
    rows = [
        ("A · Speak", "看 1 张图 / 1 件实物，说 1 个完整句", "15–20 秒", "所有孩子"),
        ("B · Decode", "点音拼 2–3 个已学 CVC，或读 1 条 mini strip", "20–30 秒", "L2 起轮换"),
        ("C · Story", "跟读 1 个重复句、读 1 页，或看图复述 1 个事件", "20–30 秒", "主线绘本第 2 次共读后"),
        ("Challenge", "多读 1–2 页 / 用 because / 三图复述", "自选", "不与贴纸或排名绑定"),
    ]
    add_grid_table(doc, ["类型", "最低任务", "建议时长", "何时使用"], rows, [Cm(2.4), Cm(7.3), Cm(2.4), Cm(4.6)], 8.8)
    add_bullets(doc, [
        "不再统一要求把老师共读的整本经典绘本录下来；未教过的文字不作为孩子独立朗读表现。",
        "同一条录音可同时完成口语或拼读目标，不固定“音频 + 描红 + 绘本”三件套。",
        "书写每课最多 1 小块：目标字母 2–4 个 + 自己的名字；重点是方向、位置和握笔，不以行数为目标。",
        "若连续未交，课堂内给孩子 2 分钟与老师完成最低任务，不把家长配合度等同于孩子能力。",
    ])
    doc.add_heading("期中与期末评价蓝图", level=2)
    rows = [
        ("Listening · 30", "听词指图 / 听句选图 / 听课堂指令", "主题理解，不放孤立冷词"),
        ("Speaking · 30", "看图说词、完整句问答、Today I can…", "独立完成 / 提示后完成 / 需帮助"),
        ("Phonics · 25", "听音选字母、CVC 合音、CVC 分音、读 1 条短句", "期中覆盖 L1–11；期末累计复现"),
        ("Story response · 15", "排序 3 幅图、选角色、回答 What happened?", "可看图，不要求背原文"),
    ]
    add_grid_table(doc, ["板块", "任务", "评价原则"], rows, [Cm(3.5), Cm(7), Cm(6.2)], 8.8)
    doc.add_heading("阶段范围", level=2)
    rows = [
        ("L12 期中", "U1–4；/p b t d f v g l s z m n/；short a/i/e/o/u；hard c/k；x；基础 CVC"),
        ("L24 期末", "U1–8；期中内容累计复现 + /h r ch sh w ng/；短句拼读；故事结构回应"),
        ("留到下学期", "soft c/g、y、th、j、wh、mb；更复杂辅音连缀；不为“学完一本书”压缩练习时间"),
    ]
    add_grid_table(doc, ["节点", "范围"], rows, [Cm(3.4), Cm(13.3)], 8.8)


def add_parent_section(doc):
    doc.add_page_break()
    add_section_heading(doc, "10", "教师备课与家长沟通")
    doc.add_heading("每节课发一张 3 行成长卡", level=2)
    add_callout(
        doc,
        "TODAY I CAN…",
        "1) I can say: book, pencil, pen…  2) I can ask: What is it?  "
        "3) I can hear: /p/ and short /a/.  家庭任务：任选 3 件文具完成 20 秒问答。",
        MINT,
    )
    add_bullets(doc, [
        "先告诉家长“本课可观察的能力”，再给词表；避免家长只按单词数量判断课程含金量。",
        "明确区分 letter name 与 letter sound：认识 A–Z 是起点，不是本课程要重复两节的终点。",
        "每四课反馈一次：词句表达、语音意识、拼读、故事理解、课堂习惯五维各 1 句话，不只发分数。",
        "绘本版权内容只用于购买后的课堂共读；家庭阅读优先使用家庭已购书或合法数字资源。",
    ])
    doc.add_heading("备课检查清单", level=2)
    rows = [
        ("□", "本课是否只有 1 个核心句型，并有真实交流目的？"),
        ("□", "目标音是否先听后看？CVC 是否只使用已教音？"),
        ("□", "绘本任务是否关注角色、问题与变化，而不是逐页翻译？"),
        ("□", "基础档孩子是否能在 30 秒内成功完成 exit ticket？"),
        ("□", "家庭最低任务是否在 20–30 秒内可完成？"),
        ("□", "是否写好了面向家长的 Today I can…？"),
    ]
    add_grid_table(doc, ["", "检查问题"], rows, [Cm(1.2), Cm(15.5)], 9)
    doc.add_heading("绘本采购 / 借阅优先级", level=2)
    rows = [
        ("首批", "Spot Goes to School; Kipper’s Toybox; Where’s Spot?; Go Away, Big Green Monster!"),
        ("中后段", "The Very Hungry Caterpillar; Pete the Cat: I Love My White Shoes; Dear Zoo"),
        ("可拼读", "Bob Books Set 1，或 Jolly Phonics Readers Level 1 / ORT Floppy’s Phonics Stage 1+；先核对音素覆盖"),
    ]
    add_grid_table(doc, ["用途", "建议"], rows, [Cm(3), Cm(13.7)], 8.7)
    p = doc.add_paragraph(style="Small Note")
    p.paragraph_format.space_before = Pt(10)
    p.add_run(
        "说明：主线书目依据“有明确情节、重复语言、主题贴合、可互动”筛选。"
        "具体版本文字量可能不同，教师备课时应以手中版本为准，不把陌生词量直接转化为孩子的朗读作业。"
    )


def build():
    doc = Document()
    configure_document(doc)
    add_cover(doc)
    add_revision_section(doc)
    add_framework_section(doc)
    add_phonics_section(doc)
    add_reading_section(doc)
    add_lesson_tables(doc)
    add_homework_assessment(doc)
    add_parent_section(doc)

    for section in doc.sections:
        section.top_margin = Cm(1.65)
        section.bottom_margin = Cm(1.55)
        section.left_margin = Cm(1.65)
        section.right_margin = Cm(1.65)
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    doc.save(OUTPUT)
    print(f"Wrote {OUTPUT}")


if __name__ == "__main__":
    build()
