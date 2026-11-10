#!/usr/bin/env python3
"""Build Lesson 1: My School! — photo-based, multi-sensory, retrieval-rich deck.

Design principles (grounded in YL-EFL research):
- real-style photos instead of vector art (AI-generated originals, see credits)
- every word enters through a DIFFERENT sense hook + its own action chant
- a retrieval game after every two new words (spacing within the lesson)
- every word gets 8-10 varied encounters; meaning + sound only, no spelling tests
- a Day1/Day3/Day7 home review plan against the forgetting curve
"""

import io
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "lesson-plans" / "our-world-starter-lesson1-my-school.pptx"
ASSETS = ROOT / "scripts" / "assets" / "lesson1"
MUSIC_DIR = ROOT / "lesson-plans" / "demo-music"
POSTER = ROOT / "scripts" / "assets" / "audio-poster.png"

W, H = 13.333, 7.5

NAVY = "1F3C88"
BLUE = "2D9CDB"
GREEN = "2BB673"
CORAL = "FF6B6B"
YELLOW = "FFC93C"
ORANGE = "F79646"
PURPLE = "8064A2"
INK = "333333"
MUTED = "777777"
WHITE = "FFFFFF"
P_BLUE = "E3F2FB"
P_GREEN = "E4F7EC"
P_CORAL = "FFE9E9"
P_YELLOW = "FFF4CC"
P_PURPLE = "EFEAF6"

EN = "Comic Sans MS"
CN = "微软雅黑"

TRACKS = {
    "enter": ("Fluffing_a_Duck.mp3", "进场音乐"),
    "mystery": ("Sneaky_Snitch.mp3", "声音袋音乐"),
    "pass": ("Monkeys_Spinning_Monkeys.mp3", "音乐停 → 回答"),
    "show": ("Carefree.mp3", "游戏背景音乐"),
}

IMGS = {
    "book": "lesson1-obj-book.jpg",
    "pencil": "lesson1-obj-pencil.jpg",
    "pen": "lesson1-obj-pen.jpg",
    "crayon": "lesson1-obj-crayon.jpg",
    "paper": "lesson1-obj-paper.jpg",
    "eraser": "lesson1-obj-eraser.jpg",
    "desk": "lesson1-obj-desk.jpg",
    "chair": "lesson1-obj-chair.jpg",
    "pip": "lesson1-char-pip.jpg",
    "pip_sad": "lesson1-char-pip-sad.jpg",
    "mia": "lesson1-char-mia.jpg",
    "monster": "lesson1-char-monster.jpg",
}

WORDS = ["book", "pencil", "pen", "crayon", "paper", "eraser", "desk", "chair"]


def img_path(key):
    return ASSETS / IMGS[key]


def rgb(color):
    return RGBColor.from_string(color)


def _east_asia(run, cn=CN):
    rpr = run._r.get_or_add_rPr()
    ea = rpr.find(qn("a:ea"))
    if ea is None:
        ea = rpr.makeelement(qn("a:ea"), {})
        latin = rpr.find(qn("a:latin"))
        if latin is not None:
            latin.addnext(ea)
        else:
            rpr.append(ea)
    ea.set("typeface", cn)


def text(slide, value, x, y, w, h, size=24, color=INK, bold=False,
         align=PP_ALIGN.CENTER, font=EN, valign=MSO_ANCHOR.MIDDLE):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.margin_left = Inches(0.05)
    frame.margin_right = Inches(0.05)
    frame.margin_top = Inches(0.02)
    frame.margin_bottom = Inches(0.02)
    frame.vertical_anchor = valign
    for i, part in enumerate(str(value).split("\n")):
        p = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        p.alignment = align
        p.space_after = Pt(0)
        run = p.add_run()
        run.text = part
        run.font.name = font
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = rgb(color)
        _east_asia(run)
    return box


def rect(slide, x, y, w, h, fill, line_color=None, radius=True):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    shape.line.color.rgb = rgb(line_color or fill)
    if radius:
        try:
            shape.adjustments[0] = 0.12
        except (IndexError, ValueError):
            pass
    return shape


def circle(slide, x, y, d, fill, line_color=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    shape.line.color.rgb = rgb(line_color or fill)
    return shape


def bar(slide, x1, y1, x2, y2, color=NAVY):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(x1), Inches(y1),
        Inches(max(0.02, x2 - x1)), Inches(max(0.02, y2 - y1)),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(color)
    shape.line.color.rgb = rgb(color)
    return shape


def photo(slide, key, x, y, w):
    return slide.shapes.add_picture(str(img_path(key)), Inches(x), Inches(y),
                                    width=Inches(w), height=Inches(w))


def photo_card(slide, key, x, y, w, label=None, covered=False):
    """Square photo in a white rounded card. `covered` hides it behind a ? card
    but keeps the same card height as labelled neighbours."""
    extra = 0.5 if (label or covered) else 0
    rect(slide, x - 0.07, y - 0.07, w + 0.14, w + 0.14 + extra, WHITE, "C9D6E4", True)
    photo(slide, key, x, y, w)
    if covered:
        rect(slide, x - 0.02, y - 0.02, w + 0.04, w + 0.04, WHITE, CORAL, True)
        text(slide, "?", x, y + w / 2 - 0.85, w, 1.7, 100, CORAL, True)
    elif label:
        text(slide, label, x, y + w + 0.02, w, 0.42, 22, NAVY, True)


def pen_zoom(slide, x, y, w):
    """Macro crop of the pen photo — the Zoom-In sense entry."""
    im = Image.open(img_path("pen"))
    c = im.width // 2
    box = (int(c - 0.24 * im.width), int(c - 0.16 * im.width),
           int(c + 0.24 * im.width), int(c + 0.28 * im.width))
    crop = im.crop(box).resize((900, int(900 * (box[3] - box[1]) / (box[2] - box[0]))), Image.LANCZOS)
    buf = io.BytesIO()
    crop.save(buf, format="JPEG", quality=90)
    buf.seek(0)
    height = w * (box[3] - box[1]) / (box[2] - box[0])
    return slide.shapes.add_picture(buf, Inches(x), Inches(y), width=Inches(w), height=Inches(height))


def white_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(WHITE)
    return slide


def page_num(slide, n):
    text(slide, str(n), 12.6, 7.05, 0.6, 0.35, 12, MUTED, align=PP_ALIGN.RIGHT)


def header(slide, title, fill):
    rect(slide, 0.5, 0.25, 12.3, 1.0, fill)
    text(slide, title, 0.7, 0.32, 11.9, 0.86, 40, WHITE, True)


def cue(slide, note, x=1.2, y=7.02):
    w = max(1.5, 12.83 - x)
    text(slide, note, x, y, w, 0.38, 14, MUTED, align=PP_ALIGN.LEFT)


def add_music(slide, key):
    filename, label = TRACKS[key]
    path = MUSIC_DIR / filename
    slide.shapes.add_movie(
        str(path), Inches(0.3), Inches(6.55), Inches(0.7), Inches(0.7),
        poster_frame_image=str(POSTER), mime_type="audio/mpeg",
    )
    text(slide, label, 1.02, 6.68, 3.6, 0.45, 14, MUTED, align=PP_ALIGN.LEFT)


def chip(slide, value, x, y, w, h, fill, color=WHITE, size=16):
    rect(slide, x, y, w, h, fill, fill, True)
    text(slide, value, x + 0.08, y, w - 0.16, h, size, color, True)


def photo_row(slide, keys, y, size=1.5, labels=False):
    total = len(keys) * size + (len(keys) - 1) * 0.35
    x = (W - total) / 2
    for k in keys:
        if labels:
            photo_card(slide, k, x, y, size, label=k)
        else:
            photo(slide, k, x, y, size)
        x += size + 0.35


def word_slide(prs, n, idx, kind, sense_label, chant, cue_note, accent):
    s = white_slide(prs)
    chip(s, sense_label, 0.45, 0.4, 3.1, 0.62, accent, WHITE, 15)
    for i in range(8):
        cx = 9.85 + i * 0.36
        circle(s, cx, 0.58, 0.24, accent if i < idx else "E4E9EF",
               accent if i < idx else "E4E9EF")
    text(s, f"{idx}/8", 12.78, 0.52, 0.55, 0.35, 14, MUTED, True)
    photo(s, kind, 4.42, 1.05, 4.5)
    text(s, kind, 0.5, 5.55, 12.3, 1.1, 72, NAVY, True)
    text(s, chant, 1.0, 6.68, 11.3, 0.42, 24, accent, True)
    cue(s, cue_note, x=0.5, y=7.1)
    page_num(s, n)
    return s


def build():
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    prs.core_properties.title = "Lesson 1: My School!"
    prs.core_properties.subject = "Our World Starter U1 × Phonics 1 (L1) · multi-sensory edition"
    prs.core_properties.author = "EnglishLearner"
    prs.core_properties.keywords = "My School, phonics, /p/, /b/, retrieval games, Pip"

    # 1 cover --------------------------------------------------------------
    s = white_slide(prs)
    text(s, "My School!", 0.5, 0.4, 12.3, 1.5, 80, NAVY, True, PP_ALIGN.LEFT)
    text(s, "第一课 ｜ Our World Starter · Unit 1 ｜ 90 分钟 ｜ 今天捉住 8 个单词", 0.55, 1.9, 12.0, 0.5, 20, MUTED, align=PP_ALIGN.LEFT)
    photo(s, "book", 0.7, 2.55, 3.7)
    photo(s, "pencil", 4.9, 2.55, 2.3)
    photo(s, "pen", 7.5, 2.55, 2.3)
    photo(s, "crayon", 10.1, 2.55, 2.3)
    photo(s, "paper", 4.9, 5.0, 2.3)
    photo(s, "eraser", 7.5, 5.0, 2.3)
    photo(s, "pip", 10.1, 4.9, 2.4)
    add_music(s, "enter")
    page_num(s, 1)

    # 2 hello --------------------------------------------------------------
    s = white_slide(prs)
    photo(s, "pip", 0.7, 1.4, 4.9)
    rect(s, 6.0, 1.5, 6.8, 2.5, BLUE, BLUE)
    text(s, "Hello! I'm Pip.\nWhat's your name?", 6.2, 1.68, 6.4, 2.2, 38, WHITE, True)
    rect(s, 6.9, 4.55, 5.0, 1.4, YELLOW, YELLOW)
    text(s, "I'm ____ .", 7.1, 4.7, 4.6, 1.1, 44, INK, True)
    cue(s, "把 Pip 的照片当布偶：先自我介绍，再传球逐个问答，要求说整句", x=0.5)
    page_num(s, 2)

    # 3 pip says -----------------------------------------------------------
    s = white_slide(prs)
    header(s, "🙋  Pip says…", GREEN)
    commands = [("🧍", "Stand up"), ("🧎", "Sit down"), ("🙌", "Hands up"), ("👀", "Eyes on me")]
    for i, (emoji, cmd) in enumerate(commands):
        x = 0.7 + (i % 2) * 6.2
        y = 1.6 + (i // 2) * 2.6
        text(s, emoji, x, y, 5.8, 2.2, 120)
        rect(s, x, y + 2.35, 5.8, 1.1, BLUE if i % 2 == 0 else CORAL)
        text(s, cmd, x, y + 2.5, 5.8, 0.8, 40, WHITE, True)
    text(s, "Hands up = I want to talk 🗣️   ｜   两三轮，越喊越快", 1.0, 6.5, 11.3, 0.5, 24, MUTED)
    page_num(s, 3)

    # 4 quest map ------------------------------------------------------------
    s = white_slide(prs)
    header(s, "🐛  Today's Quest: Catch 8 Words!", ORANGE)
    photo_row(s, WORDS[:4], 1.45, 2.0, labels=True)
    photo_row(s, WORDS[4:], 4.15, 2.0, labels=True)
    rect(s, 2.0, 6.9, 9.3, 0.48, P_YELLOW, ORANGE, True)
    text(s, "每捉住一个单词，就在探险卡上贴一个贴纸！捉满 8 个 = 单词小猎人 🏅", 2.2, 6.96, 8.9, 0.34, 15, INK, True)
    page_num(s, 4)

    # 5 book ----------------------------------------------------------------
    word_slide(prs, 5, 1, "book", "👂 Sound 入口", "book, book, open the book!",
               "翻书声入口：手臂当封面张开合上 flap flap ×2 → 揭晓照片", BLUE)

    # 6 pencil ----------------------------------------------------------------
    word_slide(prs, 6, 2, "pencil", "👀 Slow Reveal", "pencil, pencil, write, write, write!",
               "慢速揭示：把打印词卡从大书后面一点点抽出，全班喊 STOP 才全露", ORANGE)

    # 7 echo mountain ---------------------------------------------------------
    s = white_slide(prs)
    header(s, "⛰️  Echo Mountain 回声山", BLUE)
    for cx in (1.2, 7.9):
        tri = s.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(cx), Inches(2.0), Inches(4.2), Inches(3.6))
        tri.fill.solid(); tri.fill.fore_color.rgb = rgb(GREEN); tri.line.color.rgb = rgb(NAVY)
    text(s, "🏔️ 老师（山顶）", 1.4, 1.55, 4.0, 0.5, 22, NAVY, True)
    text(s, "🧒 孩子们（山谷回声）", 8.1, 1.55, 4.0, 0.5, 22, NAVY, True)
    chip(s, "book!", 2.2, 4.35, 2.2, 0.8, YELLOW, INK, 30)
    chip(s, "pencil!", 8.9, 4.35, 2.6, 0.8, YELLOW, INK, 30)
    text(s, "老师喊一声 → 全班回声两遍：大声一遍、悄悄一遍", 2.2, 6.15, 8.9, 0.5, 22, MUTED)
    cue(s, "复习刚学的 2 个词：每词 3 轮（大→小→快）。这就是『提取练习』", x=0.5)
    page_num(s, 7)

    # 8 pen zoom entry -----------------------------------------------------------
    s = white_slide(prs)
    chip(s, "🔍  ZOOM IN!", 0.45, 0.4, 2.6, 0.62, GREEN, WHITE, 15)
    pen_zoom(s, 4.27, 1.35, 4.8)
    rect(s, 2.4, 5.35, 8.5, 0.85, P_GREEN, GREEN, True)
    text(s, "这是哪个小东西的特写？猜一猜！", 2.6, 5.47, 8.1, 0.6, 26, NAVY, True)
    cue(s, "先让孩子随便猜（中英文都行），再翻到下一页揭晓照片 + chant", x=0.5, y=7.05)
    page_num(s, 8)

    # 9 pen --------------------------------------------------------------------
    word_slide(prs, 9, 3, "pen", "🔍 Zoom In", "pen, pen, click it — write it!",
               "揭晓特写答案；按笔发出 click click，再带 chant 两遍", GREEN)

    # 10 crayon ----------------------------------------------------------------
    word_slide(prs, 10, 4, "crayon", "🌈 Air Draw", "crayon, crayon, draw a rainbow!",
               "空中入口：跟老师在空中画彩虹，边画边说 draw draw draw", CORAL)

    # 11 feed the monster ---------------------------------------------------------
    s = white_slide(prs)
    header(s, "👹  Feed the Monster! 喂怪兽", GREEN)
    photo(s, "monster", 0.8, 1.6, 4.6)
    text(s, "GRRR… I'm hungry! 😋", 1.0, 6.4, 4.2, 0.5, 24, GREEN, True)
    for i, w in enumerate(["book", "pencil", "pen", "crayon"]):
        x = 6.0 + (i % 2) * 3.6
        y = 1.75 + (i // 2) * 1.5
        chip(s, w, x, y, 3.2, 1.0, [BLUE, ORANGE, GREEN, CORAL][i], WHITE, 34)
    rect(s, 6.0, 5.05, 6.6, 1.3, P_YELLOW, ORANGE, True)
    text(s, "抽一张词卡 → 大声说出单词 → 喂给怪兽\nMunch munch! 说错怪兽会吐出来哦 😝", 6.2, 5.18, 6.2, 1.0, 20, INK, True)
    cue(s, "怪兽图投屏或打印；词卡用 Slide 47-48 打印页，说对的卡才被『吃掉』", x=6.0, y=6.5)
    page_num(s, 11)

    # 12 paper -----------------------------------------------------------------
    word_slide(prs, 12, 5, "paper", "✈️ Paper Plane", "paper, paper, whoosh — a plane!",
               "纸飞机入口：老师折的纸飞机飞过教室，孩子追着说 whoosh!", PURPLE)

    # 13 eraser ----------------------------------------------------------------
    word_slide(prs, 13, 6, "eraser", "🪄 Magic Fix", "eraser, eraser, rub it out!",
               "魔术入口：白板故意写错一笔，请一位孩子上来擦掉 → rub it out!", BLUE)

    # 14 telepathy ---------------------------------------------------------------
    s = white_slide(prs)
    header(s, "🧠  Telepathy 心灵感应", PURPLE)
    text(s, "🧠", 5.9, 1.45, 1.6, 1.6, 90)
    text(s, "老师心里正想着一张卡… 你能读出老师的想法吗？", 1.5, 3.05, 10.3, 0.55, 24, NAVY, True)
    photo_row(s, ["book", "pencil", "pen", "crayon", "paper", "eraser"], 3.8, 1.45)
    for i in range(6):
        x = (W - (6 * 1.45 + 5 * 0.35)) / 2 + i * 1.8 + 0.45
        chip(s, str(i + 1), x, 5.55, 0.55, 0.5, PURPLE, WHITE, 18)
    cue(s, "报数字猜词：猜中 → 夸张地『你真会读心！』；猜不中 → 给线索（它的开头是 /p/）", x=0.5, y=6.4)
    page_num(s, 14)

    # 15 desk -------------------------------------------------------------------
    word_slide(prs, 15, 7, "desk", "🥁 Knock Knock", "desk, desk, knock, knock, knock!",
               "节奏入口：knock knock knock 敲三下桌，分组接龙问答", GREEN)

    # 16 chair -------------------------------------------------------------------
    word_slide(prs, 16, 8, "chair", "🎵 Musical Sit", "chair, chair, sit down — plonk!",
               "音乐椅子入口：音乐一停，坐下并喊 chair！每人一次（音乐可用任意儿歌）", CORAL)

    # 17 lip reading ---------------------------------------------------------------
    s = white_slide(prs)
    header(s, "👄  Lip Reading 唇语挑战", CORAL)
    text(s, "👄", 5.9, 1.5, 1.6, 1.6, 90)
    text(s, "老师不动声音，只做口型 —— 你能猜出是哪个词吗？", 1.5, 3.1, 10.3, 0.55, 24, NAVY, True)
    photo_row(s, WORDS, 3.85, 1.3)
    cue(s, "8 个词全复习！先易后难；猜对的孩子当小老师，换他做口型", x=0.5, y=5.8)
    page_num(s, 17)

    # 18 sound bag --------------------------------------------------------------
    s = white_slide(prs)
    header(s, "👂  Sound Bag 声音袋", BLUE)
    sounds = [
        ("flap flap", "翻书声", BLUE), ("scratch scratch", "写字声", ORANGE),
        ("click click", "按笔声", GREEN), ("scribble scribble", "涂色声", CORAL),
        ("whoosh!", "飞过去", PURPLE), ("rub rub", "擦擦擦", BLUE),
        ("knock knock", "敲敲敲", GREEN), ("plonk!", "坐下来", CORAL),
    ]
    for i, (snd, meaning, fill) in enumerate(sounds):
        x = 0.75 + (i % 2) * 6.2
        y = 1.7 + (i // 2) * 1.22
        chip(s, f"🔊 {snd}   {meaning}", x, y, 5.9, 0.92, fill, WHITE, 24)
    rect(s, 2.6, 6.55, 8.1, 0.5, P_YELLOW, ORANGE, True)
    text(s, "老师抽一张『声音卡』演声音 → 你指照片墙 + 喊出那个英文词！", 2.8, 6.61, 7.7, 0.36, 16, INK, True)
    add_music(s, "mystery")
    cue(s, "听觉通道第二次提取：8 个声音 ↔ 8 个词一一对应", x=4.85, y=7.12)
    page_num(s, 18)

    # 19 flash ------------------------------------------------------------------
    s = white_slide(prs)
    header(s, "⚡  FLASH! 闪卡挑战", ORANGE)
    text(s, "⚡", 0.7, 1.65, 2.6, 2.6, 110)
    rect(s, 3.6, 1.7, 9.1, 1.15, P_YELLOW, ORANGE, True)
    text(s, "1. 闭眼睛   2. 睁眼看半秒   3. 大声喊出来！", 3.8, 1.85, 8.7, 0.8, 26, INK, True)
    photo_row(s, WORDS[:4], 3.35, 1.55, labels=True)
    cue(s, "照片卡在胸前快速翻面半秒；先练前 4 个词，熟了换后 4 个；谁先喊对谁当『翻卡员』", x=0.5, y=6.65)
    page_num(s, 19)

    # 20-22 what's missing x3 -----------------------------------------------------
    grid6 = ["book", "pencil", "pen", "paper", "desk", "chair"]
    positions = [(1.35, 1.45), (5.55, 1.45), (9.75, 1.45), (1.35, 4.3), (5.55, 4.3), (9.75, 4.3)]
    s = white_slide(prs)
    header(s, "👀  What's missing?", CORAL)
    for k, (x, y) in zip(grid6, positions):
        photo_card(s, k, x, y, 2.1, label=k)
    text(s, "10… 9… 8…  仔细看！", 3.6, 7.0, 6.1, 0.45, 26, CORAL, True)
    page_num(s, 20)
    for r, missing in ((1, {"pencil"}), (2, {"pencil", "chair"})):
        s = white_slide(prs)
        header(s, "👀  What's missing?", CORAL)
        for k, (x, y) in zip(grid6, positions):
            photo_card(s, k, x, y, 2.1, label=k, covered=k in missing)
        text(s, "什么不见了？大声喊出来！" if r == 1 else "两个都不见了！You have 10 seconds!",
             3.3, 7.0, 6.7, 0.45, 26, CORAL, True)
        page_num(s, 20 + r)

    # 23 sentence model -------------------------------------------------------------
    s = white_slide(prs)
    header(s, "❓  What is it?", BLUE)
    photo(s, "pen", 0.8, 1.6, 4.4)
    rect(s, 5.5, 2.2, 7.2, 1.8, GREEN, GREEN)
    text(s, "It's a pen.", 5.7, 2.4, 6.8, 1.4, 54, WHITE, True)
    text(s, "☝️    ✌️    🤟", 5.5, 4.3, 7.2, 1.0, 54)
    text(s, "It's  –  a  –  pen  （手指打节拍，一词一拍）", 5.5, 5.5, 7.2, 0.6, 26, MUTED)
    cue(s, "Pip 问，老师答；再互换。要求整句，不要只说单词", x=0.5)
    page_num(s, 23)

    # 24 roll and say --------------------------------------------------------------
    s = white_slide(prs)
    header(s, "🎲  Roll & Say 情绪骰子", PURPLE)
    faces = [("1️⃣", "happy 😊"), ("2️⃣", "sad 😢"), ("3️⃣", "angry 😠"),
             ("4️⃣", "sleepy 😴"), ("5️⃣", "tiny 🤏"), ("6️⃣", "GIANT 🦣")]
    for i, (n, mood) in enumerate(faces):
        x = 0.7 + (i % 3) * 4.1
        y = 1.65 + (i // 3) * 1.6
        rect(s, x, y, 3.7, 1.3, P_PURPLE, PURPLE, True)
        text(s, f"{n} {mood}", x + 0.15, y + 0.15, 3.4, 1.0, 28, INK, True)
    rect(s, 2.6, 5.15, 8.1, 1.15, NAVY, NAVY)
    text(s, "掷骰子 → 用掷中的情绪说出老师指的单词！", 2.8, 5.3, 7.7, 0.8, 28, WHITE, True)
    cue(s, "骰子用打印正方体或口头报数代替；一词掷一次，全班一起变声", x=0.5, y=6.55)
    page_num(s, 24)

    # 25 pass the bag -------------------------------------------------------------
    s = white_slide(prs)
    header(s, "🎒  Pass the Bag!", CORAL)
    text(s, "🎒", 0.9, 2.0, 4.5, 4.0, 170)
    rect(s, 6.5, 2.2, 6.3, 2.6, BLUE, BLUE)
    text(s, "What is it?\nIt's a ____ .", 6.7, 2.4, 5.9, 2.2, 40, WHITE, True)
    text(s, "拿到书包的孩子摸出 1 件实物 → 问答 → 传给下一个", 6.6, 5.1, 6.1, 0.55, 20, MUTED)
    add_music(s, "pass")
    page_num(s, 25)

    # 26 peek-a-boo ---------------------------------------------------------------
    s = white_slide(prs)
    header(s, "👀  Peek-a-boo!  What is it?", GREEN)
    for i, k in enumerate(("crayon", "eraser")):
        x = 1.1 + i * 6.0
        photo(s, k, x, 1.8, 3.4)
        rect(s, x + 0.15, 1.95, 3.4, 3.1, P_YELLOW if i == 0 else P_GREEN, ORANGE if i == 0 else GREEN, True)
        text(s, "?", x + 0.15, 2.6, 3.4, 2.0, 110, ORANGE if i == 0 else GREEN, True)
    rect(s, 2.9, 5.9, 7.5, 0.85, NAVY, NAVY)
    text(s, "只露一个角，也要说整句：It's a ___ !", 3.1, 6.02, 7.1, 0.6, 26, WHITE, True)
    cue(s, "用打印照片词卡藏在卡纸后只露一角；猜对全班鼓掌 3 下", x=0.5, y=7.05)
    page_num(s, 26)

    # 27 guess my card --------------------------------------------------------------
    s = white_slide(prs)
    header(s, "🕵️  Guess My Card 猜猜我的卡", PURPLE)
    text(s, "🕵️", 5.7, 1.5, 1.9, 1.9, 100)
    photo_row(s, ["book", "pencil", "crayon", "paper"], 3.6, 1.55)
    rect(s, 1.4, 5.7, 10.5, 1.1, P_PURPLE, PURPLE, True)
    text(s, "老师心里想一张卡。你可以问：Is it red?  Can we write with it?\n老师只答 Yes / No —— 猜中的人下次抽卡！", 1.6, 5.83, 10.1, 0.85, 22, INK, True)
    cue(s, "老师示范 2 个问句后，孩子只要猜词即可；这是『带线索的提取』", x=0.5, y=6.95)
    page_num(s, 27)

    # 28 color challenge ------------------------------------------------------------
    s = white_slide(prs)
    header(s, "⭐  Add one color!", YELLOW)
    photo(s, "pen", 1.2, 1.8, 3.5)
    rect(s, 1.2, 5.5, 3.9, 0.75, P_GREEN, GREEN, True)
    text(s, "It's a blue pen.", 1.3, 5.6, 3.7, 0.5, 26, INK, True)
    photo(s, "book", 7.7, 1.8, 3.5)
    rect(s, 7.7, 5.5, 3.9, 0.75, P_BLUE, BLUE, True)
    text(s, "It's a blue book.", 7.8, 5.6, 3.7, 0.5, 26, INK, True)
    text(s, "CORE ✓  It's a pen.        CHALLENGE ★  It's a blue pen.", 2.0, 6.55, 9.3, 0.5, 22, MUTED)
    cue(s, "会说的孩子加颜色词；说原句同样得贴纸", x=0.5, y=7.1)
    page_num(s, 28)

    # 29 freeze dance ------------------------------------------------------------------
    s = white_slide(prs)
    header(s, "🎶  Freeze Dance!", PURPLE)
    text(s, "MOVE 💃", 0.9, 1.9, 3.6, 1.2, 54, CORAL, True)
    text(s, "→", 4.6, 1.95, 1.0, 1.0, 44, MUTED, True)
    text(s, "FREEZE 🧊", 5.7, 1.9, 3.9, 1.2, 54, BLUE, True)
    text(s, "→", 9.7, 1.95, 1.0, 1.0, 44, MUTED, True)
    text(s, "MIME 🎭", 10.8, 1.9, 2.5, 1.2, 54, GREEN, True)
    photo_row(s, WORDS[:4], 3.4, 1.7, labels=True)
    text(s, "音乐停 → 听到哪个词，就做那个词的动作并定住！", 2.2, 6.3, 8.9, 0.55, 22, MUTED)
    add_music(s, "show")
    page_num(s, 29)

    # 30-31 phonics letters ------------------------------------------------------------
    s = white_slide(prs)
    rect(s, 0.5, 0.6, 4.8, 3.2, YELLOW, YELLOW)
    text(s, "P p", 0.7, 0.85, 4.4, 2.7, 140, NAVY, True)
    text(s, "/p/  /p/  /p/", 0.6, 4.05, 4.6, 1.2, 44, CORAL, True)
    for i, k in enumerate(["pen", "pencil", "paper"]):
        x = 5.6 + i * 2.5
        photo(s, k, x, 1.15, 2.15)
        text(s, k, x, 3.4, 2.15, 0.55, 30, NAVY, True)
    text(s, "🤫 纸巾实验：/p/ 能吹动纸巾，喉咙不振动（字母名 pee ≠ 字母音 /p/）", 0.7, 5.35, 12.0, 0.6, 24, MUTED)
    text(s, "pen · pencil · paper —— 它们开头都住着 /p/", 0.7, 6.05, 12.0, 0.5, 22, CORAL, True)
    cue(s, "注意：paper 的 a 不发短音，今天只听开头音（L2 才学 short /a/，拼 pat · bat · tap）", x=0.7, y=6.7)
    page_num(s, 30)

    s = white_slide(prs)
    rect(s, 0.5, 0.6, 4.8, 3.2, BLUE, BLUE)
    text(s, "B b", 0.7, 0.85, 4.4, 2.7, 140, WHITE, True)
    text(s, "/b/  /b/  /b/", 0.6, 4.05, 4.6, 1.2, 44, GREEN, True)
    photo(s, "book", 5.7, 1.15, 2.15)
    text(s, "book", 5.7, 3.4, 2.15, 0.55, 30, NAVY, True)
    text(s, "🎒", 8.5, 1.15, 2.5, 2.5, 100)
    text(s, "bag（书包）", 8.35, 3.4, 2.8, 0.55, 30, NAVY, True)
    text(s, "✋ 摸喉咙：/b/ 会振动，/p/ 不振动 —— 同样双唇，差别在喉咙", 0.7, 5.35, 12.0, 0.6, 24, MUTED)
    text(s, "book 和 bag 都是 /b/ 开头！", 0.7, 6.05, 12.0, 0.5, 22, GREEN, True)
    cue(s, "一手纸巾一手摸喉咙，/p/ 和 /b/ 各做三遍对比", x=0.7, y=6.7)
    page_num(s, 31)

    # 32 pop or bounce -----------------------------------------------------------------
    s = white_slide(prs)
    header(s, "🍿  Pop for /p/ !  Bounce for /b/ !", CORAL)
    moves = [
        ("🍿", "POP UP!", "/p/  pen · pencil · paper", P_YELLOW, ORANGE),
        ("🫧", "BOUNCE!", "/b/  book · bag", P_BLUE, BLUE),
        ("🧊", "FREEZE!", "desk · chair · crayon · eraser", P_GREEN, GREEN),
    ]
    for i, (emoji, name, words_str, tint, fill) in enumerate(moves):
        x = 0.75 + i * 4.15
        rect(s, x, 1.7, 3.6, 4.2, tint, fill, True)
        text(s, emoji, x, 2.0, 3.6, 1.5, 80)
        text(s, name, x, 3.7, 3.6, 0.6, 30, INK, True)
        text(s, words_str, x + 0.15, 4.45, 3.3, 1.0, 18, MUTED)
    cue(s, "随机喊词，孩子做对应动作；最后一轮用今天全部 8 个词", x=0.5, y=6.2)
    page_num(s, 32)

    # 33 swat it -----------------------------------------------------------------------
    s = white_slide(prs)
    header(s, "🪰  Swat it!", YELLOW)
    rect(s, 1.0, 1.8, 4.8, 3.6, P_YELLOW, ORANGE, True)
    text(s, "P p", 1.2, 2.1, 4.4, 3.0, 150, NAVY, True)
    rect(s, 7.5, 1.8, 4.8, 3.6, P_BLUE, BLUE, True)
    text(s, "B b", 7.7, 2.1, 4.4, 3.0, 150, NAVY, True)
    text(s, "🪰", 5.95, 2.5, 1.6, 2.0, 90)
    text(s, "Listen  →  Turn  →  Swat!  →  Say the sound", 1.0, 5.8, 11.3, 0.7, 28, NAVY, True)
    cue(s, "两只苍蝇拍；听到词跑到对应首音卡片拍下并大声说 /p/ 或 /b/", x=0.5, y=6.6)
    page_num(s, 33)

    # 34 robot talk ---------------------------------------------------------------------
    s = white_slide(prs)
    header(s, "🤖  Robot Talk!", BLUE)
    robots = [
        ("/p/  -  /e/  -  /n/", "pen!", "pen", P_YELLOW),
        ("/b/  -  /u/  -  /k/", "book!", "book", P_BLUE),
        ("/d/  -  /e/  -  /s/  -  /k/", "desk!", "desk", P_GREEN),
    ]
    for i, (seg, whole, k, tint) in enumerate(robots):
        y = 1.65 + i * 1.5
        rect(s, 0.8, y, 11.7, 1.25, tint, WHITE, True)
        text(s, "🤖 ROBOT:   " + seg, 1.1, y + 0.2, 6.2, 0.85, 26, NAVY, True, PP_ALIGN.LEFT)
        text(s, "→", 7.35, y + 0.2, 0.7, 0.85, 30, CORAL, True)
        text(s, whole, 8.1, y + 0.2, 1.8, 0.85, 30, GREEN, True)
        photo(s, k, 11.0, y + 0.12, 1.0)
    text(s, "机器人慢速说 → 你快快说成一个词！（纯听音，不看拼写）", 1.5, 6.25, 10.3, 0.5, 20, MUTED)
    text(s, "下节课 L2 预告：short /a/ 来啦，拼出 pat · bat · tap！", 1.5, 6.8, 10.3, 0.45, 18, CORAL, True)
    page_num(s, 34)

    # 35 board race ----------------------------------------------------------------------
    s = white_slide(prs)
    header(s, "🏃  Board Race 触卡接力!", GREEN)
    photo_row(s, WORDS[:4], 1.65, 2.0, labels=True)
    photo_row(s, WORDS[4:], 4.35, 2.0, labels=True)
    rect(s, 1.7, 7.0, 10.0, 0.42, P_GREEN, GREEN, True)
    text(s, "两队排队 → 老师喊词 → 第一人跑上前拍对应照片卡 → 拍对得分！", 1.85, 7.05, 9.7, 0.3, 15, INK, True)
    page_num(s, 35)

    # 36 story cover -----------------------------------------------------------------------
    s = white_slide(prs)
    header(s, "📖  Story Time:  Pip's First Day", CORAL)
    photo(s, "pip", 0.9, 1.6, 4.6)
    text(s, "Pip 要上学啦！", 0.9, 6.35, 4.6, 0.5, 24, GREEN, True)
    rect(s, 6.1, 1.7, 6.5, 1.0, P_GREEN, GREEN, True)
    text(s, "What will Pip pack?", 6.3, 1.82, 6.1, 0.75, 30, NAVY, True)
    for i, k in enumerate(["book", "pencil", "crayon"]):
        photo(s, k, 6.5 + i * 2.1, 3.2, 1.8)
    text(s, "指一指，猜一猜！", 6.5, 5.35, 5.7, 0.5, 22, CORAL, True)
    cue(s, "指照片让孩子预测并说词（第 1 次故事接触）；猜完进入下一幕", x=0.5, y=6.1)
    page_num(s, 36)

    # 37 story 1 ---------------------------------------------------------------------------
    s = white_slide(prs)
    header(s, "📖  1  Pip packs for school.", GREEN)
    photo(s, "pip", 0.6, 1.7, 3.6)
    photo_row(s, ["book", "pencil", "pen", "paper"], 2.4, 1.55, labels=True)
    rect(s, 1.3, 6.15, 10.7, 0.8, P_YELLOW, ORANGE, True)
    text(s, "A book, a pencil, a pen, and paper.", 1.5, 6.25, 10.3, 0.6, 26, NAVY, True)
    cue(s, "读一遍；孩子只跟读实物词。问：What is in the bag?", x=0.5, y=7.05)
    page_num(s, 37)

    # 38 story 2 ---------------------------------------------------------------------------
    s = white_slide(prs)
    header(s, "📖  2  Oh no! No crayon.", CORAL)
    photo(s, "pip_sad", 0.7, 1.6, 4.6)
    rect(s, 6.0, 1.9, 3.2, 3.2, P_CORAL, CORAL, True)
    photo(s, "crayon", 6.55, 2.35, 2.1)
    bar(s, 6.3, 2.0, 8.9, 5.0, CORAL)
    text(s, "NO crayon! 😢", 6.2, 5.15, 3.6, 0.6, 30, CORAL, True)
    rect(s, 10.0, 1.9, 3.0, 1.0, P_YELLOW, ORANGE, True)
    text(s, "How does Pip feel?", 10.1, 2.05, 2.8, 0.7, 20, NAVY, True)
    text(s, "I need a crayon.", 6.0, 5.95, 3.4, 0.55, 26, NAVY, True)
    cue(s, "停在这里：书包里 4 样都有了，少了什么？让孩子指照片说 crayon", x=0.5, y=6.75)
    page_num(s, 38)

    # 39 story 3 ---------------------------------------------------------------------------
    s = white_slide(prs)
    header(s, "📖  3  Mia shares her crayon.", BLUE)
    photo(s, "mia", 0.9, 1.7, 4.6)
    text(s, "Here you are!  🖍️", 5.9, 2.3, 3.4, 0.7, 30, GREEN, True)
    text(s, "Thank you!  🤗", 5.9, 4.6, 3.4, 0.7, 30, NAVY, True)
    rect(s, 9.6, 2.1, 3.3, 2.9, P_BLUE, BLUE, True)
    text(s, "A crayon + a kind friend\n= a happy first day!\n💛", 9.75, 2.35, 3.0, 2.4, 22, NAVY, True)
    cue(s, "两人一组用一支蜡笔演 Here you are! / Thank you!（第 3 次故事接触）", x=0.5, y=6.7)
    page_num(s, 39)

    # 40 story order ------------------------------------------------------------------------
    s = white_slide(prs)
    header(s, "🃏  First — Then — Finally", PURPLE)
    order_cards = [
        ("FIRST", "Pip packs.", GREEN, "pip"),
        ("THEN", "No crayon!", CORAL, "pip_sad"),
        ("FINALLY", "Mia shares.", BLUE, "mia"),
    ]
    for i, (marker, caption, fill, k) in enumerate(order_cards):
        x = 0.8 + i * 4.15
        rect(s, x, 1.6, 3.6, 4.6, WHITE, fill, True)
        rect(s, x, 1.6, 3.6, 0.62, fill, fill, False)
        text(s, marker, x + 0.15, 1.7, 3.3, 0.4, 16, WHITE, True)
        photo(s, k, x + 0.55, 2.4, 2.5)
        text(s, caption, x + 0.2, 5.2, 3.2, 0.65, 24, NAVY, True)
    text(s, "3 个孩子各拿 1 张卡，全班喊出正确顺序再站好！", 1.5, 6.45, 10.3, 0.5, 22, MUTED)
    cue(s, "排序卡打印在 Slide 50；说错不扣分，重来一次", x=0.5, y=7.0)
    page_num(s, 40)

    # 41 story detectives ---------------------------------------------------------------------
    s = white_slide(prs)
    header(s, "🔍  Story Detectives", GREEN)
    questions = [
        ("POINT", "Where is the crayon?", P_BLUE, BLUE),
        ("SAY", "What is it?", P_YELLOW, ORANGE),
        ("THINK", "Who helps Pip?", P_CORAL, CORAL),
        ("ACT", "Here you are!\nThank you!", P_GREEN, GREEN),
    ]
    for i, (level, q, tint, fill) in enumerate(questions):
        x = 0.7 + (i % 2) * 6.2
        y = 1.65 + (i // 2) * 2.5
        rect(s, x, y, 5.8, 2.15, tint, fill, True)
        rect(s, x, y, 1.5, 2.15, fill, fill, False)
        text(s, level, x + 0.05, y + 0.8, 1.4, 0.5, 15, WHITE, True)
        text(s, q, x + 1.75, y + 0.35, 3.8, 1.45, 26, INK, True)
    cue(s, "每个孩子至少答 1 题：指 → 说词 → 短句 → 表演", x=0.5, y=6.6)
    page_num(s, 41)

    # 42 exit doors --------------------------------------------------------------------------
    s = white_slide(prs)
    header(s, "🚪  Choose one door!", CORAL)
    doors = [
        ("DOOR 1", "Say 3 words", "看照片说 3 个词", P_BLUE, BLUE),
        ("DOOR 2", "Ask + answer", "What is it?\nIt's a…", P_YELLOW, ORANGE),
        ("DOOR 3", "Catch /p/!", "pen · pencil · paper", P_GREEN, GREEN),
    ]
    for i, (door, task, hint, tint, fill) in enumerate(doors):
        x = 0.78 + i * 4.15
        rect(s, x, 1.8, 3.6, 4.2, tint, fill, True)
        circle(s, x + 2.95, 4.35, 0.3, NAVY, NAVY)
        text(s, door, x + 0.2, 2.1, 3.2, 0.5, 16, fill, True)
        text(s, task, x + 0.2, 2.9, 3.2, 0.7, 28, INK, True)
        text(s, hint, x + 0.2, 4.9, 3.2, 0.9, 20, MUTED)
    cue(s, "出门考：每人 30–40 秒；记录 独立 / 提示后 / 需帮助", x=0.5, y=6.35)
    page_num(s, 42)

    # 43 20-sec challenge ----------------------------------------------------------------------
    s = white_slide(prs)
    header(s, "🏠  20-second School Bag Challenge", GREEN)
    photo_row(s, ["book", "pencil", "pen"], 1.9, 2.1, labels=True)
    steps = [
        ("1", "Pick 3 school things."),
        ("2", "Say: It's a ___ , It's a ___ , It's a ___ ."),
        ("3", "Bonus: /p/ /p/ pen!"),
    ]
    for i, (n, phrase) in enumerate(steps):
        y = 5.15 + i * 0.6
        circle(s, 1.6, y, 0.5, [ORANGE, BLUE, CORAL][i], NAVY)
        text(s, n, 1.6, y + 0.07, 0.5, 0.35, 15, WHITE, True)
        text(s, phrase, 2.35, y - 0.02, 10.5, 0.5, 20, INK, True, PP_ALIGN.LEFT)
    cue(s, "录 20-30 秒小视频发群；不要求整本朗读", x=0.5, y=7.05)
    page_num(s, 43)

    # 44 three-day review plan --------------------------------------------------------------------
    s = white_slide(prs)
    header(s, "📅  回家温习计划：3 天捉牢 8 个词", NAVY)
    plan = [
        ("今晚 Day 1", "20 秒 School Bag Challenge", "挑 3 件文具说整句 It's a ___.，录小视频发群", "30 秒", P_BLUE, BLUE),
        ("周三 Day 3", "Teach Your Toys 教玩偶", "把 3 张照片词卡『教』给你的玩偶，玩偶说错了你来纠正", "2 分钟", P_GREEN, GREEN),
        ("周六 Day 7", "Home I-Spy 家里侦察", "在家找到 book / paper / pen，找到就大声说 It's a ___ !", "3 分钟", P_CORAL, CORAL),
    ]
    for i, (day, title_, desc, dur, tint, fill) in enumerate(plan):
        x = 0.6 + i * 4.3
        rect(s, x, 1.7, 3.95, 4.1, tint, fill, True)
        rect(s, x, 1.7, 3.95, 0.7, fill, fill, False)
        text(s, f"{day}  ⏱ {dur}", x + 0.15, 1.8, 3.65, 0.45, 17, WHITE, True)
        text(s, title_, x + 0.2, 2.65, 3.55, 0.6, 22, INK, True)
        text(s, desc, x + 0.25, 3.4, 3.45, 1.5, 16, MUTED)
    rect(s, 1.4, 6.1, 10.5, 0.75, P_YELLOW, ORANGE, True)
    text(s, "遗忘曲线告诉我们：隔天再取一次，记得才牢 —— 三次小复习 > 一次长复习！", 1.6, 6.22, 10.1, 0.5, 18, INK, True)
    cue(s, "本页拍照发家长群；下节课开场会先复习这 8 个词", x=0.5, y=6.98)
    page_num(s, 44)

    # 45 today we learned -----------------------------------------------------------------------
    s = white_slide(prs)
    header(s, "🏠  今天我们学了", NAVY)
    rows = [
        ("🧾 单词", "book, pencil, pen, crayon, paper, eraser, desk, chair", P_BLUE, BLUE),
        ("💬 句子", "What is it? — It's a ___.   挑战：It's a blue pen.", P_YELLOW, ORANGE),
        ("🔤 字母的声音", "P 说 /p/（pen · pencil · paper）  B 说 /b/（book · bag）—— 字母有名字，也有声音", P_GREEN, GREEN),
        ("🎒 回家任务", "Day 1 今晚 20 秒挑战 ｜ Day 3 教玩偶 ｜ Day 7 家里 I-Spy（详见上页）", P_CORAL, CORAL),
    ]
    for i, (label, content, tint, fill) in enumerate(rows):
        y = 1.55 + i * 1.35
        rect(s, 0.6, y, 3.0, 1.1, fill, fill)
        text(s, label, 0.75, y + 0.12, 2.7, 0.85, 26, WHITE, True)
        rect(s, 3.8, y, 9.0, 1.1, tint, fill, True)
        text(s, content, 4.0, y + 0.1, 8.6, 0.9, 21, INK, True, PP_ALIGN.LEFT)
    cue(s, "拍照发家长群，附一句 Today I can… 成长说明", x=0.6, y=7.02)
    page_num(s, 45)

    # 46 goodbye -------------------------------------------------------------------------------
    s = white_slide(prs)
    photo(s, "pip", 0.8, 1.3, 4.9)
    rect(s, 6.2, 2.2, 6.6, 2.8, BLUE, BLUE)
    text(s, "Goodbye!\n👋", 6.4, 2.4, 6.2, 2.4, 50, WHITE, True)
    text(s, "下次课预告：认识 short /a/，拼出第一批 CVC：pat · bat · tap！", 5.9, 5.5, 7.4, 0.55, 22, MUTED)
    cue(s, "Goodbye Song 建议用 Super Simple Songs 流媒体版（见 Credits）", x=0.5, y=6.35)
    page_num(s, 46)

    # 47-48 printable photo cards --------------------------------------------------------------
    for idx, (name, keys) in enumerate((("Set A", WORDS[:4]), ("Set B", WORDS[4:]))):
        s = white_slide(prs)
        text(s, f"Printable Photo Cards · {name}", 0.45, 0.18, 9.4, 0.45, 22, NAVY, True, PP_ALIGN.LEFT)
        text(s, "PRINT • CUT • LAMINATE (optional)", 9.55, 0.2, 3.25, 0.35, 9, MUTED, True, PP_ALIGN.RIGHT)
        bar(s, 0.45, 0.7, 12.85, 0.74, "C9D6E4")
        for i, k in enumerate(keys):
            x = 2.1 + (i % 2) * 5.4
            y = 1.05 + (i // 2) * 3.1
            photo_card(s, k, x, y, 2.55, label=k)
        page_num(s, 47 + idx)

    # 49 sound basket cards ----------------------------------------------------------------------
    s = white_slide(prs)
    text(s, "Sound Basket Cards ( p vs b )", 0.45, 0.18, 9.4, 0.45, 22, NAVY, True, PP_ALIGN.LEFT)
    text(s, "PRINT • CUT", 9.55, 0.2, 3.25, 0.35, 9, MUTED, True, PP_ALIGN.RIGHT)
    bar(s, 0.45, 0.7, 12.85, 0.74, "C9D6E4")
    rect(s, 0.7, 1.05, 5.65, 5.8, P_YELLOW, NAVY, False)
    text(s, "p", 0.7, 1.35, 5.65, 3.3, 130, NAVY, True)
    text(s, "/p/  quiet puff", 1.0, 4.75, 5.05, 0.5, 24, CORAL, True)
    text(s, "pen · pencil · paper", 1.0, 5.45, 5.05, 0.5, 20, INK, True)
    rect(s, 6.98, 1.05, 5.65, 5.8, P_BLUE, NAVY, False)
    text(s, "b", 6.98, 1.35, 5.65, 3.3, 130, NAVY, True)
    text(s, "/b/  throat buzz", 7.28, 4.75, 5.05, 0.5, 24, GREEN, True)
    text(s, "book · bag", 7.28, 5.45, 5.05, 0.5, 20, INK, True)
    page_num(s, 49)

    # 50 story sequence cards ----------------------------------------------------------------------
    s = white_slide(prs)
    text(s, "Pip's First Day · Sequence Cards", 0.45, 0.18, 9.4, 0.45, 22, NAVY, True, PP_ALIGN.LEFT)
    text(s, "PRINT • CUT", 9.55, 0.2, 3.25, 0.35, 9, MUTED, True, PP_ALIGN.RIGHT)
    bar(s, 0.45, 0.7, 12.85, 0.74, "C9D6E4")
    seq = [("1", "Pip packs.", GREEN, "pip"), ("2", "No crayon!", CORAL, "pip_sad"), ("3", "Mia shares.", BLUE, "mia")]
    for i, (n, caption, color, k) in enumerate(seq):
        x = 0.55 + i * 4.25
        rect(s, x, 1.1, 3.85, 5.9, WHITE, NAVY, False)
        rect(s, x, 1.1, 3.85, 0.68, color, color, False)
        text(s, n, x + 1.45, 1.19, 0.95, 0.43, 20, WHITE, True)
        photo(s, k, x + 0.68, 2.05, 2.5)
        text(s, caption, x + 0.4, 5.95, 3.05, 0.6, 24, NAVY, True)
    page_num(s, 50)

    # 51 spare games box ---------------------------------------------------------------------------
    s = white_slide(prs)
    header(s, "🎲  备用游戏盒（时间富余时任选）", PURPLE)
    spare = [
        ("🗣️ Whisper Chain", "两队传悄悄话，最后一人跑去拍照片卡", P_BLUE, BLUE),
        ("🥔 Hot Potato", "实物烫手山芋：音乐停拿着的孩子说词", P_CORAL, CORAL),
        ("🔎 I Spy", "I spy something blue! 孩子找照片墙猜词", P_GREEN, GREEN),
        ("🧩 Odd One Out", "book · pencil · ball —— 哪个不是学校的？", P_YELLOW, ORANGE),
        ("🦊 Duck, Duck, FOX!", "拍头巡回说词，被叫 FOX 的追跑", P_PURPLE, PURPLE),
        ("📸 Memory Match", "照片配对翻翻乐，翻到说词才能拿走", P_BLUE, BLUE),
    ]
    for i, (name, desc, tint, fill) in enumerate(spare):
        x = 0.7 + (i % 3) * 4.15
        y = 1.7 + (i // 3) * 2.5
        rect(s, x, y, 3.7, 2.15, tint, fill, True)
        text(s, name, x + 0.15, y + 0.25, 3.4, 0.55, 22, INK, True)
        text(s, desc, x + 0.2, y + 0.95, 3.3, 1.05, 16, MUTED)
    cue(s, "每个 2-3 分钟，全是已学词的再提取；不要为新词加时", x=0.5, y=6.6)
    page_num(s, 51)

    # 52 teacher prep ------------------------------------------------------------------------------
    s = white_slide(prs)
    header(s, "🛒  明天上课：购买与准备清单", CORAL)
    columns = [
        ("① 外卖 / 文具店", [
            "不透明书包或布袋 1 个（兼作 bag）",
            "book / pencil / pen / crayon /\npaper / eraser 实物各 1–2 份",
            "纸巾 1 包（/p/ 吹气 + /b/ 摸喉咙）",
            "两只苍蝇拍（Swat it）",
            "A4 纸 2–3 张（现场折纸飞机）",
            "骰子 1 个（口头报数也行）",
        ], P_YELLOW, ORANGE),
        ("② 家里找", [
            "小玩偶 1–2 个（Day 3 任务要教玩偶）",
            "白板 / 大张纸 + 笔（eraser 魔术）",
            "两个篮子 / 纸盒（p 篮 + b 篮）",
            "一张挡照片的卡纸（Peek-a-boo）",
            "蓝牙音箱或手机（课件内嵌 4 段音乐）",
            "奖励贴纸（探险卡捉虫贴纸）",
        ], P_BLUE, BLUE),
        ("③ 今晚打印", [
            "照片词卡 Set A / B：Slide 47–48",
            "p / b 声音篮卡：Slide 49",
            "故事排序卡：Slide 50",
            "怪兽图：投屏或打印 1 张",
            "彩印最佳；黑白也可以",
            "没有实体绘本 → 用课件里的",
            "原创故事 Pip's First Day",
        ], P_GREEN, GREEN),
    ]
    for i, (title_, items, tint, fill) in enumerate(columns):
        x = 0.55 + i * 4.25
        rect(s, x, 1.55, 3.95, 4.9, tint, fill, True)
        rect(s, x, 1.55, 3.95, 0.62, fill, fill, False)
        text(s, title_, x + 0.15, 1.66, 3.65, 0.4, 17, WHITE, True)
        for j, item in enumerate(items):
            y = 2.4 + j * 0.58
            circle(s, x + 0.28, y + 0.05, 0.14, fill, fill)
            text(s, item, x + 0.52, y - 0.05, 3.28, 0.56, 11.5, INK, True, PP_ALIGN.LEFT)
    text(s, "最重要：实物书包 + 8 种物品实物 + 照片词卡。其余都能替代，不必网购绘本。",
         1.5, 6.65, 10.3, 0.4, 16, CORAL, True)
    page_num(s, 52)

    # 53 credits --------------------------------------------------------------------------------
    s = white_slide(prs)
    text(s, "素材来源 Credits", 0.5, 0.6, 12.3, 0.9, 44, NAVY, True, PP_ALIGN.LEFT)
    credits = [
        "图片：AI 生成原创素材（12 张，见 scripts/assets/lesson1/），风格统一、无第三方版权",
        "音乐（已内嵌课件）：Kevin MacLeod（incompetech.com），CC BY 4.0 —— Fluffing a Duck / Sneaky Snitch / Monkeys Spinning Monkeys / Carefree",
        "原创故事：Pip's First Day —— 本课件原创角色与情节，可自由用于课堂教学",
        "课程依据：《Our World Starter × Phonics 1 · 24 课整合大纲》L1（Unit 1 My School）",
        "游戏设计参考：幼儿 EFL 的提取练习 / 间隔复习 / 多感官教学（回声山、喂怪兽、闪卡、触卡接力等）",
        "课件生成：scripts/build_starter_lesson1_deck.py（python-pptx），可重复构建",
    ]
    for i, line in enumerate(credits):
        y = 1.9 + i * 0.85
        circle(s, 0.8, y + 0.14, 0.18, GREEN, GREEN)
        text(s, line, 1.25, y, 11.4, 0.75, 17, INK, True, PP_ALIGN.LEFT)
    text(s, "Hello Song / Goodbye Song 建议用 Super Simple Songs 流媒体版本，因版权原因未嵌入。",
         1.25, 6.95, 11.4, 0.4, 14, MUTED, True, PP_ALIGN.LEFT)
    page_num(s, 53)

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(OUTPUT)
    print(f"Wrote {OUTPUT} ({len(prs.slides)} slides)")


if __name__ == "__main__":
    build()
