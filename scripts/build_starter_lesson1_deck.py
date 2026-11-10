#!/usr/bin/env python3
"""Build Lesson 1: My School as a demo-style interactive PPTX deck.

Style follows lesson-plans/demo-slides-starter-toys.pptx:
giant one-word slides, emoji command cards, progressive What's-missing rounds,
140pt phonics letter cards with tissue/throat experiments, embedded mp3 cues,
a Chinese "what we learned" summary and a credits slide.
"""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "lesson-plans" / "our-world-starter-lesson1-my-school.pptx"
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
P_ORANGE = "FDEEDD"
BROWN = "9B6A46"
PINK = "F58DB2"
SKY = "7BDFF2"

EN = "Comic Sans MS"
CN = "微软雅黑"

TRACKS = {
    "enter": ("Fluffing_a_Duck.mp3", "进场音乐"),
    "mystery": ("Sneaky_Snitch.mp3", "神秘袋音乐"),
    "pass": ("Monkeys_Spinning_Monkeys.mp3", "音乐停 → 回答"),
    "show": ("Carefree.mp3", "游戏背景音乐"),
}

NATURAL = {
    "book": (1.15, 0.82),
    "pencil": (1.39, 0.26),
    "pen": (0.18, 1.33),
    "crayon": (0.35, 1.28),
    "paper": (0.92, 1.12),
    "eraser": (1.03, 0.53),
    "desk": (1.35, 1.06),
    "chair": (1.1, 1.65),
    "bag": (2.05, 2.2),
}


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
            shape.adjustments[0] = 0.14
        except (IndexError, ValueError):
            pass
    return shape


def circle(slide, x, y, d, fill, line_color=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    shape.line.color.rgb = rgb(line_color or fill)
    return shape


def bar(slide, x1, y1, x2, y2, color=NAVY, thick=0.06):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(x1), Inches(y1),
        Inches(max(0.02, x2 - x1)), Inches(max(0.02, y2 - y1)),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(color)
    shape.line.color.rgb = rgb(color)
    if thick != 0.06 and y2 - y1 <= 0.08:
        pass
    return shape


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


# ---------------------------------------------------------------- vector art

def draw_book(slide, x, y, scale=1, color=SKY):
    rect(slide, x, y, 1.15 * scale, 0.82 * scale, color, NAVY, False)
    rect(slide, x + 0.53 * scale, y + 0.03 * scale, 0.05 * scale, 0.76 * scale, WHITE, WHITE, False)
    text(slide, "★", x + 0.08 * scale, y + 0.10 * scale, 0.4 * scale, 0.4 * scale,
         max(10, int(15 * scale)), WHITE, True)
    return 1.15 * scale, 0.82 * scale


def draw_pencil(slide, x, y, scale=1, color=YELLOW):
    body = rect(slide, x, y, 1.25 * scale, 0.18 * scale, color, BROWN, False)
    body.rotation = -12
    tip = slide.shapes.add_shape(
        MSO_SHAPE.ISOSCELES_TRIANGLE,
        Inches(x - 0.12 * scale), Inches(y - 0.02 * scale),
        Inches(0.27 * scale), Inches(0.22 * scale),
    )
    tip.rotation = 258
    tip.fill.solid()
    tip.fill.fore_color.rgb = rgb(WHITE)
    tip.line.color.rgb = rgb(BROWN)
    eraser = rect(slide, x + 1.17 * scale, y, 0.22 * scale, 0.18 * scale, PINK, BROWN, False)
    eraser.rotation = -12


def draw_pen(slide, x, y, scale=1, color=BLUE):
    body = rect(slide, x, y, 0.18 * scale, 1.22 * scale, color, NAVY, False)
    rect(slide, x, y - 0.11 * scale, 0.18 * scale, 0.2 * scale, NAVY, NAVY, False)
    clip = rect(slide, x + 0.12 * scale, y + 0.12 * scale, 0.05 * scale, 0.45 * scale, NAVY, NAVY, False)
    body.rotation = 6
    clip.rotation = 6


def draw_crayon(slide, x, y, scale=1, color=CORAL):
    rect(slide, x, y, 0.35 * scale, 1.05 * scale, color, BROWN, False)
    tip = slide.shapes.add_shape(
        MSO_SHAPE.ISOSCELES_TRIANGLE,
        Inches(x), Inches(y - 0.23 * scale),
        Inches(0.35 * scale), Inches(0.3 * scale),
    )
    tip.fill.solid()
    tip.fill.fore_color.rgb = rgb(color)
    tip.line.color.rgb = rgb(BROWN)
    rect(slide, x, y + 0.63 * scale, 0.35 * scale, 0.08 * scale, P_YELLOW, P_YELLOW, False)


def draw_paper(slide, x, y, scale=1):
    rect(slide, x, y, 0.92 * scale, 1.12 * scale, WHITE, NAVY, False)
    for i in range(3):
        rect(slide, x + 0.15 * scale, y + (0.28 + i * 0.2) * scale, 0.6 * scale, 0.03 * scale, SKY, SKY, False)
    circle(slide, x + 0.14 * scale, y + 0.9 * scale, 0.09 * scale, CORAL)


def draw_eraser(slide, x, y, scale=1):
    rect(slide, x, y, 1.03 * scale, 0.53 * scale, PINK, NAVY, True)
    rect(slide, x + 0.62 * scale, y, 0.41 * scale, 0.53 * scale, SKY, NAVY, True)


def draw_desk(slide, x, y, scale=1):
    rect(slide, x, y, 1.35 * scale, 0.26 * scale, BROWN, NAVY, False)
    rect(slide, x + 0.1 * scale, y + 0.25 * scale, 0.15 * scale, 0.8 * scale, BROWN, NAVY, False)
    rect(slide, x + 1.08 * scale, y + 0.25 * scale, 0.15 * scale, 0.8 * scale, BROWN, NAVY, False)
    rect(slide, x + 0.23 * scale, y + 0.3 * scale, 0.85 * scale, 0.36 * scale, P_YELLOW, BROWN, False)


def draw_chair(slide, x, y, scale=1):
    rect(slide, x + 0.15 * scale, y, 0.8 * scale, 0.72 * scale, SKY, NAVY, True)
    rect(slide, x, y + 0.72 * scale, 1.1 * scale, 0.22 * scale, SKY, NAVY, False)
    rect(slide, x + 0.12 * scale, y + 0.9 * scale, 0.13 * scale, 0.75 * scale, NAVY, NAVY, False)
    rect(slide, x + 0.82 * scale, y + 0.9 * scale, 0.13 * scale, 0.75 * scale, NAVY, NAVY, False)


def draw_bag(slide, x, y, scale=1, color=GREEN):
    rect(slide, x, y + 0.35 * scale, 2.05 * scale, 1.85 * scale, color, NAVY, True)
    handle = slide.shapes.add_shape(
        MSO_SHAPE.ARC,
        Inches(x + 0.55 * scale), Inches(y),
        Inches(0.95 * scale), Inches(0.85 * scale),
    )
    handle.fill.background()
    handle.line.color.rgb = rgb(NAVY)
    handle.line.width = Pt(4 * scale)
    rect(slide, x + 0.35 * scale, y + 1.3 * scale, 1.35 * scale, 0.58 * scale, P_YELLOW, NAVY, True)
    circle(slide, x + 0.92 * scale, y + 1.5 * scale, 0.18 * scale, YELLOW, NAVY)


def draw_object(slide, kind, x, y, scale=1, color=None):
    if kind == "book":
        draw_book(slide, x, y, scale, color or SKY)
    elif kind == "pencil":
        draw_pencil(slide, x, y, scale, color or YELLOW)
    elif kind == "pen":
        draw_pen(slide, x, y, scale, color or BLUE)
    elif kind == "crayon":
        draw_crayon(slide, x, y, scale, color or CORAL)
    elif kind == "paper":
        draw_paper(slide, x, y, scale)
    elif kind == "eraser":
        draw_eraser(slide, x, y, scale)
    elif kind == "desk":
        draw_desk(slide, x, y, scale)
    elif kind == "chair":
        draw_chair(slide, x, y, scale)
    elif kind == "bag":
        draw_bag(slide, x, y, scale, color or GREEN)


def draw_centered(slide, kind, cx, cy, box_w, box_h, color=None):
    """Draw object kind centered inside a (box_w x box_h) box at (cx, cy)."""
    nw, nh = NATURAL[kind]
    scale = min(box_w / nw, box_h / nh)
    x = cx - nw * scale / 2
    y = cy - nh * scale / 2
    if kind == "pencil":
        draw_pencil(slide, x + 0.12 * scale, y + 0.04 * scale, scale, color or YELLOW)
    elif kind == "pen":
        draw_pen(slide, x, y + 0.11 * scale, scale, color or BLUE)
    elif kind == "crayon":
        draw_crayon(slide, x, y + 0.23 * scale, scale, color or CORAL)
    else:
        draw_object(slide, kind, x, y, scale, color)


def draw_face(slide, x, y, scale=1, animal="puppy", mood="happy", shirt=GREEN):
    rect(slide, x + 0.35 * scale, y + 1.25 * scale, 1.2 * scale, 1.25 * scale, shirt, NAVY, True)
    circle(slide, x + 0.28 * scale, y + 0.28 * scale, 1.35 * scale, "FFEFD3", NAVY)
    if animal == "puppy":
        ears = []
        for ex in (x, x + 1.35 * scale):
            ear = slide.shapes.add_shape(
                MSO_SHAPE.TEAR, Inches(ex), Inches(y + 0.25 * scale),
                Inches(0.55 * scale), Inches(0.9 * scale),
            )
            ear.rotation = 20 if ex == x else 200
            ear.fill.solid()
            ear.fill.fore_color.rgb = rgb(BROWN)
            ear.line.color.rgb = rgb(NAVY)
            ears.append(ear)
    else:
        for dx in (0.28, 1.15):
            ear = slide.shapes.add_shape(
                MSO_SHAPE.ISOSCELES_TRIANGLE,
                Inches(x + dx * scale), Inches(y),
                Inches(0.5 * scale), Inches(0.58 * scale),
            )
            ear.fill.solid()
            ear.fill.fore_color.rgb = rgb(PINK)
            ear.line.color.rgb = rgb(NAVY)
    circle(slide, x + 0.62 * scale, y + 0.7 * scale, 0.11 * scale, NAVY)
    circle(slide, x + 1.18 * scale, y + 0.7 * scale, 0.11 * scale, NAVY)
    circle(slide, x + 0.88 * scale, y + 0.92 * scale, 0.15 * scale, CORAL)
    if mood == "sad":
        mouth = bar(slide, x + 0.79 * scale, y + 1.19 * scale, x + 1.12 * scale, y + 1.22 * scale, NAVY)
        circle(slide, x + 0.52 * scale, y + 0.86 * scale, 0.1 * scale, SKY)
    else:
        mouth = slide.shapes.add_shape(
            MSO_SHAPE.ARC,
            Inches(x + 0.72 * scale), Inches(y + 0.94 * scale),
            Inches(0.5 * scale), Inches(0.38 * scale),
        )
        mouth.fill.background()
        mouth.line.color.rgb = rgb(NAVY)
        mouth.line.width = Pt(2)


def speech(slide, value, x, y, w, h, fill=WHITE, border=NAVY, size=22, color=NAVY):
    rect(slide, x, y, w, h, fill, border, True)
    text(slide, value, x + 0.15, y + 0.08, w - 0.3, h - 0.16, size, color, True)


# ------------------------------------------------------------------- slides

def word_slide(prs, n, kind, word, hint):
    slide = white_slide(prs)
    draw_centered(slide, kind, W / 2, 3.25, 5.6, 4.5)
    text(slide, word, 0.5, 5.75, 12.3, 1.2, 72, NAVY, True)
    text(slide, hint, 0.5, 7.02, 11.0, 0.38, 16, MUTED, align=PP_ALIGN.LEFT)
    page_num(slide, n)
    return slide


def grid_card(slide, kind, label, x, y, w=4.0, h=2.6, show_label=True, missing=False):
    rect(slide, x, y, w, h, WHITE, "C9D6E4", True)
    if missing:
        text(slide, "?", x, y + 0.1, w, h - 0.75, 120, CORAL, True)
    else:
        draw_centered(slide, kind, x + w / 2, y + 1.15, w - 0.7, 1.75)
    if show_label:
        text(slide, "" if missing else label, x, y + h - 0.62, w, 0.55, 22, NAVY, True)


def game_card(slide, emoji, title, sub, x, y, w, h, fill, tint):
    rect(slide, x, y, w, h, tint, fill, True)
    text(slide, emoji, x, y + 0.18, w, h * 0.52, int(h * 28), align=PP_ALIGN.CENTER)
    text(slide, title, x + 0.1, y + h * 0.62, w - 0.2, 0.5, 24, INK, True)
    text(slide, sub, x + 0.1, y + h * 0.62 + 0.5, w - 0.2, 0.42, 16, MUTED)


def build():
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    prs.core_properties.title = "Lesson 1: My School!"
    prs.core_properties.subject = "Our World Starter U1 × Phonics 1 (L1)"
    prs.core_properties.author = "EnglishLearner"
    prs.core_properties.keywords = "My School, phonics, /p/, /b/, games, Pip"

    # 1 cover -------------------------------------------------------------
    s = white_slide(prs)
    text(s, "My School!", 0.5, 0.5, 12.3, 1.6, 80, NAVY, True, PP_ALIGN.LEFT)
    text(s, "第一课 ｜ Our World Starter · Unit 1 ｜ 90 分钟", 0.55, 2.05, 11, 0.5, 20, MUTED, align=PP_ALIGN.LEFT)
    draw_centered(s, "book", 2.6, 4.6, 4.0, 4.0)
    draw_centered(s, "pencil", 6.25, 3.35, 2.1, 2.0)
    draw_centered(s, "pen", 8.9, 3.35, 1.6, 2.0)
    draw_centered(s, "crayon", 11.6, 3.35, 1.8, 2.0)
    draw_centered(s, "paper", 6.25, 5.6, 1.9, 1.9)
    draw_centered(s, "eraser", 8.9, 5.6, 2.4, 1.5)
    draw_centered(s, "chair", 11.75, 5.55, 1.9, 2.2)
    add_music(s, "enter")
    page_num(s, 1)

    # 2 hello --------------------------------------------------------------
    s = white_slide(prs)
    draw_face(s, 0.9, 1.5, 2.3, "puppy", "happy", GREEN)
    rect(s, 5.8, 1.4, 7.0, 2.6, BLUE, BLUE)
    text(s, "Hello! I'm Pip.\nWhat's your name?", 6.0, 1.55, 6.6, 2.3, 40, WHITE, True)
    rect(s, 6.8, 4.6, 5.0, 1.5, YELLOW, YELLOW)
    text(s, "I'm ____ .", 7.0, 4.75, 4.6, 1.2, 44, INK, True)
    cue(s, "布偶示范一遍 → 传球逐个问答，要求说整句；第 3 轮加 Nice to meet you!", x=0.5)
    page_num(s, 2)

    # 3 routine ------------------------------------------------------------
    s = white_slide(prs)
    header(s, "🙋  Pip says…", GREEN)
    commands = [("🧍", "Stand up"), ("🧎", "Sit down"), ("🙌", "Hands up"), ("👀", "Eyes on me")]
    for i, (emoji, cmd) in enumerate(commands):
        x = 0.7 + (i % 2) * 6.2
        y = 1.6 + (i // 2) * 2.6
        text(s, emoji, x, y, 5.8, 2.2, 120)
        rect(s, x, y + 2.35, 5.8, 1.1, BLUE if i % 2 == 0 else CORAL)
        text(s, cmd, x, y + 2.5, 5.8, 0.8, 40, WHITE, True)
    text(s, "Hands up = I want to talk  🗣️      做 2–3 轮，越喊越快！", 1.0, 6.45, 11.3, 0.5, 24, MUTED)
    page_num(s, 3)

    # 4 mystery bag --------------------------------------------------------
    s = white_slide(prs)
    header(s, "🎒  What's in the bag?", CORAL)
    draw_centered(s, "bag", 6.66, 3.8, 5.0, 4.6)
    text(s, "摇一摇 → 听一听 → 猜一猜", 4.6, 6.15, 4.1, 0.45, 20, MUTED)
    add_music(s, "mystery")
    cue(s, "不透明书包；每次请 1 个孩子伸手摸 1 件，先猜再揭晓（对应下一页）")
    page_num(s, 4)

    # 5-12 one-word slides --------------------------------------------------
    words = [
        ("book", "第 1 件：从神秘袋里摸出来！跟读 2 遍"),
        ("pencil", "拍手数一数：pen-cil，2 个音节"),
        ("pen", "举起来转一转：It's a pen!"),
        ("crayon", "假装在纸上画一画"),
        ("paper", "两手捧住，轻轻扇一扇"),
        ("eraser", "在桌上做擦一擦的动作"),
        ("desk", "伸手摸一摸自己的桌子"),
        ("chair", "坐下 → 站起，边做边说"),
    ]
    for i, (kind, hint) in enumerate(words):
        word_slide(prs, 5 + i, kind, kind, hint)

    # 13 show me ------------------------------------------------------------
    s = white_slide(prs)
    header(s, "🙌  Show me… FAST!", YELLOW)
    for i, kind in enumerate(["book", "pencil", "pen", "crayon"]):
        x = 0.9 + i * 3.1
        rect(s, x, 1.7, 2.7, 2.7, P_BLUE if i % 2 == 0 else P_YELLOW, WHITE, True)
        draw_centered(s, kind, x + 1.35, 3.05, 2.0, 2.2)
        text(s, kind, x, 4.55, 2.7, 0.55, 26, NAVY, True)
    rect(s, 2.8, 5.5, 7.7, 0.95, NAVY, NAVY)
    text(s, "Ready?  3 — 2 — 1 — SHOW ME!", 3.0, 5.62, 7.3, 0.7, 32, WHITE, True)
    cue(s, "孩子用自己桌上的实物；打乱顺序喊 8–10 轮，最后 2 轮加速")
    page_num(s, 13)

    # 14 listen and do ------------------------------------------------------
    s = white_slide(prs)
    header(s, "🎧  Listen and do!", BLUE)
    actions = [
        ("📖", "Open your book", P_BLUE, BLUE),
        ("✏️", "Pick up a pencil", P_YELLOW, ORANGE),
        ("🖐️", "Touch the desk", P_GREEN, GREEN),
        ("🪑", "Sit on the chair", P_CORAL, CORAL),
    ]
    for i, (emoji, cmd, tint, fill) in enumerate(actions):
        x = 0.7 + (i % 2) * 6.2
        y = 1.65 + (i // 2) * 2.5
        rect(s, x, y, 5.8, 2.15, tint, fill, True)
        text(s, emoji, x + 0.2, y + 0.15, 1.9, 1.85, 60)
        text(s, cmd, x + 2.2, y + 0.55, 3.5, 1.05, 30, INK, True)
    text(s, "做完动作 → 再跟读一遍 ｜ 两步指令：Pick up a pencil and touch the desk!", 1.0, 6.55, 11.3, 0.5, 22, MUTED)
    page_num(s, 14)

    # 15-17 what's missing x3 ----------------------------------------------
    grid_words = ["book", "pencil", "pen", "paper", "desk", "chair"]
    positions = [(0.6, 1.5), (4.7, 1.5), (8.8, 1.5), (0.6, 4.35), (4.7, 4.35), (8.8, 4.35)]

    s = white_slide(prs)
    header(s, "👀  What's missing?", CORAL)
    for kind, (x, y) in zip(grid_words, positions):
        grid_card(s, kind, kind, x, y)
    text(s, "10… 9… 8…  👀  仔细看！", 3.6, 7.0, 6.1, 0.45, 26, CORAL, True)
    page_num(s, 15)

    for round_no, missing in ((1, {"pencil"}), (2, {"pencil", "chair"})):
        s = white_slide(prs)
        header(s, "👀  What's missing?", CORAL)
        for kind, (x, y) in zip(grid_words, positions):
            grid_card(s, kind, kind, x, y, missing=kind in missing)
        if round_no == 1:
            text(s, "什么不见了？大声喊出来！", 3.6, 7.0, 6.1, 0.45, 26, CORAL, True)
            cue(s, "先接受单词，再引导整句 It's a pencil!", x=0.5, y=7.02)
        else:
            text(s, "两个都不见了！You have 10 seconds!", 3.0, 7.0, 7.3, 0.45, 26, CORAL, True)
            cue(s, "挑战：把两句连起来说", x=0.5, y=7.02)
        page_num(s, 15 + round_no)

    # 18 sentence model -----------------------------------------------------
    s = white_slide(prs)
    header(s, "❓  What is it?", BLUE)
    draw_centered(s, "pencil", 3.1, 3.7, 4.2, 4.2)
    rect(s, 5.5, 2.2, 7.2, 1.8, GREEN, GREEN)
    text(s, "It's a pencil.", 5.7, 2.4, 6.8, 1.4, 54, WHITE, True)
    text(s, "☝️    ✌️    🤟", 5.5, 4.3, 7.2, 1.0, 54)
    text(s, "It's  –  a  –  pen·cil  （手指打节拍，一词一拍）", 5.5, 5.5, 7.2, 0.6, 26, MUTED)
    cue(s, "布偶问，老师答；再师生互换。只教这一个句型，要求说整句", x=0.5)
    page_num(s, 18)

    # 19 four voices ----------------------------------------------------------
    s = white_slide(prs)
    header(s, "🎭  Ask it four ways!", PURPLE)
    voices = [
        ("🗣️", "GIANT", "巨人大声音", CORAL, P_CORAL),
        ("🐭", "MOUSE", "老鼠小声音", BLUE, P_BLUE),
        ("🤖", "ROBOT", "机器人声", GREEN, P_GREEN),
        ("🤫", "SECRET", "悄悄话", PURPLE, P_PURPLE),
    ]
    for i, (emoji, name, sub, fill, tint) in enumerate(voices):
        x = 0.75 + i * 3.1
        rect(s, x, 1.7, 2.75, 3.1, tint, fill, True)
        text(s, emoji, x, 1.9, 2.75, 1.5, 72)
        text(s, name, x, 3.45, 2.75, 0.5, 26, INK, True)
        text(s, sub, x, 3.95, 2.75, 0.45, 18, MUTED)
    rect(s, 2.9, 5.35, 7.5, 1.05, NAVY, NAVY)
    text(s, "What is it?  —  It's a ___ !", 3.1, 5.48, 7.1, 0.75, 30, WHITE, True)
    cue(s, "老师指卡片定声音，全班一起问；布偶故意答错一次，让孩子纠正")
    page_num(s, 19)

    # 20 pass the bag ---------------------------------------------------------
    s = white_slide(prs)
    header(s, "🎒  Pass the Bag!", CORAL)
    draw_centered(s, "bag", 3.0, 3.8, 4.2, 4.4)
    rect(s, 6.5, 2.2, 6.3, 2.6, BLUE, BLUE)
    text(s, "What is it?\nIt's a ____ .", 6.7, 2.4, 5.9, 2.2, 40, WHITE, True)
    text(s, "拿到书包的孩子从袋里摸 1 件 → 问答 → 传给下一个", 6.6, 5.1, 6.1, 0.55, 20, MUTED)
    add_music(s, "pass")
    page_num(s, 20)

    # 21 peek-a-boo ------------------------------------------------------------
    s = white_slide(prs)
    header(s, "👀  Peek-a-boo!  What is it?", GREEN)
    for i, (kind, tint, fill) in enumerate((("pen", P_YELLOW, ORANGE), ("eraser", P_BLUE, BLUE))):
        x = 1.0 + i * 6.2
        rect(s, x, 1.7, 5.4, 3.7, tint, fill, True)
        text(s, "?", x + 1.2, 1.85, 3.0, 3.2, 140, fill, True)
        draw_centered(s, kind, x + 4.55, 4.75, 1.7, 1.4)
    rect(s, 2.9, 5.9, 7.5, 0.85, NAVY, NAVY)
    text(s, "只露一个角，也要说整句：It's a ___ !", 3.1, 6.0, 7.1, 0.6, 26, WHITE, True)
    cue(s, "把实物藏在卡纸后只露一角；猜对全班鼓掌 3 下")
    page_num(s, 21)

    # 22 color challenge ---------------------------------------------------------
    s = white_slide(prs)
    header(s, "⭐  Add one color!", YELLOW)
    rect(s, 1.0, 1.8, 5.1, 3.3, P_CORAL, CORAL, True)
    draw_centered(s, "pen", 3.55, 3.3, 2.6, 2.4, color=CORAL)
    text(s, "It's a red pen.", 1.2, 5.25, 4.7, 0.6, 30, INK, True)
    rect(s, 7.2, 1.8, 5.1, 3.3, P_BLUE, BLUE, True)
    draw_centered(s, "book", 9.75, 3.3, 3.2, 2.4, color=SKY)
    text(s, "It's a blue book.", 7.4, 5.25, 4.7, 0.6, 30, INK, True)
    text(s, "CORE ✓  It's a pen.          CHALLENGE ★  It's a red pen.", 2.2, 6.25, 9.0, 0.5, 22, MUTED)
    cue(s, "会说的孩子加颜色，不会的孩子说原句同样得贴纸", x=0.5)
    page_num(s, 22)

    # 23 freeze dance ----------------------------------------------------------
    s = white_slide(prs)
    header(s, "🎶  Freeze Dance!", PURPLE)
    text(s, "MOVE 💃", 0.9, 2.0, 3.6, 1.2, 54, CORAL, True)
    text(s, "→", 4.6, 2.05, 1.0, 1.0, 44, MUTED, True)
    text(s, "FREEZE 🧊", 5.7, 2.0, 3.9, 1.2, 54, BLUE, True)
    text(s, "→", 9.7, 2.05, 1.0, 1.0, 44, MUTED, True)
    text(s, "MIME 🎭", 10.8, 2.0, 2.5, 1.2, 54, GREEN, True)
    rect(s, 1.6, 3.9, 10.1, 1.15, YELLOW, YELLOW)
    text(s, "BOOK!   PENCIL!   DESK!   CHAIR!", 1.8, 4.05, 9.7, 0.8, 34, NAVY, True)
    text(s, "音乐停 → 听到哪个词，就做那个词的动作并定住！", 2.2, 5.35, 8.9, 0.55, 24, MUTED)
    add_music(s, "show")
    page_num(s, 23)

    # 24 letter p ---------------------------------------------------------------
    s = white_slide(prs)
    rect(s, 0.5, 0.6, 4.8, 3.2, YELLOW, YELLOW)
    text(s, "P p", 0.7, 0.85, 4.4, 2.7, 140, NAVY, True)
    text(s, "/p/  /p/  /p/", 0.6, 4.05, 4.6, 1.2, 44, CORAL, True)
    for i, kind in enumerate(["pen", "pencil", "paper"]):
        x = 5.6 + i * 2.5
        rect(s, x, 1.2, 2.3, 3.6, P_BLUE if i % 2 else P_YELLOW, WHITE, True)
        draw_centered(s, kind, x + 1.15, 2.5, 1.8, 2.2)
        text(s, kind, x, 4.05, 2.3, 0.6, 34, NAVY, True)
    text(s, "🤫 纸巾实验：/p/ 能吹动纸巾，喉咙不振动（字母名 pee ≠ 字母音 /p/）", 0.7, 5.7, 12.0, 0.6, 24, MUTED)
    text(s, "今天只听开头音：pen · pencil · paper 都住着 /p/", 0.7, 6.4, 12.0, 0.5, 20, MUTED)
    page_num(s, 24)

    # 25 letter b ---------------------------------------------------------------
    s = white_slide(prs)
    rect(s, 0.5, 0.6, 4.8, 3.2, BLUE, BLUE)
    text(s, "B b", 0.7, 0.85, 4.4, 2.7, 140, WHITE, True)
    text(s, "/b/  /b/  /b/", 0.6, 4.05, 4.6, 1.2, 44, GREEN, True)
    for i, kind in enumerate(["book", "bag"]):
        x = 5.7 + i * 3.6
        rect(s, x, 1.2, 3.3, 3.6, P_GREEN if i % 2 else P_BLUE, WHITE, True)
        draw_centered(s, kind, x + 1.65, 2.5, 2.7, 2.2)
        text(s, kind, x, 4.05, 3.3, 0.6, 34, NAVY, True)
    text(s, "✋ 摸喉咙：/b/ 会振动，/p/ 不振动 —— 同样双唇，差别在喉咙", 0.7, 5.7, 12.0, 0.6, 24, MUTED)
    text(s, "book 和 bag（书包）都是 /b/ 开头！", 0.7, 6.4, 12.0, 0.5, 20, MUTED)
    page_num(s, 25)

    # 26 pop or bounce -----------------------------------------------------------
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
    cue(s, "随机喊词：pen / book / pencil / bag / paper / desk / chair / crayon，孩子做对应动作")
    page_num(s, 26)

    # 27 swat it ---------------------------------------------------------------
    s = white_slide(prs)
    header(s, "🪰  Swat it!", YELLOW)
    rect(s, 1.0, 1.8, 4.8, 3.6, P_YELLOW, ORANGE, True)
    text(s, "P p", 1.2, 2.1, 4.4, 3.0, 150, NAVY, True)
    rect(s, 7.5, 1.8, 4.8, 3.6, P_BLUE, BLUE, True)
    text(s, "B b", 7.7, 2.1, 4.4, 3.0, 150, NAVY, True)
    text(s, "🪰", 5.95, 2.5, 1.6, 2.0, 90)
    text(s, "Listen  →  Turn  →  Swat!  →  Say the sound", 1.0, 5.8, 11.3, 0.7, 28, NAVY, True)
    cue(s, "两只苍蝇拍；听到词跑到对应首音卡片拍下并大声说 /p/ 或 /b/", x=0.5, y=6.75)
    page_num(s, 27)

    # 28 robot talk --------------------------------------------------------------
    s = white_slide(prs)
    header(s, "🤖  Robot Talk!", BLUE)
    robots = [
        ("/p/  -  /e/  -  /n/", "pen!", "pen", P_YELLOW),
        ("/b/  -  /u/  -  /k/", "book!", "book", P_BLUE),
        ("/d/  -  /e/  -  /s/  -  /k/", "desk!", "desk", P_GREEN),
    ]
    for i, (seg, whole, kind, tint) in enumerate(robots):
        y = 1.65 + i * 1.5
        rect(s, 0.8, y, 11.7, 1.25, tint, WHITE, True)
        text(s, "🤖 ROBOT:   " + seg, 1.1, y + 0.2, 6.2, 0.85, 26, NAVY, True, PP_ALIGN.LEFT)
        text(s, "→", 7.35, y + 0.2, 0.7, 0.85, 30, CORAL, True)
        text(s, whole, 8.1, y + 0.2, 1.8, 0.85, 30, GREEN, True)
        draw_centered(s, kind, 11.4, y + 0.62, 1.5, 0.95)
    text(s, "机器人慢速说 → 你快快说成一个词！（纯听音，不看拼写）", 1.5, 6.25, 10.3, 0.5, 20, MUTED)
    text(s, "下节课 L2 预告：short /a/ 来啦，我们拼 pat · bat · tap！", 1.5, 6.8, 10.3, 0.45, 18, CORAL, True)
    page_num(s, 28)

    # 29 sound basket relay -------------------------------------------------------
    s = white_slide(prs)
    header(s, "🧺  Sound Basket Relay!", GREEN)
    steps = [
        ("1  PICK", "抽 1 件实物", P_YELLOW, ORANGE),
        ("2  SAY", "说首音 /p/ 或 /b/", P_BLUE, BLUE),
        ("3  SPEAK", "说整句 It's a ___.", P_GREEN, GREEN),
        ("4  SORT", "投进 p 篮或 b 篮", P_PURPLE, PURPLE),
    ]
    for i, (title, sub, tint, fill) in enumerate(steps):
        x = 0.7 + i * 3.1
        rect(s, x, 1.7, 2.75, 2.6, tint, fill, True)
        text(s, title, x + 0.1, 2.0, 2.55, 0.6, 26, INK, True)
        text(s, sub, x + 0.15, 2.75, 2.45, 1.2, 18, MUTED)
    rect(s, 1.6, 4.75, 4.6, 1.5, P_YELLOW, ORANGE, True)
    text(s, "🧺  p  篮", 1.8, 4.95, 4.2, 1.1, 40, NAVY, True)
    rect(s, 7.1, 4.75, 4.6, 1.5, P_BLUE, BLUE, True)
    text(s, "🧺  b  篮", 7.3, 4.95, 4.2, 1.1, 40, NAVY, True)
    cue(s, "两人一组比赛；投错篮不要批评，回来说对首音即可重投", x=0.5, y=6.55)
    page_num(s, 29)

    # 30 story cover ------------------------------------------------------------
    s = white_slide(prs)
    header(s, "📖  Story Time:  Pip's First Day", CORAL)
    draw_face(s, 1.0, 1.7, 1.7, "puppy", "happy", GREEN)
    draw_centered(s, "bag", 5.0, 3.6, 3.3, 3.6, color=CORAL)
    speech(s, "What will Pip pack?", 7.3, 1.7, 4.9, 1.0, P_GREEN, GREEN, 28, NAVY)
    for i, kind in enumerate(["book", "pencil", "crayon"]):
        draw_centered(s, kind, 8.3 + i * 1.8, 3.6, 1.6, 1.8)
    text(s, "指一指，猜一猜：Pip 会装哪几件？", 7.3, 4.7, 4.9, 0.5, 22, CORAL, True)
    cue(s, "没有实体绘本也可用本原创故事；先预测，再翻页验证", x=0.5)
    page_num(s, 30)

    # 31 story 1 ----------------------------------------------------------------
    s = white_slide(prs)
    header(s, "📖  1  Pip packs for school.", GREEN)
    draw_face(s, 0.7, 1.8, 1.35, "puppy", "happy", GREEN)
    draw_centered(s, "bag", 4.2, 3.7, 3.0, 3.4, color=CORAL)
    for i, kind in enumerate(["book", "pencil", "pen", "paper"]):
        draw_centered(s, kind, 7.6 + (i % 2) * 2.6, 2.4 + (i // 2) * 2.6, 2.2, 2.3)
    rect(s, 7.3, 5.35, 5.4, 1.35, P_YELLOW, ORANGE, True)
    text(s, "A book, a pencil, a pen,\nand paper.", 7.5, 5.5, 5.0, 1.05, 26, NAVY, True)
    cue(s, "读一遍；孩子只跟读彩色实物词，再问 What is in the bag?", x=0.5)
    page_num(s, 31)

    # 32 story 2 ----------------------------------------------------------------
    s = white_slide(prs)
    header(s, "📖  2  Oh no! No crayon.", CORAL)
    draw_face(s, 0.9, 1.9, 1.6, "puppy", "sad", GREEN)
    draw_centered(s, "paper", 4.6, 3.5, 2.6, 3.2)
    rect(s, 7.0, 1.8, 5.5, 3.4, P_CORAL, CORAL, True)
    text(s, "NO", 7.3, 2.1, 2.0, 1.0, 44, CORAL, True)
    draw_centered(s, "crayon", 10.0, 3.4, 1.7, 2.4, color=CORAL)
    bar(s, 8.9, 2.5, 10.9, 4.3, CORAL)
    speech(s, "I need a crayon.", 1.0, 5.35, 4.9, 0.9, WHITE, NAVY, 26, NAVY)
    speech(s, "How does Pip feel?", 6.3, 5.35, 4.9, 0.9, P_YELLOW, ORANGE, 26, NAVY)
    cue(s, "先停在这里让孩子找 crayon 在不在；再演 Pip 的表情 sad", x=0.5)
    page_num(s, 32)

    # 33 story 3 ----------------------------------------------------------------
    s = white_slide(prs)
    header(s, "📖  3  Mia shares her crayon.", BLUE)
    draw_face(s, 0.8, 1.9, 1.5, "puppy", "happy", GREEN)
    draw_face(s, 10.0, 1.9, 1.5, "cat", "happy", CORAL)
    draw_centered(s, "crayon", 6.66, 3.3, 1.9, 2.6, color=CORAL)
    speech(s, "Here you are!", 7.3, 1.75, 2.6, 0.85, P_CORAL, CORAL, 24, NAVY)
    speech(s, "Thank you!", 3.3, 1.9, 2.6, 0.85, P_GREEN, GREEN, 24, NAVY)
    rect(s, 3.0, 5.5, 7.3, 0.95, P_YELLOW, ORANGE, True)
    text(s, "A crayon + a kind friend\n= a happy first day!", 3.2, 5.6, 6.9, 0.75, 22, NAVY, True)
    cue(s, "两人一组用一支蜡笔演 Here you are! / Thank you!", x=0.5)
    page_num(s, 33)

    # 34 story order game --------------------------------------------------------
    s = white_slide(prs)
    header(s, "🃏  First — Then — Finally", PURPLE)
    order = [
        ("FIRST", "Pip packs.", GREEN, "bag"),
        ("THEN", "No crayon!", CORAL, "crayon"),
        ("FINALLY", "Mia shares.", BLUE, "chair"),
    ]
    for i, (marker, caption, fill, kind) in enumerate(order):
        x = 0.8 + i * 4.15
        rect(s, x, 1.7, 3.6, 4.3, WHITE, fill, True)
        rect(s, x, 1.7, 3.6, 0.62, fill, fill, False)
        text(s, marker, x + 0.15, 1.8, 3.3, 0.4, 16, WHITE, True)
        draw_centered(s, kind, x + 1.8, 3.6, 2.9, 2.3, color=CORAL if kind == "crayon" else None)
        text(s, caption, x + 0.2, 5.15, 3.2, 0.65, 24, NAVY, True)
    text(s, "3 个孩子各拿 1 张场景卡，全班喊出正确顺序再站好！", 1.5, 6.35, 10.3, 0.5, 22, MUTED)
    cue(s, "场景卡打印页在 Slide 43；说错顺序不扣分，重来一次", x=0.5, y=6.9)
    page_num(s, 34)

    # 35 story detectives ---------------------------------------------------------
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
    cue(s, "每个孩子至少答 1 题：指 → 说词 → 短句 → 表演", x=0.5, y=6.55)
    page_num(s, 35)

    # 36 exit ticket --------------------------------------------------------------
    s = white_slide(prs)
    header(s, "🚪  Choose one door!", CORAL)
    doors = [
        ("DOOR 1", "Say 3 words", "book · pen · chair", P_BLUE, BLUE),
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
    cue(s, "出门考：每人 30–40 秒，记录 独立 / 提示后 / 需帮助", x=0.5, y=6.35)
    page_num(s, 36)

    # 37 take home -----------------------------------------------------------------
    s = white_slide(prs)
    header(s, "🏠  20-second School Bag Challenge", GREEN)
    draw_centered(s, "bag", 2.4, 3.8, 3.2, 4.2)
    steps = [
        ("1", "Pick 3 school things.", P_YELLOW, ORANGE),
        ("2", "Say: It's a ___ , It's a ___ , It's a ___ .", P_BLUE, BLUE),
        ("3", "Bonus: /p/ /p/ pen!", P_CORAL, CORAL),
    ]
    for i, (n, phrase, tint, fill) in enumerate(steps):
        y = 1.9 + i * 1.35
        circle(s, 5.2, y, 0.8, fill, NAVY)
        text(s, n, 5.2, y + 0.18, 0.8, 0.4, 20, WHITE, True)
        rect(s, 6.3, y, 6.4, 1.05, tint, fill, True)
        text(s, phrase, 6.5, y + 0.1, 6.0, 0.85, 22, INK, True, PP_ALIGN.LEFT)
    text(s, "Example: It's a book. It's a pencil. It's a pen.", 5.2, 6.0, 7.5, 0.5, 20, MUTED)
    cue(s, "不要求整本朗读！20–30 秒小视频发群即可", x=5.2, y=6.6)
    page_num(s, 37)

    # 38 today we learned (Chinese summary) ------------------------------------------
    s = white_slide(prs)
    header(s, "🏠  今天我们学了", NAVY)
    rows = [
        ("🧾 单词", "book, pencil, pen, crayon, paper, eraser, desk, chair", P_BLUE, BLUE),
        ("💬 句子", "What is it? — It's a ___.   挑战：It's a red pen.", P_YELLOW, ORANGE),
        ("🔤 字母的声音", "P 说 /p/（pen · pencil · paper）  B 说 /b/（book · bag）—— 字母有名字，也有声音", P_GREEN, GREEN),
        ("🎒 回家任务", "20 秒 School Bag Challenge：挑 3 件文具说整句，录视频发群", P_CORAL, CORAL),
    ]
    for i, (label, content, tint, fill) in enumerate(rows):
        y = 1.55 + i * 1.35
        rect(s, 0.6, y, 3.0, 1.1, fill, fill)
        text(s, label, 0.75, y + 0.12, 2.7, 0.85, 26, WHITE, True)
        rect(s, 3.8, y, 9.0, 1.1, tint, fill, True)
        text(s, content, 4.0, y + 0.1, 8.6, 0.9, 22, INK, True, PP_ALIGN.LEFT)
    cue(s, "拍照发家长群，附一句 Today I can… 的成长说明", x=0.6, y=7.02)
    page_num(s, 38)

    # 39 goodbye -------------------------------------------------------------------
    s = white_slide(prs)
    draw_face(s, 0.8, 1.4, 2.2, "puppy", "happy", GREEN)
    rect(s, 6.2, 2.2, 6.6, 2.8, BLUE, BLUE)
    text(s, "Goodbye, Pip!\n👋", 6.4, 2.4, 6.2, 2.4, 50, WHITE, True)
    text(s, "下次课预告：认识 short /a/，拼出第一批 CVC：pat · bat · tap！", 2.2, 5.6, 8.9, 0.55, 22, MUTED)
    text(s, "（Goodbye Song 用课件底部 Credits 推荐的流媒体版本）", 2.2, 6.25, 8.9, 0.4, 16, MUTED)
    page_num(s, 39)

    # 40-41 printable vocab cards -----------------------------------------------
    print_words_a = [("book", SKY), ("pencil", YELLOW), ("pen", BLUE), ("crayon", CORAL)]
    print_words_b = [("paper", PURPLE), ("eraser", PINK), ("desk", BROWN), ("chair", GREEN)]
    for idx, (name, words) in enumerate((("Set A", print_words_a), ("Set B", print_words_b))):
        s = white_slide(prs)
        text(s, f"Printable Vocabulary Cards · {name}", 0.45, 0.18, 9.4, 0.45, 22, NAVY, True, PP_ALIGN.LEFT)
        text(s, "PRINT • CUT • LAMINATE (optional)", 9.55, 0.2, 3.25, 0.35, 9, MUTED, True, PP_ALIGN.RIGHT)
        bar(s, 0.45, 0.7, 12.85, 0.74, "C9D6E4")
        for i, (kind, color) in enumerate(words):
            x = 0.5 + (i % 2) * 6.42
            y = 0.95 + (i // 2) * 3.05
            rect(s, x, y, 5.9, 2.75, WHITE, NAVY, False)
            rect(s, x, y, 0.18, 2.75, color, color, False)
            draw_centered(s, kind, x + 1.55, y + 1.38, 2.1, 1.9)
            text(s, kind, x + 2.85, y + 0.78, 2.9, 0.8, 30, NAVY, True)
            text(s, "What is it?  It's a ______.", x + 2.65, y + 1.65, 3.25, 0.42, 12, MUTED)
        page_num(s, 40 + idx)

    # 42 printable p/b basket cards ------------------------------------------------
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
    page_num(s, 42)

    # 43 story sequence cards --------------------------------------------------------
    s = white_slide(prs)
    text(s, "Pip's First Day · Sequence Cards", 0.45, 0.18, 9.4, 0.45, 22, NAVY, True, PP_ALIGN.LEFT)
    text(s, "PRINT • CUT", 9.55, 0.2, 3.25, 0.35, 9, MUTED, True, PP_ALIGN.RIGHT)
    bar(s, 0.45, 0.7, 12.85, 0.74, "C9D6E4")
    seq = [("1", "Pip packs.", GREEN), ("2", "No crayon!", CORAL), ("3", "Mia shares.", BLUE)]
    for i, (n, caption, color) in enumerate(seq):
        x = 0.55 + i * 4.25
        rect(s, x, 1.1, 3.85, 5.9, WHITE, NAVY, False)
        rect(s, x, 1.1, 3.85, 0.68, color, color, False)
        text(s, n, x + 1.45, 1.19, 0.95, 0.43, 20, WHITE, True)
        if i == 0:
            draw_face(s, x + 0.55, 2.15, 0.95, "puppy", "happy", GREEN)
            draw_centered(s, "bag", x + 2.9, 3.6, 1.6, 2.2, color=CORAL)
        elif i == 1:
            draw_face(s, x + 0.7, 2.15, 1.15, "puppy", "sad", GREEN)
            draw_centered(s, "crayon", x + 2.95, 3.6, 1.3, 2.2, color=CORAL)
            bar(s, x + 2.55, 2.7, x + 3.5, 4.4, CORAL)
        else:
            draw_face(s, x + 0.1, 2.2, 0.85, "puppy", "happy", GREEN)
            draw_face(s, x + 2.15, 2.2, 0.85, "cat", "happy", CORAL)
            draw_centered(s, "crayon", x + 1.9, 4.2, 1.2, 1.6, color=CORAL)
        text(s, caption, x + 0.4, 5.75, 3.05, 0.6, 24, NAVY, True)
    page_num(s, 43)

    # 44 teacher prep ------------------------------------------------------------
    s = white_slide(prs)
    header(s, "🛒  明天上课：购买与准备清单", CORAL)
    columns = [
        ("① 外卖 / 文具店", [
            "不透明书包或布袋 1 个（兼作 /b/ bag）",
            "book / pencil / pen / crayon /\npaper / eraser 各 2–4 份",
            "便利贴 / 美纹纸 / 奖励贴纸",
            "纸巾 1 包（/p/ 吹气 + /b/ 摸喉咙）",
            "两只苍蝇拍（Swat it 游戏）",
        ], P_YELLOW, ORANGE),
        ("② 家里找", [
            "小玩偶 1–2 个（扮演 Pip / Mia）",
            "蓝牙音箱或手机（课件已内嵌音乐）",
            "两个篮子 / 纸盒（p 篮 + b 篮）",
            "1 张挡住实物的卡纸（Peek-a-boo）",
            "红、蓝各 1 件（颜色挑战）",
        ], P_BLUE, BLUE),
        ("③ 今晚打印", [
            "词卡 Set A / Set B：Slide 40–41",
            "p / b 声音篮卡：Slide 42",
            "故事排序卡：Slide 43",
            "彩印最佳；黑白打印后手涂也行",
            "没有实体绘本 → 直接用课件里的",
            "原创故事 Pip's First Day",
        ], P_GREEN, GREEN),
    ]
    for i, (title, items, tint, fill) in enumerate(columns):
        x = 0.55 + i * 4.25
        rect(s, x, 1.55, 3.95, 4.9, tint, fill, True)
        rect(s, x, 1.55, 3.95, 0.62, fill, fill, False)
        text(s, title, x + 0.15, 1.66, 3.65, 0.4, 17, WHITE, True)
        for j, item in enumerate(items):
            y = 2.4 + j * 0.68
            circle(s, x + 0.28, y + 0.06, 0.16, fill, fill)
            text(s, item, x + 0.55, y - 0.04, 3.25, 0.62, 12, INK, True, PP_ALIGN.LEFT)
    text(s, "最重要：实物书包 + 8 种物品 + 词卡。其余都能替代，不用临时网购绘本。",
         1.5, 6.65, 10.3, 0.4, 16, CORAL, True)
    page_num(s, 44)

    # 45 credits -----------------------------------------------------------------
    s = white_slide(prs)
    text(s, "素材来源 Credits", 0.5, 0.6, 12.3, 0.9, 44, NAVY, True, PP_ALIGN.LEFT)
    credits = [
        "音乐（已内嵌课件）：Kevin MacLeod（incompetech.com），CC BY 4.0 —— Fluffing a Duck / Sneaky Snitch / Monkeys Spinning Monkeys / Carefree",
        "原创故事：Pip's First Day —— 本课件原创角色与情节，可自由用于课堂教学",
        "课程依据：《Our World Starter × Phonics 1 · 24 课整合大纲》L1（Unit 1 My School）",
        "词汇与句型：Our World Starter Unit 1；字母音 /p/ /b/ 依据 Phonics 1（课堂内部使用）",
        "打印页（Slide 40–43）：词卡 Set A / Set B、p·b 声音篮卡、故事排序卡",
        "课件生成：scripts/build_starter_lesson1_deck.py（python-pptx），可重复构建",
    ]
    for i, line in enumerate(credits):
        y = 1.9 + i * 0.85
        circle(s, 0.8, y + 0.14, 0.18, GREEN, GREEN)
        text(s, line, 1.25, y, 11.4, 0.75, 18, INK, True, PP_ALIGN.LEFT)
    text(s, "Hello Song / Goodbye Song 建议用 Super Simple Songs 版本（流媒体搜索），因版权原因未嵌入。",
         1.25, 6.95, 11.4, 0.4, 14, MUTED, True, PP_ALIGN.LEFT)
    page_num(s, 45)

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(OUTPUT)
    print(f"Wrote {OUTPUT} ({len(prs.slides)} slides)")


if __name__ == "__main__":
    build()
