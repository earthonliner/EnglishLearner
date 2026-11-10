#!/usr/bin/env python3
"""Build Lesson 1: My School as an editable PPTX deck."""

from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "lesson-plans" / "our-world-starter-lesson1-my-school.pptx"

W = 13.333
H = 7.5

NAVY = "173F5F"
TEAL = "20BFA9"
TEAL_DARK = "129988"
YELLOW = "FFD166"
CORAL = "FF6B6B"
SKY = "7BDFF2"
PURPLE = "A78BFA"
CREAM = "FFF9ED"
WHITE = "FFFFFF"
INK = "23445D"
MUTED = "678093"
MINT = "E7FAF5"
PALE_YELLOW = "FFF5CC"
PALE_BLUE = "EAF7FC"
PALE_CORAL = "FFE9E6"
PALE_PURPLE = "F1EBFF"
BROWN = "9B6A46"
GREEN = "70C36B"
PINK = "F58DB2"
BLACK = "1F2937"
LIGHT_GRAY = "E9EEF2"


def rgb(hex_color):
    return RGBColor.from_string(hex_color)


def rect(slide, x, y, w, h, fill, line=None, radius=True, transparency=0):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    shape.fill.transparency = transparency
    shape.line.color.rgb = rgb(line or fill)
    if radius:
        try:
            shape.adjustments[0] = 0.12
        except (IndexError, ValueError):
            pass
    return shape


def circle(slide, x, y, d, fill, line=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d))
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    shape.line.color.rgb = rgb(line or fill)
    return shape


def line(slide, x1, y1, x2, y2, color=NAVY, width=2):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(x1),
        Inches(y1),
        Inches(max(0.02, x2 - x1)),
        Inches(max(0.02, y2 - y1)),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(color)
    shape.line.color.rgb = rgb(color)
    return shape


def text(
    slide,
    value,
    x,
    y,
    w,
    h,
    size=24,
    color=INK,
    bold=False,
    align=PP_ALIGN.LEFT,
    font="Aptos",
    valign=MSO_ANCHOR.MIDDLE,
    margin=0.05,
):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.margin_left = Inches(margin)
    frame.margin_right = Inches(margin)
    frame.margin_top = Inches(margin)
    frame.margin_bottom = Inches(margin)
    frame.vertical_anchor = valign
    p = frame.paragraphs[0]
    p.alignment = align
    p.space_after = Pt(0)
    run = p.add_run()
    run.text = value
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = rgb(color)
    return box


def rich_text(slide, runs, x, y, w, h, size=24, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.MIDDLE):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.margin_left = Inches(0.05)
    frame.margin_right = Inches(0.05)
    frame.margin_top = Inches(0.02)
    frame.margin_bottom = Inches(0.02)
    frame.vertical_anchor = valign
    p = frame.paragraphs[0]
    p.alignment = align
    for value, color, bold, run_size in runs:
        r = p.add_run()
        r.text = value
        r.font.name = "Aptos"
        r.font.size = Pt(run_size or size)
        r.font.bold = bold
        r.font.color.rgb = rgb(color)
    return box


def add_background(slide, color=CREAM):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(color)
    # Soft decorative bubbles.
    c1 = circle(slide, -0.35, -0.35, 1.3, YELLOW)
    c1.fill.transparency = 30
    c2 = circle(slide, 12.55, 6.85, 1.25, SKY)
    c2.fill.transparency = 35


def add_header(slide, section, title, subtitle=None, accent=TEAL, printable=False):
    if printable:
        text(slide, title, 0.45, 0.18, 9.4, 0.45, 22, NAVY, True)
        text(slide, "PRINT • CUT • LAMINATE (optional)", 9.55, 0.2, 3.25, 0.35, 9, MUTED, True, PP_ALIGN.RIGHT)
        line(slide, 0.45, 0.7, 12.85, 0.72, LIGHT_GRAY)
        return
    rect(slide, 0.52, 0.35, 1.42, 0.38, accent, accent)
    text(slide, section.upper(), 0.58, 0.36, 1.3, 0.31, 10, WHITE, True, PP_ALIGN.CENTER)
    text(slide, title, 0.55, 0.83, 11.7, 0.65, 28, NAVY, True)
    if subtitle:
        text(slide, subtitle, 0.58, 1.43, 11.8, 0.34, 11.5, MUTED)


def add_footer(slide, number, time_label=None, cue=None):
    if cue:
        rect(slide, 0.5, 6.86, 10.7, 0.34, PALE_BLUE, PALE_BLUE)
        text(slide, "TEACHER  " + cue, 0.67, 6.87, 10.35, 0.28, 8.5, MUTED, False)
    else:
        line(slide, 0.55, 6.96, 11.3, 6.98, LIGHT_GRAY)
    if time_label:
        rect(slide, 11.35, 6.82, 1.08, 0.42, YELLOW, YELLOW)
        text(slide, time_label, 11.38, 6.86, 1.02, 0.27, 9, NAVY, True, PP_ALIGN.CENTER)
    text(slide, f"{number:02d}", 12.53, 6.85, 0.38, 0.3, 9, MUTED, True, PP_ALIGN.RIGHT)


def new_slide(prs, section, title, subtitle=None, accent=TEAL, bg=CREAM, cue=None, time_label=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide, bg)
    add_header(slide, section, title, subtitle, accent)
    add_footer(slide, len(prs.slides), time_label, cue)
    return slide


def label_card(slide, x, y, w, h, title, color=WHITE, accent=TEAL, subtitle=None):
    rect(slide, x, y, w, h, color, color)
    rect(slide, x, y, 0.12, h, accent, accent, radius=False)
    text(slide, title, x + 0.25, y + 0.16, w - 0.42, 0.43, 18, NAVY, True)
    if subtitle:
        text(slide, subtitle, x + 0.25, y + 0.65, w - 0.42, h - 0.78, 11, MUTED)


def draw_book(slide, x, y, scale=1, color=SKY):
    rect(slide, x, y, 1.15 * scale, 0.82 * scale, color, NAVY, False)
    rect(slide, x + 0.53 * scale, y + 0.03 * scale, 0.05 * scale, 0.76 * scale, WHITE, WHITE, False)
    text(slide, "A", x + 0.1 * scale, y + 0.12 * scale, 0.38 * scale, 0.35 * scale, 16 * scale, WHITE, True, PP_ALIGN.CENTER)
    return 1.15 * scale, 0.82 * scale


def draw_pencil(slide, x, y, scale=1, color=YELLOW, rotation=-15):
    body = rect(slide, x, y, 1.25 * scale, 0.18 * scale, color, BROWN, False)
    body.rotation = rotation
    tip = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(x - 0.12 * scale), Inches(y - 0.02 * scale), Inches(0.27 * scale), Inches(0.22 * scale))
    tip.rotation = 270 + rotation
    tip.fill.solid()
    tip.fill.fore_color.rgb = rgb(CREAM)
    tip.line.color.rgb = rgb(BROWN)
    eraser = rect(slide, x + 1.17 * scale, y, 0.22 * scale, 0.18 * scale, PINK, BROWN, False)
    eraser.rotation = rotation


def draw_pen(slide, x, y, scale=1, color=TEAL):
    body = rect(slide, x, y, 0.18 * scale, 1.22 * scale, color, NAVY, False)
    rect(slide, x, y - 0.11 * scale, 0.18 * scale, 0.2 * scale, NAVY, NAVY, False)
    clip = rect(slide, x + 0.12 * scale, y + 0.12 * scale, 0.05 * scale, 0.45 * scale, NAVY, NAVY, False)
    body.rotation = 8
    clip.rotation = 8


def draw_crayon(slide, x, y, scale=1, color=CORAL):
    rect(slide, x, y, 0.35 * scale, 1.05 * scale, color, BROWN, False)
    tip = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(x), Inches(y - 0.23 * scale), Inches(0.35 * scale), Inches(0.3 * scale))
    tip.fill.solid()
    tip.fill.fore_color.rgb = rgb(color)
    tip.line.color.rgb = rgb(BROWN)
    rect(slide, x, y + 0.63 * scale, 0.35 * scale, 0.08 * scale, PALE_YELLOW, PALE_YELLOW, False)


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
    rect(slide, x + 0.23 * scale, y + 0.3 * scale, 0.85 * scale, 0.36 * scale, PALE_YELLOW, BROWN, False)


def draw_chair(slide, x, y, scale=1):
    rect(slide, x + 0.15 * scale, y, 0.8 * scale, 0.72 * scale, SKY, NAVY, True)
    rect(slide, x, y + 0.72 * scale, 1.1 * scale, 0.22 * scale, SKY, NAVY, False)
    rect(slide, x + 0.12 * scale, y + 0.9 * scale, 0.13 * scale, 0.75 * scale, NAVY, NAVY, False)
    rect(slide, x + 0.82 * scale, y + 0.9 * scale, 0.13 * scale, 0.75 * scale, NAVY, NAVY, False)


def draw_bag(slide, x, y, scale=1, color=TEAL):
    rect(slide, x, y + 0.35 * scale, 2.05 * scale, 1.85 * scale, color, NAVY, True)
    handle = slide.shapes.add_shape(MSO_SHAPE.ARC, Inches(x + 0.55 * scale), Inches(y), Inches(0.95 * scale), Inches(0.85 * scale))
    handle.fill.background()
    handle.line.color.rgb = rgb(NAVY)
    handle.line.width = Pt(4 * scale)
    rect(slide, x + 0.35 * scale, y + 1.3 * scale, 1.35 * scale, 0.58 * scale, SKY, NAVY, True)
    circle(slide, x + 0.92 * scale, y + 1.5 * scale, 0.18 * scale, YELLOW, NAVY)


def draw_object(slide, kind, x, y, scale=1, color=None):
    if kind == "book":
        draw_book(slide, x, y, scale, color or SKY)
    elif kind == "pencil":
        draw_pencil(slide, x, y + 0.35 * scale, scale, color or YELLOW)
    elif kind == "pen":
        draw_pen(slide, x + 0.48 * scale, y, scale, color or TEAL)
    elif kind == "crayon":
        draw_crayon(slide, x + 0.4 * scale, y, scale, color or CORAL)
    elif kind == "paper":
        draw_paper(slide, x + 0.12 * scale, y, scale)
    elif kind == "eraser":
        draw_eraser(slide, x + 0.08 * scale, y + 0.35 * scale, scale)
    elif kind == "desk":
        draw_desk(slide, x, y + 0.12 * scale, scale)
    elif kind == "chair":
        draw_chair(slide, x + 0.1 * scale, y, scale * 0.82)


def object_card(slide, kind, label, x, y, w=2.6, h=2.05, bg=WHITE, show_label=True, scale=1):
    rect(slide, x, y, w, h, bg, bg)
    draw_object(slide, kind, x + w / 2 - 0.62 * scale, y + 0.22, scale)
    if show_label:
        rect(slide, x + 0.18, y + h - 0.55, w - 0.36, 0.38, NAVY, NAVY)
        text(slide, label, x + 0.22, y + h - 0.53, w - 0.44, 0.28, 16, WHITE, True, PP_ALIGN.CENTER)


def speech_bubble(slide, value, x, y, w, h, fill=WHITE, border=TEAL, size=22):
    rect(slide, x, y, w, h, fill, border, True)
    text(slide, value, x + 0.18, y + 0.1, w - 0.36, h - 0.2, size, NAVY, True, PP_ALIGN.CENTER)


def draw_face(slide, x, y, scale=1, animal="puppy", mood="happy", shirt=TEAL):
    # Body and head.
    rect(slide, x + 0.35 * scale, y + 1.25 * scale, 1.2 * scale, 1.25 * scale, shirt, NAVY, True)
    circle(slide, x + 0.28 * scale, y + 0.28 * scale, 1.35 * scale, PALE_YELLOW, NAVY)
    if animal == "puppy":
        ear1 = slide.shapes.add_shape(MSO_SHAPE.TEAR, Inches(x), Inches(y + 0.25 * scale), Inches(0.55 * scale), Inches(0.9 * scale))
        ear2 = slide.shapes.add_shape(MSO_SHAPE.TEAR, Inches(x + 1.35 * scale), Inches(y + 0.25 * scale), Inches(0.55 * scale), Inches(0.9 * scale))
        for e, rotation in ((ear1, 20), (ear2, 200)):
            e.rotation = rotation
            e.fill.solid()
            e.fill.fore_color.rgb = rgb(BROWN)
            e.line.color.rgb = rgb(NAVY)
    else:
        for dx in (0.28, 1.15):
            ear = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE, Inches(x + dx * scale), Inches(y), Inches(0.5 * scale), Inches(0.58 * scale))
            ear.fill.solid()
            ear.fill.fore_color.rgb = rgb(PINK)
            ear.line.color.rgb = rgb(NAVY)
    circle(slide, x + 0.62 * scale, y + 0.7 * scale, 0.11 * scale, NAVY)
    circle(slide, x + 1.18 * scale, y + 0.7 * scale, 0.11 * scale, NAVY)
    circle(slide, x + 0.88 * scale, y + 0.92 * scale, 0.15 * scale, CORAL)
    if mood == "sad":
        line(slide, x + 0.79 * scale, y + 1.18 * scale, x + 1.12 * scale, y + 1.21 * scale, NAVY)
        circle(slide, x + 0.52 * scale, y + 0.86 * scale, 0.1 * scale, SKY)
    else:
        mouth = slide.shapes.add_shape(MSO_SHAPE.ARC, Inches(x + 0.72 * scale), Inches(y + 0.94 * scale), Inches(0.5 * scale), Inches(0.38 * scale))
        mouth.fill.background()
        mouth.line.color.rgb = rgb(NAVY)
        mouth.line.width = Pt(2)


def draw_big_letter(slide, letter, x, y, d, fill, subtitle=None):
    circle(slide, x, y, d, fill, NAVY)
    text(slide, letter, x, y + 0.03 * d, d, 0.7 * d, int(70 * d / 2), NAVY, True, PP_ALIGN.CENTER)
    if subtitle:
        text(slide, subtitle, x - 0.1, y + 0.7 * d, d + 0.2, 0.25 * d, int(14 * d / 2), NAVY, True, PP_ALIGN.CENTER)


def add_vocab_grid(slide, items, y=2.0, show_labels=True):
    for i, (kind, label, bg) in enumerate(items):
        x = 0.65 + i * 3.05
        object_card(slide, kind, label, x, y, 2.65, 2.25, bg, show_labels, 1.05)


def add_print_card(slide, kind, label, x, y, color):
    rect(slide, x, y, 5.9, 2.75, WHITE, NAVY, False)
    rect(slide, x, y, 0.18, 2.75, color, color, False)
    draw_object(slide, kind, x + 0.65, y + 0.45, 1.5)
    text(slide, label, x + 2.55, y + 0.78, 2.9, 0.8, 30, NAVY, True, PP_ALIGN.CENTER)
    text(slide, "What is it?  It’s a ______.", x + 2.35, y + 1.65, 3.25, 0.42, 12, MUTED, False, PP_ALIGN.CENTER)


def build():
    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)
    prs.core_properties.title = "Lesson 1: My School Adventure"
    prs.core_properties.subject = "Our World Starter × Phonics 1"
    prs.core_properties.author = "EnglishLearner"
    prs.core_properties.keywords = "My School, Starter, phonics, p, short a, CVC preparation"

    # 1 Cover
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background(slide, CREAM)
    rect(slide, 0.62, 0.52, 2.25, 0.45, TEAL, TEAL)
    text(slide, "OUR WORLD STARTER · L1", 0.72, 0.58, 2.05, 0.25, 10, WHITE, True, PP_ALIGN.CENTER)
    text(slide, "My School\nAdventure!", 0.75, 1.35, 6.2, 2.0, 38, NAVY, True)
    text(slide, "Words • Speaking • Sound Detectives • Story", 0.82, 3.55, 6.0, 0.5, 18, TEAL_DARK, True)
    draw_bag(slide, 8.55, 1.3, 1.55, TEAL)
    draw_pencil(slide, 7.9, 1.45, 1.25)
    draw_book(slide, 10.35, 3.95, 1.3)
    draw_crayon(slide, 8.2, 4.65, 1.05, CORAL)
    speech_bubble(slide, "What’s in my bag?", 7.0, 5.2, 4.6, 0.72, WHITE, CORAL, 21)
    add_footer(slide, 1, "90 min", "Keep the real school bag hidden until Slide 4.")

    # 2 Outcomes
    slide = new_slide(prs, "WELCOME", "Today I can…", "Three things to take home after Lesson 1", CORAL, cue="Point to each badge; children repeat the short promise.", time_label="0–3′")
    outcomes = [
        ("8", "school words", SKY, "I can name it."),
        ("Q+A", "one full sentence", YELLOW, "What is it? It’s a…"),
        ("p · a", "two sounds", MINT, "I can hear it."),
    ]
    for i, (big, title_value, bg, small) in enumerate(outcomes):
        x = 0.75 + i * 4.15
        rect(slide, x, 2.05, 3.6, 3.6, bg, bg)
        text(slide, big, x + 0.25, 2.35, 3.1, 1.0, 36, NAVY, True, PP_ALIGN.CENTER)
        text(slide, title_value, x + 0.3, 3.52, 3.0, 0.5, 19, NAVY, True, PP_ALIGN.CENTER)
        text(slide, small, x + 0.35, 4.35, 2.9, 0.55, 13, MUTED, False, PP_ALIGN.CENTER)
        rect(slide, x + 1.2, 5.12, 1.2, 0.16, CORAL if i == 1 else TEAL, CORAL if i == 1 else TEAL, False)

    # 3 Hello routine
    slide = new_slide(prs, "ROUTINE", "Hello, school explorers!", "Listen • Move • Freeze", TEAL, cue="Model quickly: Stand up / Sit down / Eyes on me / Hands on knees.", time_label="3–5′")
    actions = [
        ("1", "Stand up!", TEAL),
        ("2", "Sit down!", SKY),
        ("3", "Eyes on me!", YELLOW),
        ("4", "Hands on knees!", CORAL),
    ]
    for i, (n, phrase, color) in enumerate(actions):
        x = 0.72 + (i % 2) * 6.0
        y = 2.0 + (i // 2) * 2.05
        rect(slide, x, y, 5.4, 1.45, WHITE, WHITE)
        circle(slide, x + 0.25, y + 0.25, 0.9, color, color)
        text(slide, n, x + 0.25, y + 0.31, 0.9, 0.55, 22, NAVY, True, PP_ALIGN.CENTER)
        text(slide, phrase, x + 1.45, y + 0.35, 3.5, 0.55, 24, NAVY, True)

    # 4 Mystery bag
    slide = new_slide(prs, "EXPLORE", "The Mystery School Bag", "Shake it. Listen. Guess!", CORAL, cue="Use a real opaque bag. Let 3 children feel one object without looking.", time_label="5–8′")
    draw_bag(slide, 1.3, 2.0, 1.7, TEAL)
    text(slide, "?", 3.9, 1.85, 1.2, 1.4, 72, CORAL, True, PP_ALIGN.CENTER)
    speech_bubble(slide, "What’s in my bag?", 6.0, 2.0, 5.2, 0.95, WHITE, TEAL, 28)
    rect(slide, 6.25, 3.45, 4.7, 1.6, PALE_YELLOW, PALE_YELLOW)
    text(slide, "Touch → Think → Whisper", 6.5, 3.72, 4.2, 0.45, 22, NAVY, True, PP_ALIGN.CENTER)
    text(slide, "No peeking!", 6.55, 4.34, 4.1, 0.35, 14, CORAL, True, PP_ALIGN.CENTER)

    # 5 first four words
    slide = new_slide(prs, "WORDS 1", "Look • Say • Tap", "Four things inside the school bag", SKY, cue="Reveal one real object at a time. Say naturally, then clap syllables once.", time_label="8–13′")
    add_vocab_grid(slide, [
        ("book", "book", PALE_BLUE),
        ("pencil", "pencil", PALE_YELLOW),
        ("pen", "pen", MINT),
        ("crayon", "crayon", PALE_CORAL),
    ])
    text(slide, "Teacher: normal voice → robot voice → tiny voice", 1.3, 5.05, 10.7, 0.55, 17, NAVY, True, PP_ALIGN.CENTER)

    # 6 Show me
    slide = new_slide(prs, "GAME 1", "Show me… FAST!", "Listen first. Hold it up. Say it once.", YELLOW, cue="Children use their own objects. Say 8–10 commands in mixed order.", time_label="13–17′")
    for i, (kind, label, color) in enumerate([
        ("book", "BOOK!", SKY), ("pencil", "PENCIL!", YELLOW), ("pen", "PEN!", TEAL), ("crayon", "CRAYON!", CORAL)
    ]):
        x = 0.8 + i * 3.1
        circle(slide, x + 0.45, 2.0, 1.5, PALE_BLUE if i % 2 == 0 else PALE_YELLOW, color)
        draw_object(slide, kind, x + 0.62, 2.32, 0.75)
        text(slide, label, x, 4.05, 2.4, 0.6, 19, NAVY, True, PP_ALIGN.CENTER)
    rect(slide, 3.25, 5.05, 6.85, 0.65, NAVY, NAVY)
    text(slide, "Ready?  3 — 2 — 1 — SHOW ME!", 3.35, 5.13, 6.65, 0.42, 21, WHITE, True, PP_ALIGN.CENTER)

    # 7 second four words
    slide = new_slide(prs, "WORDS 2", "Look • Say • Do", "Four more things in our classroom", PURPLE, cue="Connect each word to an action: hold / rub / touch / sit.", time_label="17–22′")
    add_vocab_grid(slide, [
        ("paper", "paper", PALE_BLUE),
        ("eraser", "eraser", PALE_CORAL),
        ("desk", "desk", PALE_YELLOW),
        ("chair", "chair", MINT),
    ])
    text(slide, "paper — hold it   •   eraser — rub   •   desk — touch   •   chair — sit", 0.9, 5.05, 11.6, 0.55, 16.5, NAVY, True, PP_ALIGN.CENTER)

    # 8 action map
    slide = new_slide(prs, "MOVE", "Listen and do!", "Don’t repeat until the action is finished.", TEAL, cue="Give one-step commands, then combine two: Pick up a pencil and touch the desk.", time_label="22–25′")
    actions = [
        ("OPEN", "your book", "book", SKY),
        ("PICK UP", "a pencil", "pencil", YELLOW),
        ("TOUCH", "the desk", "desk", CORAL),
        ("SIT ON", "the chair", "chair", TEAL),
    ]
    for i, (verb, obj, kind, color) in enumerate(actions):
        x = 0.7 + (i % 2) * 6.1
        y = 1.95 + (i // 2) * 2.2
        rect(slide, x, y, 5.5, 1.55, WHITE, WHITE)
        circle(slide, x + 0.25, y + 0.23, 1.05, color, color)
        draw_object(slide, kind, x + 0.47, y + 0.42, 0.62)
        text(slide, verb, x + 1.65, y + 0.28, 1.55, 0.38, 17, CORAL, True)
        text(slide, obj, x + 1.65, y + 0.75, 3.25, 0.45, 22, NAVY, True)

    # 9 Find and touch
    slide = new_slide(prs, "GAME 2", "Find it. Touch it. Freeze!", "Walk — don’t run.", CORAL, cue="Place 8 large cards around the room. Call a word; children touch the matching card.", time_label="25–28′")
    words = [("book", SKY), ("pencil", YELLOW), ("pen", TEAL), ("crayon", CORAL), ("paper", PURPLE), ("eraser", PINK), ("desk", BROWN), ("chair", GREEN)]
    for i, (value, color) in enumerate(words):
        x = 0.72 + (i % 4) * 3.05
        y = 2.0 + (i // 4) * 1.55
        rect(slide, x, y, 2.55, 1.05, color, color)
        text(slide, value, x + 0.1, y + 0.25, 2.35, 0.45, 20, NAVY if color in (YELLOW, SKY, PINK, GREEN) else WHITE, True, PP_ALIGN.CENTER)
    speech_bubble(slide, "Touch the… PENCIL!", 3.7, 5.2, 5.7, 0.68, WHITE, NAVY, 22)

    # 10 memory tray
    slide = new_slide(prs, "GAME 3", "Look carefully…", "You have 10 seconds!", PURPLE, cue="Show this slide for 10 seconds, then advance. Children close eyes while you count down.", time_label="28–31′")
    items = [("book", "book", PALE_BLUE), ("pencil", "pencil", PALE_YELLOW), ("pen", "pen", MINT), ("crayon", "crayon", PALE_CORAL),
             ("paper", "paper", PALE_PURPLE), ("eraser", "eraser", PALE_CORAL), ("desk", "desk", PALE_YELLOW), ("chair", "chair", MINT)]
    for i, (kind, label, bg) in enumerate(items):
        x = 0.72 + (i % 4) * 3.05
        y = 1.9 + (i // 4) * 2.35
        object_card(slide, kind, label, x, y, 2.55, 1.95, bg, True, 0.85)
    text(slide, "10… 9… 8…", 4.8, 6.05, 3.7, 0.45, 22, CORAL, True, PP_ALIGN.CENTER)

    # 11 What's missing
    slide = new_slide(prs, "GAME 3", "What’s missing?", "Say the missing words!", PURPLE, cue="Missing: pencil and eraser. Accept the word first; challenge children to say two words.", time_label="28–31′")
    remaining = [("book", "book", PALE_BLUE), ("pen", "pen", MINT), ("crayon", "crayon", PALE_CORAL),
                 ("paper", "paper", PALE_PURPLE), ("desk", "desk", PALE_YELLOW), ("chair", "chair", MINT)]
    positions = [(0.72, 1.9), (3.77, 1.9), (6.82, 1.9), (9.87, 1.9), (2.25, 4.25), (7.35, 4.25)]
    for (kind, label, bg), (x, y) in zip(remaining, positions):
        object_card(slide, kind, label, x, y, 2.55, 1.95, bg, True, 0.85)
    text(slide, "___  +  ___", 4.6, 5.95, 4.1, 0.55, 27, CORAL, True, PP_ALIGN.CENTER)

    # 12 sentence model
    slide = new_slide(prs, "SPEAK", "One question. One full answer.", "Use the whole sentence — not just one word.", TEAL, cue="Model with a puppet twice. Children point to each color chunk as they repeat.", time_label="31–34′")
    draw_face(slide, 0.8, 2.25, 1.05, "puppy", "happy", TEAL)
    draw_face(slide, 10.55, 2.25, 1.05, "cat", "happy", CORAL)
    speech_bubble(slide, "What is it?", 2.55, 2.0, 3.25, 0.85, PALE_BLUE, TEAL, 25)
    speech_bubble(slide, "It’s a pencil.", 7.05, 3.2, 3.25, 0.85, PALE_YELLOW, CORAL, 25)
    draw_pencil(slide, 5.85, 4.5, 1.45)
    rich_text(slide, [
        ("What is it?", TEAL_DARK, True, 25),
        ("     →     ", MUTED, False, 22),
        ("It’s a ", NAVY, True, 25),
        ("pencil.", CORAL, True, 25),
    ], 2.6, 5.45, 8.2, 0.6, 25, PP_ALIGN.CENTER)

    # 13 voice switch
    slide = new_slide(prs, "PRACTICE", "Ask it four ways!", "Same sentence. New voice.", SKY, cue="Teacher points to one card; whole class asks. Puppet answers incorrectly once for children to correct.", time_label="34–38′")
    voices = [
        ("GIANT", "BIG voice", CORAL),
        ("MOUSE", "tiny voice", SKY),
        ("ROBOT", "ro-bot voice", YELLOW),
        ("SECRET", "whisper", PURPLE),
    ]
    for i, (name, desc, color) in enumerate(voices):
        x = 0.75 + i * 3.05
        rect(slide, x, 2.15, 2.55, 2.55, color, color)
        text(slide, name, x + 0.15, 2.55, 2.25, 0.55, 23, NAVY, True, PP_ALIGN.CENTER)
        text(slide, desc, x + 0.2, 3.35, 2.15, 0.4, 15, NAVY, False, PP_ALIGN.CENTER)
        text(slide, "?", x + 0.83, 4.0, 0.9, 0.55, 32, WHITE, True, PP_ALIGN.CENTER)
    speech_bubble(slide, "What is it?  —  It’s a…", 3.65, 5.35, 6.0, 0.65, WHITE, NAVY, 20)

    # 14 pass the bag
    slide = new_slide(prs, "PAIR GAME", "Pass the bag!", "Music on: pass. Music off: ask + answer.", CORAL, cue="Put 6–8 objects in the bag. A asks; B draws and answers; switch roles.", time_label="38–42′")
    draw_bag(slide, 1.1, 2.15, 1.45, TEAL)
    steps = [
        ("1", "PASS", "music on"),
        ("2", "STOP", "music off"),
        ("3", "ASK", "What is it?"),
        ("4", "ANSWER", "It’s a ___."),
    ]
    for i, (n, action, detail) in enumerate(steps):
        x = 4.6 + (i % 2) * 3.7
        y = 2.0 + (i // 2) * 1.8
        circle(slide, x, y + 0.12, 0.72, YELLOW if i % 2 == 0 else SKY, NAVY)
        text(slide, n, x, y + 0.25, 0.72, 0.3, 16, NAVY, True, PP_ALIGN.CENTER)
        text(slide, action, x + 0.95, y, 2.2, 0.35, 18, CORAL, True)
        text(slide, detail, x + 0.95, y + 0.48, 2.35, 0.38, 15, NAVY, True)

    # 15 optional stretch
    slide = new_slide(prs, "CHALLENGE", "Add one color!", "Only if the class is ready.", YELLOW, cue="Optional differentiation: show red/blue objects. Keep the core answer acceptable without color.", time_label="42–44′")
    rect(slide, 1.0, 2.2, 5.1, 2.9, PALE_CORAL, PALE_CORAL)
    draw_pen(slide, 2.0, 2.75, 1.4, CORAL)
    text(slide, "It’s a red pen.", 3.0, 3.15, 2.55, 0.55, 24, NAVY, True, PP_ALIGN.CENTER)
    rect(slide, 7.2, 2.2, 5.1, 2.9, PALE_BLUE, PALE_BLUE)
    draw_book(slide, 8.0, 2.9, 1.3, SKY)
    text(slide, "It’s a blue book.", 9.55, 3.15, 2.35, 0.55, 24, NAVY, True, PP_ALIGN.CENTER)
    text(slide, "CORE ✓  It’s a pen.          CHALLENGE ★  It’s a red pen.", 2.35, 5.6, 8.7, 0.45, 17, NAVY, True, PP_ALIGN.CENTER)

    # 16 checkpoint
    slide = new_slide(prs, "CHECKPOINT", "Pack my school bag!", "Listen carefully. Pick the right object.", TEAL, cue="Call 4 children: Put a book / pencil / pen / crayon in the bag. Class confirms the word.", time_label="44–45′")
    draw_bag(slide, 5.25, 2.25, 1.45, TEAL)
    draw_book(slide, 1.4, 2.55, 1.1)
    draw_pencil(slide, 2.85, 3.35, 1.1)
    draw_pen(slide, 9.45, 2.35, 1.1)
    draw_crayon(slide, 11.15, 2.45, 1.05)
    speech_bubble(slide, "Put a pencil in the bag.", 3.4, 5.25, 6.6, 0.72, WHITE, CORAL, 23)

    # 17 movement break
    slide = new_slide(prs, "RESET", "60-second School Freeze!", "Move when the music plays. Freeze on the word.", PURPLE, cue="Use any upbeat 60-second clip. Freeze and mime: book / pencil / desk / chair.", time_label="45′")
    text(slide, "MOVE", 1.0, 2.3, 3.0, 1.0, 48, CORAL, True, PP_ALIGN.CENTER)
    text(slide, "→", 4.25, 2.35, 1.0, 0.8, 44, TEAL, True, PP_ALIGN.CENTER)
    text(slide, "FREEZE", 5.5, 2.3, 3.2, 1.0, 48, NAVY, True, PP_ALIGN.CENTER)
    text(slide, "→", 8.85, 2.35, 1.0, 0.8, 44, TEAL, True, PP_ALIGN.CENTER)
    text(slide, "MIME", 9.95, 2.3, 2.3, 1.0, 48, PURPLE, True, PP_ALIGN.CENTER)
    rect(slide, 2.0, 4.25, 9.3, 0.9, YELLOW, YELLOW)
    text(slide, "BOOK!   PENCIL!   DESK!   CHAIR!", 2.1, 4.42, 9.1, 0.45, 26, NAVY, True, PP_ALIGN.CENTER)

    # 18 /p/ detective
    slide = new_slide(prs, "PHONICS", "Sound Detective: /p/", "Not the letter name “pee” — a quick, quiet /p/.", CORAL, cue="Hold tissue before lips. Say /p/ without adding ‘uh’. Children feel the puff.", time_label="45–49′")
    draw_big_letter(slide, "p", 1.0, 2.15, 2.5, YELLOW, "sound /p/")
    rect(slide, 4.4, 2.2, 7.1, 2.7, WHITE, WHITE)
    text(slide, "Lips together", 4.8, 2.58, 2.0, 0.45, 21, NAVY, True)
    text(slide, "→", 6.7, 2.58, 0.7, 0.45, 24, CORAL, True, PP_ALIGN.CENTER)
    text(slide, "tiny puff of air", 7.4, 2.58, 2.7, 0.45, 21, NAVY, True)
    rect(slide, 4.8, 3.45, 5.9, 0.62, PALE_BLUE, PALE_BLUE)
    text(slide, "/p/  /p/  /p/", 5.0, 3.56, 5.5, 0.36, 26, CORAL, True, PP_ALIGN.CENTER)
    text(slide, "Try it with a tissue.", 5.1, 4.32, 5.3, 0.36, 16, MUTED, True, PP_ALIGN.CENTER)

    # 19 p sort
    slide = new_slide(prs, "HEAR IT", "Does it start with /p/?", "Thumbs up or hands on head.", SKY, cue="Say words without showing labels first. Reveal pictures as confirmation.", time_label="49–53′")
    p_words = [("pen", "pen"), ("pencil", "pencil"), ("paper", "paper")]
    no_words = [("book", "book"), ("desk", "desk"), ("chair", "chair")]
    rect(slide, 0.75, 1.95, 5.8, 4.2, MINT, MINT)
    rect(slide, 6.78, 1.95, 5.8, 4.2, PALE_CORAL, PALE_CORAL)
    text(slide, "YES  /p/", 1.1, 2.15, 5.1, 0.45, 22, TEAL_DARK, True, PP_ALIGN.CENTER)
    text(slide, "NO  /p/", 7.15, 2.15, 5.1, 0.45, 22, CORAL, True, PP_ALIGN.CENTER)
    for i, (kind, label) in enumerate(p_words):
        object_card(slide, kind, label, 1.0 + i * 1.78, 3.0, 1.55, 2.15, WHITE, True, 0.7)
    for i, (kind, label) in enumerate(no_words):
        object_card(slide, kind, label, 7.03 + i * 1.78, 3.0, 1.55, 2.15, WHITE, True, 0.7)

    # 20 sound hunt
    slide = new_slide(prs, "HUNT", "Find three /p/ things!", "Point. Say. Place a sticky note.", TEAL, cue="Pairs search the room or picture cards for pen, pencil, paper. 45 seconds only.", time_label="53–55′")
    text(slide, "/p/", 0.95, 2.1, 2.8, 1.5, 72, CORAL, True, PP_ALIGN.CENTER)
    for i in range(3):
        rect(slide, 4.15 + i * 2.55, 2.25, 1.85, 1.85, YELLOW, YELLOW)
        text(slide, "?", 4.15 + i * 2.55, 2.55, 1.85, 0.8, 46, NAVY, True, PP_ALIGN.CENTER)
    text(slide, "pen  •  pencil  •  paper", 3.8, 4.75, 7.8, 0.65, 28, NAVY, True, PP_ALIGN.CENTER)
    rect(slide, 3.5, 5.65, 8.4, 0.42, PALE_BLUE, PALE_BLUE)
    text(slide, "Challenge: Can you think of one more /p/ word?", 3.65, 5.71, 8.1, 0.27, 13, MUTED, True, PP_ALIGN.CENTER)

    # 21 short a
    slide = new_slide(prs, "PHONICS", "Meet short /a/", "Open your mouth wide: /a/ /a/ apple.", YELLOW, cue="Keep it short /æ/, not the letter name ‘A’. Use mirror or partner mouth check.", time_label="55–60′")
    draw_big_letter(slide, "a", 1.0, 2.15, 2.5, SKY, "short /a/")
    examples = [
        ("APPLE", "round mouth", CORAL),
        ("ANT", "tiny steps", TEAL),
        ("ALLIGATOR", "big snap", GREEN),
    ]
    for i, (word, action, color) in enumerate(examples):
        x = 4.15 + i * 2.7
        rect(slide, x, 2.15, 2.25, 2.75, WHITE, WHITE)
        circle(slide, x + 0.62, 2.55, 1.0, color, color)
        text(slide, "a", x + 0.62, 2.72, 1.0, 0.5, 29, WHITE, True, PP_ALIGN.CENTER)
        text(slide, word, x + 0.18, 3.75, 1.9, 0.38, 16, NAVY, True, PP_ALIGN.CENTER)
        text(slide, action, x + 0.18, 4.25, 1.9, 0.28, 11, MUTED, False, PP_ALIGN.CENTER)
    text(slide, "Quick. Short. Smile a little.", 4.5, 5.45, 6.85, 0.5, 21, CORAL, True, PP_ALIGN.CENTER)

    # 22 p or a movement
    slide = new_slide(prs, "GAME 4", "/p/ or /a/?", "Hear the sound. Do the move.", CORAL, cue="Say one sound only, then use mixed words: pen, apple, paper, ant. Do not show the letter first.", time_label="60–63′")
    rect(slide, 0.95, 2.0, 5.25, 3.65, PALE_YELLOW, PALE_YELLOW)
    draw_big_letter(slide, "p", 1.5, 2.55, 1.55, YELLOW)
    text(slide, "PAT YOUR KNEES", 3.35, 2.7, 2.35, 0.9, 22, NAVY, True, PP_ALIGN.CENTER)
    text(slide, "/p/ /p/ /p/", 3.25, 4.05, 2.55, 0.55, 24, CORAL, True, PP_ALIGN.CENTER)
    rect(slide, 7.1, 2.0, 5.25, 3.65, PALE_BLUE, PALE_BLUE)
    draw_big_letter(slide, "a", 7.65, 2.55, 1.55, SKY)
    text(slide, "OPEN YOUR ARMS", 9.5, 2.7, 2.35, 0.9, 22, NAVY, True, PP_ALIGN.CENTER)
    text(slide, "/a/ /a/ /a/", 9.4, 4.05, 2.55, 0.55, 24, TEAL_DARK, True, PP_ALIGN.CENTER)

    # 23 sound slide
    slide = new_slide(prs, "BLEND PREP", "Build a Sound Road", "Touch each sound. Then slide.", PURPLE, cue="This is oral blending preparation, not a word to memorize. Keep sounds clean and separate first.", time_label="63–67′")
    draw_big_letter(slide, "p", 1.2, 2.25, 2.2, YELLOW, "/p/")
    rect(slide, 3.75, 3.15, 4.0, 0.18, TEAL, TEAL, False)
    for i in range(4):
        circle(slide, 3.75 + i * 1.25, 2.96, 0.55, TEAL, TEAL)
    text(slide, "slide the sound →", 4.3, 2.35, 3.0, 0.42, 17, TEAL_DARK, True, PP_ALIGN.CENTER)
    draw_big_letter(slide, "a", 8.4, 2.25, 2.2, SKY, "/a/")
    rich_text(slide, [
        ("/p/  …  /a/", NAVY, True, 28),
        ("     then     ", MUTED, False, 18),
        ("/p-a/", CORAL, True, 30),
    ], 3.3, 5.15, 6.8, 0.65, 28, PP_ALIGN.CENTER)

    # 24 floor road
    slide = new_slide(prs, "GAME 5", "Walk the Sound Road!", "Step • Say • Slide", TEAL, cue="Put printed p and a cards on floor 1 metre apart. One child steps; class voices the sound.", time_label="67–70′")
    steps = [
        ("1", "STEP", "p", YELLOW),
        ("2", "SAY", "/p/", PALE_YELLOW),
        ("3", "STEP", "a", SKY),
        ("4", "SLIDE", "/p-a/", MINT),
    ]
    for i, (n, action, sound, color) in enumerate(steps):
        x = 0.75 + i * 3.05
        rect(slide, x, 2.2, 2.55, 3.15, color, color)
        circle(slide, x + 0.85, 2.5, 0.85, WHITE, NAVY)
        text(slide, n, x + 0.85, 2.7, 0.85, 0.3, 17, NAVY, True, PP_ALIGN.CENTER)
        text(slide, action, x + 0.2, 3.65, 2.15, 0.35, 17, CORAL, True, PP_ALIGN.CENTER)
        text(slide, sound, x + 0.2, 4.32, 2.15, 0.55, 28, NAVY, True, PP_ALIGN.CENTER)

    # 25 story cover
    slide = new_slide(prs, "STORY LAB", "Pip’s First Day", "An original mini-story: predict before you read.", CORAL, cue="Use this as the fallback if the physical Spot book is unavailable. Point to Pip and ask: How does Pip feel?", time_label="70–72′")
    draw_face(slide, 1.25, 2.2, 1.35, "puppy", "happy", TEAL)
    draw_bag(slide, 4.0, 2.75, 1.15, CORAL)
    speech_bubble(slide, "What will Pip pack?", 7.0, 2.2, 4.4, 0.9, WHITE, TEAL, 26)
    choices = [("book", 7.35, 3.65), ("pencil", 8.85, 4.15), ("crayon", 10.65, 3.7)]
    for kind, x, y in choices:
        draw_object(slide, kind, x, y, 0.9)
    text(slide, "Point and predict!", 7.3, 5.45, 4.1, 0.45, 20, CORAL, True, PP_ALIGN.CENTER)

    # 26 story beginning
    slide = new_slide(prs, "STORY 1", "Pip packs for school.", "Beginning: Pip is ready!", SKY, cue="Read once; children echo only the bold object words. Ask: What is in the bag?", time_label="72–76′")
    draw_face(slide, 0.8, 2.05, 1.05, "puppy", "happy", TEAL)
    draw_bag(slide, 3.1, 2.45, 1.05, CORAL)
    for kind, x, y in [("book", 6.2, 2.25), ("pencil", 7.9, 3.0), ("pen", 9.35, 2.15), ("paper", 10.8, 2.25)]:
        draw_object(slide, kind, x, y, 0.85)
    rect(slide, 1.25, 5.0, 10.9, 0.85, WHITE, WHITE)
    rich_text(slide, [
        ("A ", NAVY, False, 22), ("book", TEAL_DARK, True, 24), (", a ", NAVY, False, 22),
        ("pencil", CORAL, True, 24), (", a ", NAVY, False, 22), ("pen", TEAL_DARK, True, 24),
        (", and ", NAVY, False, 22), ("paper", PURPLE, True, 24), (".", NAVY, False, 22),
    ], 1.45, 5.12, 10.5, 0.55, 22, PP_ALIGN.CENTER)

    # 27 story problem
    slide = new_slide(prs, "STORY 2", "Oh no! No crayon.", "Middle: Pip wants to draw.", CORAL, cue="Pause before reading. Let children search the picture and say crayon. Show Pip’s feeling with face/body.", time_label="76–79′")
    draw_face(slide, 1.0, 2.0, 1.3, "puppy", "sad", TEAL)
    draw_paper(slide, 5.05, 2.65, 1.25)
    rect(slide, 7.25, 2.4, 4.25, 2.3, PALE_CORAL, PALE_CORAL)
    text(slide, "NO", 7.6, 2.8, 1.3, 0.65, 34, CORAL, True, PP_ALIGN.CENTER)
    draw_crayon(slide, 9.3, 2.9, 1.15, CORAL)
    line(slide, 9.0, 2.75, 10.45, 4.1, CORAL, 5)
    speech_bubble(slide, "I need a crayon.", 2.5, 5.15, 4.3, 0.68, WHITE, NAVY, 22)
    speech_bubble(slide, "How does Pip feel?", 7.15, 5.15, 4.3, 0.68, WHITE, CORAL, 22)

    # 28 story ending
    slide = new_slide(prs, "STORY 3", "Mia shares a crayon.", "Ending: Friends help friends.", TEAL, cue="Teacher: Here you are. Pip: Thank you! Children act it out in pairs with one crayon.", time_label="79–82′")
    draw_face(slide, 0.85, 2.0, 1.15, "puppy", "happy", TEAL)
    draw_face(slide, 9.9, 2.0, 1.15, "cat", "happy", CORAL)
    draw_crayon(slide, 6.25, 3.1, 1.2, CORAL)
    speech_bubble(slide, "Here you are!", 7.25, 1.95, 2.4, 0.72, WHITE, CORAL, 20)
    speech_bubble(slide, "Thank you!", 3.0, 2.1, 2.4, 0.72, WHITE, TEAL, 20)
    rect(slide, 3.2, 5.15, 7.0, 0.65, PALE_YELLOW, PALE_YELLOW)
    text(slide, "A crayon + a kind friend = a happy first day!", 3.35, 5.27, 6.7, 0.38, 18, NAVY, True, PP_ALIGN.CENTER)

    # 29 sequence
    slide = new_slide(prs, "STORY GAME", "First — Then — Finally", "Stand by the right picture.", PURPLE, cue="Print Slide 36 or point on screen. Three children hold scene cards; class puts them in order.", time_label="82–85′")
    cards = [
        ("FIRST", "Pip packs.", TEAL),
        ("THEN", "No crayon!", CORAL),
        ("FINALLY", "Mia shares.", YELLOW),
    ]
    for i, (marker, event, color) in enumerate(cards):
        x = 0.8 + i * 4.15
        rect(slide, x, 2.05, 3.6, 3.65, WHITE, WHITE)
        rect(slide, x, 2.05, 3.6, 0.58, color, color, False)
        text(slide, marker, x + 0.15, 2.16, 3.3, 0.3, 15, NAVY, True, PP_ALIGN.CENTER)
        if i == 0:
            draw_bag(slide, x + 1.05, 2.95, 0.72, TEAL)
        elif i == 1:
            draw_crayon(slide, x + 1.55, 3.0, 0.85, CORAL)
            line(slide, x + 1.2, 2.85, x + 2.2, 4.15, CORAL, 4)
        else:
            draw_crayon(slide, x + 1.55, 3.0, 0.85, CORAL)
            circle(slide, x + 0.55, 3.2, 0.65, PALE_YELLOW, NAVY)
            circle(slide, x + 2.35, 3.2, 0.65, PALE_YELLOW, NAVY)
        text(slide, event, x + 0.25, 4.75, 3.1, 0.5, 20, NAVY, True, PP_ALIGN.CENTER)

    # 30 story response
    slide = new_slide(prs, "RESPOND", "Story detectives", "Point, answer, act.", SKY, cue="Ask three levels. Every child answers at least one: point → word → short sentence.", time_label="85–87′")
    questions = [
        ("POINT", "Where is the crayon?", TEAL),
        ("SAY", "What is it?", YELLOW),
        ("THINK", "Who helps Pip?", CORAL),
        ("ACT", "Here you are! / Thank you!", PURPLE),
    ]
    for i, (level, question, color) in enumerate(questions):
        x = 0.75 + (i % 2) * 6.05
        y = 1.95 + (i // 2) * 2.1
        rect(slide, x, y, 5.45, 1.55, WHITE, WHITE)
        rect(slide, x, y, 1.15, 1.55, color, color, False)
        text(slide, level, x + 0.08, y + 0.55, 0.98, 0.3, 12, NAVY, True, PP_ALIGN.CENTER)
        text(slide, question, x + 1.45, y + 0.35, 3.65, 0.75, 20, NAVY, True, PP_ALIGN.CENTER)

    # 31 exit ticket
    slide = new_slide(prs, "EXIT TICKET", "Choose one door!", "Everyone leaves with one success.", CORAL, cue="Children choose one task; teacher records ✓ independent / ○ with prompt. Aim for 30–40 seconds each.", time_label="87–89′")
    doors = [
        ("DOOR 1", "Say 3 words", SKY, "book • pen • chair"),
        ("DOOR 2", "Ask + answer", YELLOW, "What is it? It’s a…"),
        ("DOOR 3", "Hear /p/ or /a/", MINT, "point to the sound"),
    ]
    for i, (door, task, color, hint) in enumerate(doors):
        x = 0.78 + i * 4.15
        rect(slide, x, 2.0, 3.55, 3.75, color, NAVY, False)
        circle(slide, x + 2.85, 3.65, 0.25, NAVY, NAVY)
        text(slide, door, x + 0.25, 2.35, 3.05, 0.38, 15, CORAL, True, PP_ALIGN.CENTER)
        text(slide, task, x + 0.38, 3.15, 2.8, 0.85, 23, NAVY, True, PP_ALIGN.CENTER)
        text(slide, hint, x + 0.38, 4.62, 2.8, 0.45, 12.5, NAVY, False, PP_ALIGN.CENTER)

    # 32 home task
    slide = new_slide(prs, "TAKE HOME", "20-second School Bag Challenge", "One small task — no whole-book recording.", TEAL, cue="Model the exact recording before dismissal. Parents only need to hold the phone.", time_label="89–90′")
    draw_bag(slide, 0.9, 2.25, 1.3, TEAL)
    steps = [
        ("1", "Pick 3 things."),
        ("2", "Show each thing."),
        ("3", "Say: It’s a ___."),
    ]
    for i, (n, phrase) in enumerate(steps):
        y = 2.05 + i * 1.12
        circle(slide, 4.25, y, 0.72, YELLOW if i == 0 else SKY if i == 1 else CORAL, NAVY)
        text(slide, n, 4.25, y + 0.18, 0.72, 0.28, 16, NAVY, True, PP_ALIGN.CENTER)
        text(slide, phrase, 5.25, y + 0.05, 4.9, 0.55, 23, NAVY, True)
    rect(slide, 3.95, 5.6, 7.1, 0.55, PALE_YELLOW, PALE_YELLOW)
    text(slide, "Example: It’s a book. It’s a pencil. It’s a pen.", 4.1, 5.7, 6.8, 0.32, 15.5, NAVY, True, PP_ALIGN.CENTER)

    # 33 printable cards 1
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(WHITE)
    add_header(slide, "PRINTABLE", "Printable Vocabulary Cards · Set A", printable=True)
    add_print_card(slide, "book", "book", 0.5, 0.95, SKY)
    add_print_card(slide, "pencil", "pencil", 6.92, 0.95, YELLOW)
    add_print_card(slide, "pen", "pen", 0.5, 4.0, TEAL)
    add_print_card(slide, "crayon", "crayon", 6.92, 4.0, CORAL)

    # 34 printable cards 2
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(WHITE)
    add_header(slide, "PRINTABLE", "Printable Vocabulary Cards · Set B", printable=True)
    add_print_card(slide, "paper", "paper", 0.5, 0.95, PURPLE)
    add_print_card(slide, "eraser", "eraser", 6.92, 0.95, PINK)
    add_print_card(slide, "desk", "desk", 0.5, 4.0, BROWN)
    add_print_card(slide, "chair", "chair", 6.92, 4.0, GREEN)

    # 35 sound cards
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(WHITE)
    add_header(slide, "PRINTABLE", "Floor Sound Cards", printable=True)
    rect(slide, 0.7, 1.05, 5.65, 5.8, PALE_YELLOW, NAVY, False)
    text(slide, "p", 0.7, 1.55, 5.65, 3.3, 128, NAVY, True, PP_ALIGN.CENTER)
    text(slide, "/p/  quick puff", 1.0, 5.4, 5.05, 0.55, 22, CORAL, True, PP_ALIGN.CENTER)
    rect(slide, 6.98, 1.05, 5.65, 5.8, PALE_BLUE, NAVY, False)
    text(slide, "a", 6.98, 1.55, 5.65, 3.3, 128, NAVY, True, PP_ALIGN.CENTER)
    text(slide, "short /a/  apple", 7.28, 5.4, 5.05, 0.55, 22, TEAL_DARK, True, PP_ALIGN.CENTER)

    # 36 story sequence print
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(WHITE)
    add_header(slide, "PRINTABLE", "Pip’s First Day · Sequence Cards", printable=True)
    story_cards = [
        ("1", "Pip packs.", TEAL),
        ("2", "No crayon!", CORAL),
        ("3", "Mia shares.", YELLOW),
    ]
    for i, (n, caption, color) in enumerate(story_cards):
        x = 0.55 + i * 4.25
        rect(slide, x, 1.1, 3.85, 5.9, WHITE, NAVY, False)
        rect(slide, x, 1.1, 3.85, 0.68, color, color, False)
        text(slide, n, x + 1.45, 1.19, 0.95, 0.43, 20, NAVY, True, PP_ALIGN.CENTER)
        if i == 0:
            draw_face(slide, x + 0.45, 2.05, 0.8, "puppy", "happy", TEAL)
            draw_bag(slide, x + 1.85, 3.0, 0.65, CORAL)
        elif i == 1:
            draw_face(slide, x + 0.8, 2.05, 0.95, "puppy", "sad", TEAL)
            draw_crayon(slide, x + 2.55, 3.05, 0.75, CORAL)
            line(slide, x + 2.35, 2.9, x + 3.2, 4.0, CORAL, 4)
        else:
            draw_face(slide, x + 0.15, 2.1, 0.72, "puppy", "happy", TEAL)
            draw_face(slide, x + 2.15, 2.1, 0.72, "cat", "happy", CORAL)
            draw_crayon(slide, x + 1.72, 3.45, 0.65, CORAL)
        text(slide, caption, x + 0.4, 5.65, 3.05, 0.6, 24, NAVY, True, PP_ALIGN.CENTER)

    # 37 final prep checklist
    slide = new_slide(prs, "TEACHER PREP", "明天上课：购买与准备清单", "按优先级准备；没有绘本也可直接用本课原创故事替代", CORAL, bg=CREAM, cue=None, time_label=None)
    columns = [
        ("① 外卖 / 附近文具店", [
            "不透明书包或布袋 1 个",
            "书、铅笔、笔、蜡笔、纸、橡皮各 2–4",
            "便利贴 / 美纹纸 / 奖励贴",
            "纸巾 1 包（/p/ 吹气实验）",
        ], PALE_YELLOW, YELLOW),
        ("② 家中可找", [
            "小玩偶 1–2 个（代替 Pip / Mia）",
            "蓝牙音箱或手机音乐",
            "两个篮子 / 纸盒、计时器",
            "桌椅；红、蓝物品作挑战",
        ], PALE_BLUE, SKY),
        ("③ 今晚打印", [
            "词卡：Slides 33–34",
            "p / a 地板卡：Slide 35",
            "故事排序卡：Slide 36",
            "建议彩印；也可黑白印后手涂",
        ], MINT, TEAL),
    ]
    for i, (heading, items, bg, accent) in enumerate(columns):
        x = 0.55 + i * 4.25
        rect(slide, x, 1.9, 3.95, 4.45, bg, bg)
        rect(slide, x, 1.9, 3.95, 0.62, accent, accent, False)
        text(slide, heading, x + 0.18, 2.05, 3.58, 0.3, 14, NAVY, True, PP_ALIGN.CENTER)
        for j, item in enumerate(items):
            circle(slide, x + 0.25, 2.85 + j * 0.78, 0.22, accent, accent)
            text(slide, item, x + 0.62, 2.72 + j * 0.78, 3.0, 0.54, 12.3, NAVY, False)
    rect(slide, 1.55, 6.53, 10.2, 0.42, PALE_CORAL, PALE_CORAL)
    text(slide, "最重要：实物书包 + 8 种课堂物品 + 词卡；其他都可替代，不必临时网购绘本。", 1.7, 6.6, 9.9, 0.26, 13, CORAL, True, PP_ALIGN.CENTER)

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(OUTPUT)
    print(f"Wrote {OUTPUT} ({len(prs.slides)} slides)")


if __name__ == "__main__":
    build()
