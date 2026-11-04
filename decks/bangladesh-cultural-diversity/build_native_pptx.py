#!/usr/bin/env python3
"""Build a fully EDITABLE native PPTX for the Bangladesh deck.

Every slide is composed of real PowerPoint objects (text boxes, shapes,
native charts, pictures) — no baked-in slide images. Speaker notes come
from speech.md.
"""
from __future__ import annotations

import argparse
import io
import re
from pathlib import Path

from PIL import Image, ImageOps
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_MARKER_STYLE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

IVORY = RGBColor(0xF7, 0xF3, 0xE8)
CARD = RGBColor(0xFD, 0xFB, 0xF4)
GREEN = RGBColor(0x00, 0x6A, 0x4E)
GREEN_DK = RGBColor(0x00, 0x4D, 0x38)
VERM = RGBColor(0xF4, 0x2A, 0x41)
INK = RGBColor(0x14, 0x14, 0x14)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREY = RGBColor(0x8A, 0x8A, 0x8A)
LGREY = RGBColor(0xD9, 0xD4, 0xC5)

HEAD_FONT = "Arial Black"
BODY = "Arial"

W, H = 13.333, 7.5


# ---------- helpers ----------
def set_spc(run, val=60):  # letter spacing (1/100 pt units)
    run.font._rPr.set("spc", str(val))


def txt(slide, x, y, w, h, text, size=14, color=INK, bold=False, font=BODY,
        align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spc=None, line=1.0, caps=False):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    lines = text.split("\n")
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line
        r = p.add_run()
        r.text = ln.upper() if caps else ln
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.name = font
        r.font.color.rgb = color
        if spc:
            set_spc(r, spc)
    return tb


def rect(slide, x, y, w, h, fill=None, line=None, lw=1.0, rounded=False, radius=0.08, shadow_off=True):
    shp = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h))
    if rounded:
        try:
            shp.adjustments[0] = radius
        except Exception:
            pass
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(lw)
    if shadow_off:
        try:
            shp.shadow.inherit = False
        except Exception:
            pass
    return shp


def shape_text(shp, text, size=13, color=INK, bold=False, font=BODY,
               align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, spc=None, line=1.0):
    tf = shp.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.06)
    tf.margin_top = tf.margin_bottom = Inches(0.03)
    for i, ln in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line
        r = p.add_run()
        r.text = ln
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.name = font
        r.font.color.rgb = color
        if spc:
            set_spc(r, spc)
    return shp


def badge(slide, n):
    b = rect(slide, W - 1.05, 0.12, 0.9, 0.42, fill=VERM, rounded=False)
    shape_text(b, f"{n:02d} / 17", size=12, color=WHITE, bold=True, font=BODY, spc=50)


def rules(slide, x, y, w, color=INK):
    rect(slide, x, y, w, 0.025, fill=color)
    rect(slide, x, y + 0.07, w, 0.015, fill=color)


def eyebrow(slide, x, y, text, color=GREEN):
    txt(slide, x, y, 6.5, 0.3, text, size=12, color=color, bold=True, spc=250, caps=True)


def motif(slide, corner="tr", color=GREEN):
    # simple jamdani-flavored diagonal line cluster
    import math
    if corner == "tr":
        bx, by = W - 1.7, 0.15
    elif corner == "tl":
        bx, by = 0.18, 0.15
    elif corner == "bl":
        bx, by = 0.18, H - 1.15
    else:
        bx, by = W - 1.7, H - 1.15
    for i in range(3):
        ln = rect(slide, bx + i * 0.16, by + i * 0.13, 1.05, 0.03, fill=color)
        ln.rotation = -45 if corner in ("tr", "bl") else 45


def motif_tr_safe(slide, color=GREEN):
    """Top-right corner motif that clears the badge (x<12.2) and headlines."""
    for i in range(3):
        ln = rect(slide, 10.95 + i * 0.16, 0.15 + i * 0.13, 1.05, 0.03, fill=color)
        ln.rotation = -45


def photo(slide, figs: Path, name, x, y, w, h, caption=None, frame=True):
    src = figs / name
    img = Image.open(src).convert("RGB")
    tw, th = int(w * 150), int(h * 150)
    img = ImageOps.fit(img, (tw, th), Image.LANCZOS, centering=(0.5, 0.45))
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=88)
    buf.seek(0)
    if frame:
        rect(slide, x - 0.05, y - 0.05, w + 0.10, h + 0.10 + (0.34 if caption else 0.05),
             fill=WHITE, line=INK, lw=0.75)
    slide.shapes.add_picture(buf, Inches(x), Inches(y), Inches(w), Inches(h))
    if caption:
        txt(slide, x, y + h + 0.07, w, 0.26, caption, size=9.5, color=INK, bold=True,
            align=PP_ALIGN.CENTER, spc=180, caps=True)


def headline(slide, text, y=0.42, size=30, w=None, x=0.55, color=INK, split_at=None):
    w = w or (W - x - 1.6)
    t = text
    if split_at and split_at in t:
        t = t.replace(split_at, split_at + "\n", 1)
    nlines = t.count("\n") + 1
    txt(slide, x, y, w, 0.6 * nlines, t, size=size, color=color, bold=True,
        font=HEAD_FONT, line=0.95, caps=True)
    rules(slide, x, y + 0.62 * nlines + 0.03, w)
    return y + 0.62 * nlines + 0.2


def chip(slide, x, y, w, h, text, fill=GREEN, color=WHITE, size=11.5, bold=True, rounded=True):
    c = rect(slide, x, y, w, h, fill=fill, rounded=rounded, radius=0.5 if rounded else 0)
    shape_text(c, text, size=size, color=color, bold=bold)
    return c


def card(slide, x, y, w, h, fill=CARD, line=INK, lw=1.0):
    return rect(slide, x, y, w, h, fill=fill, line=line, lw=lw, rounded=True, radius=0.06)


def load_notes(deck: Path):
    content = (deck / "speech.md").read_text(encoding="utf-8")
    notes, cur, buf = {}, None, []
    pat = re.compile(r"^#{2,4}\s*Slide\s*(\d+)\b.*$", re.I)
    for line in content.splitlines():
        m = pat.match(line.strip())
        if m:
            if cur:
                txt_ = "\n".join(buf).strip()
                if txt_:
                    notes[cur] = txt_
            cur, buf = int(m.group(1)), []
        elif cur:
            buf.append(line)
    if cur:
        txt_ = "\n".join(buf).strip()
        if txt_:
            notes[cur] = txt_
    return notes


# ---------- slides ----------
def jamdani_band(slide, y=None):
    y = (H - 0.45) if y is None else y
    rect(slide, 0, y, W, H - y, fill=GREEN)
    for i in range(int(W / 0.5) + 1):
        d = rect(slide, 0.2 + i * 0.5, y + 0.12, 0.18, 0.18, fill=IVORY)
        d.rotation = 45


def s01(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    eyebrow(s, 0.7, 0.5, "South Asia · Seminar Series")
    txt(s, 0.7, 0.95, 12.3, 1.75,
        "CULTURAL DIVERSITY AND\nNATIONAL IDENTITY IN BANGLADESH",
        size=40, color=INK, bold=True, font=HEAD_FONT, line=0.98)
    txt(s, 0.7, 2.8, 11, 0.5, "Historical and Political Perspectives", size=23, color=INK)
    rect(s, 0.72, 3.28, 4.5, 0.035, fill=VERM)
    txt(s, 0.7, 3.48, 11, 0.4, "Masters Seminar Presentation", size=15, color=INK, spc=40)
    motif(s, "tr", INK)
    rect(s, 0.0, 4.47, W, 0.02, fill=INK)
    return {"strip": ("lalbagh_fort_pano.jpg", 0.0, 4.5, W, 3.0)}


def s02(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    txt(s, 1.6, 0.5, 10.2, 1.5, "LANGUAGE MOVEMENT\nAND EKUSHEY FEBRUARY",
        size=38, color=GREEN, bold=True, font=HEAD_FONT, align=PP_ALIGN.CENTER, line=0.98)
    rect(s, 0.55, 2.12, W - 1.1, 0.03, fill=VERM)
    rect(s, 0.55, 2.19, W - 1.1, 0.015, fill=VERM)
    badge(s, 2)
    motif(s, "tl", GREEN)
    cards = [
        ("1952", "Students sacrifice their lives demanding Bangla as the state language"),
        ("1972", "Article 3: Bangla is the state language of the Republic"),
        ("1999", "UNESCO declares 21 February International Mother Language Day"),
    ]
    y = 2.55
    for yr, line in cards:
        card(s, 6.95, y, 5.85, 1.3, line=INK, lw=1.0)
        yc = rect(s, 7.2, y + 0.2, 1.05, 0.9, fill=GREEN, rounded=True, radius=0.18)
        shape_text(yc, yr, size=17, color=WHITE, bold=True, font=HEAD_FONT)
        txt(s, 8.5, y + 0.2, 4.1, 0.95, line, size=12.5, color=INK, line=1.05,
            anchor=MSO_ANCHOR.MIDDLE)
        y += 1.48
    return {
        "p1": ("shaheed_minar_monument.jpg", 0.7, 2.55, 2.85, 3.85, "Shaheed Minar"),
        "p2": ("shaheed_minar_ekushey.webp", 3.85, 2.55, 2.85, 3.85, "Ekushey observance"),
    }


def s03(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    headline(s, "From Partition to Liberation, 1947–1971", y=0.42, size=30, w=11.18)
    badge(s, 3)
    motif(s, "bl", GREEN)
    band_y = 5.5
    rect(s, 0.55, band_y, 12.25, 0.85, fill=GREEN)
    nodes = [
        (0.9, 1.65, 3.3, "1947 — Partition: East Bengal joins Pakistan; mass refugee flows"),
        (5.3, 6.05, 2.4, "1952 — Language movement"),
        (8.6, 9.35, 3.3, "1971 — Nine-month Liberation War; victory on 16 December"),
    ]
    for nx, tx, tw, t in nodes:
        d = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(nx + 0.4), Inches(band_y - 0.17),
                               Inches(0.34), Inches(0.34))
        d.fill.solid(); d.fill.fore_color.rgb = VERM
        d.line.fill.background(); d.shadow.inherit = False
        txt(s, tx, band_y + 0.15, tw, 0.62, t, size=11, color=WHITE, bold=True, line=1.0,
            anchor=MSO_ANCHOR.MIDDLE)
    txt(s, 5.35, band_y + 0.2, 0.5, 0.4, "→", size=16, color=WHITE, bold=True)
    txt(s, 8.55, band_y + 0.2, 0.5, 0.4, "→", size=16, color=WHITE, bold=True)
    txt(s, 0.55, 6.6, 12.2, 0.4,
        "Memory of 1971 anchors a civic, language-based national identity.",
        size=14, color=INK, bold=True, align=PP_ALIGN.CENTER)
    return {
        "p1": ("partition_1947_refugees.jpg", 0.6, 1.9, 5.9, 3.05, "Partition refugees, 1947"),
        "p2": ("libwar_refugee_sculpture.jpg", 6.9, 1.9, 5.9, 3.05, "Liberation War memorial"),
    }


def s04(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    headline(s, "Constitution, Principles, and the Secularism Debate", y=0.42,
             size=27, w=10.6, split_at=", Principles,")
    badge(s, 4)
    motif_tr_safe(s, GREEN)
    rect(s, 6.5, 1.9, 0.5, 4.65, fill=GREEN, rounded=True, radius=0.5)
    steps = [
        ("1972", GREEN, "1972", "Four principles: nationalism, socialism, democracy, secularism", 1.95),
        (None, GREEN_DK, "Fifth Amendment (1977/1979)", "Secularism removed; Bismillah inserted", 3.15),
        ("1988", VERM, "Eighth Amendment — Islam, the state religion",
         "2005 High Court: removal of secularism illegal", 4.35),
        ("2011", VERM, "Fifteenth Amendment — secularism restored",
         "Article 12 anti-communalism; Article 2A keeps Islam as state religion", 5.55),
    ]
    for pill_year, pill_color, head, body, y in steps:
        if pill_year:
            pill = rect(s, 6.48, y, 0.54, 0.95, fill=pill_color, rounded=True, radius=0.4)
            tb = txt(s, 6.48 - 0.33, y + 0.3, 1.2, 0.35, pill_year, size=11.5, color=WHITE,
                     bold=True, font=HEAD_FONT, align=PP_ALIGN.CENTER)
            tb.rotation = 270
        else:
            d = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(6.62), Inches(y + 0.32),
                                   Inches(0.26), Inches(0.26))
            d.fill.solid(); d.fill.fore_color.rgb = VERM
            d.line.color.rgb = IVORY; d.line.width = Pt(1.5); d.shadow.inherit = False
        txt(s, 7.25, y, 5.55, 0.4, head, size=12.5, color=pill_color, bold=True,
            font=HEAD_FONT, caps=True, line=1.0)
        txt(s, 7.25, y + 0.42, 5.55, 0.75, body, size=11.5, color=INK, line=1.0)
    chip(s, 6.5, 6.75, 3.1, 0.42, "Art. 8 · 2A · 12 · 41", fill=GREEN, size=11.5)
    return {"p1": ("martyrs_memorial_savar.jpg", 0.7, 1.9, 5.1, 4.55,
                   "National Martyrs’ Memorial, Savar")}


def s05(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    eyebrow(s, 0.7, 0.5, "Faith · Majority Tradition")
    headline(s, "The Islamic Heritage of Bengal", y=0.9, size=28, w=6.1, x=0.7,
             split_at=" Heritage")
    badge(s, 5)
    txt(s, 0.7, 2.55, 6.0, 1.25, "91.04%", size=70, color=VERM, bold=True, font=HEAD_FONT)
    txt(s, 0.7, 3.85, 6.0, 0.4, "Muslims of 169.8M — Census 2022", size=16, color=INK, bold=True)
    notes = [
        "Baitul Mukarram: the national mosque, Dhaka",
        "Sufi shrines (mazars) knit local devotion across faiths",
        "Baul–Sufi synthesis — UNESCO Intangible Heritage, 2005",
    ]
    y = 4.5
    for n in notes:
        rect(s, 0.72, y + 0.09, 0.15, 0.15, fill=GREEN)
        txt(s, 1.0, y, 5.6, 0.5, n, size=13, color=INK, line=1.0)
        y += 0.7
    jamdani_band(s, H - 0.45)
    return {"p1": ("mosque_baitul_mukarram.jpg", 7.15, 0.7, 5.6, 5.85,
                   "Baitul Mukarram, Dhaka")}


def s06(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    motif_tr_safe(s, GREEN)
    headline(s, "Four Threads of Identity", y=0.45, size=40, w=11.18)
    badge(s, 6)
    items = [
        ("I", "LANGUAGE", "Bangla state language + 50+ tongues of the peoples"),
        ("II", "FAITH", "Muslim-majority public life, multi-faith institutions"),
        ("III", "PEOPLES", "~54 indigenous groups, 35+ languages, hill and plain"),
        ("IV", "MEMORY", "1971 Liberation War as civic religion"),
    ]
    x = 0.55
    for num, ttl, body in items:
        card(s, x, 2.0, 3.0, 4.1, line=INK, lw=1.25)
        txt(s, x, 2.35, 3.0, 0.9, num, size=44, color=GREEN, bold=True, font=HEAD_FONT,
            align=PP_ALIGN.CENTER)
        rect(s, x + 0.5, 3.35, 2.0, 0.02, fill=INK)
        txt(s, x + 0.15, 3.55, 2.7, 0.55, ttl, size=20, color=INK, bold=True, font=HEAD_FONT,
            align=PP_ALIGN.CENTER)
        txt(s, x + 0.25, 4.3, 2.5, 1.6, body, size=13, color=INK, align=PP_ALIGN.CENTER,
            line=1.1)
        x += 3.16
    jamdani_band(s, H - 0.45)
    return {}


def s07(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    motif_tr_safe(s, GREEN)
    headline(s, "Who Is Indigenous? Counting and Naming the Peoples", y=0.42,
             size=25, w=8.4, split_at=" Counting")
    badge(s, 7)
    txt(s, 0.6, 1.78, 8.2, 0.35,
        "Census 2022 counts 1,650,159 indigenous persons across 50 groups",
        size=13.5, color=INK, bold=True)
    bars = [
        ("CHAKMA", "483,299", 1.0), ("MARMA", "224,261", 0.46),
        ("TRIPURI", "156,578", 0.32), ("SANTAL", "129,049", 0.265),
        ("ORAON", "85,846", 0.177), ("GARO", "76,846", 0.158),
    ]
    y = 2.3
    maxw = 4.7
    for name, fig, frac in bars:
        bw = max(maxw * frac, 1.35)
        rect(s, 0.6, y, bw, 0.42, fill=GREEN, rounded=True, radius=0.5)
        txt(s, 0.78, y + 0.06, 1.6, 0.3, name, size=12, color=WHITE, bold=True, spc=60)
        txt(s, 0.6 + bw + 0.15, y + 0.05, 1.5, 0.3, fig, size=13, color=INK, bold=True)
        y += 0.54
    rect(s, 0.6, 5.9, 7.9, 1.35, fill=None, line=VERM, lw=1.75, rounded=True, radius=0.05)
    tab = rect(s, 0.75, 5.75, 1.3, 0.4, fill=VERM, rounded=True, radius=0.3)
    shape_text(tab, "CAVEAT", size=12, color=WHITE, bold=True, spc=80)
    txt(s, 0.8, 6.25, 7.5, 0.95,
        "Definitions are contested: the government says “ethnic minority”; the UN and most "
        "movements say “indigenous”. Activist estimates exceed 3 million; the census count is far lower.",
        size=12, color=INK, line=1.05)
    return {"p1": ("indigenous_family_ch.jpg", 9.0, 1.78, 3.7, 4.1,
                   "Chakma girls in traditional dress")}


def s08(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    txt(s, 0.6, 0.45, W - 1.2, 0.8, "TEMPLES, VIHARAS, AND CHURCHES", size=40, color=INK,
        bold=True, font=HEAD_FONT, align=PP_ALIGN.CENTER, caps=True)
    rules(s, 1.6, 1.3, W - 3.2)
    txt(s, 0.6, 1.5, W - 1.2, 0.35, "Three living traditions, one constitutional promise",
        size=15, color=INK, align=PP_ALIGN.CENTER)
    badge(s, 8)
    items = [
        ("temple_dhakeshwari.webp", "Dhakeshwari Temple", "The state temple of Dhaka"),
        ("church_holy_cross.jpg", "Holy Cross Church",
         "Christianity since the Portuguese era; Holy Cross educational tradition"),
        ("buddhist_golden_temple.webp", "Buddha Dhatu Jadi", "Bandarban’s golden temple"),
    ]
    x = 0.75
    for fn, cap, line in items:
        photo(s, FIGS, fn, x, 2.1, 3.6, 2.6)
        card(s, x, 4.95, 3.6, 1.35, line=INK, lw=1.0)
        txt(s, x + 0.1, 5.08, 3.4, 0.35, cap, size=15, color=INK, bold=True, font=HEAD_FONT,
            align=PP_ALIGN.CENTER, caps=True)
        rect(s, x + 0.7, 5.48, 2.2, 0.02, fill=INK)
        txt(s, x + 0.2, 5.58, 3.2, 0.68, line, size=11, color=INK, align=PP_ALIGN.CENTER,
            line=1.0)
        x += 4.05
    rect(s, 0.55, 6.55, 12.25, 0.55, fill=GREEN)
    chip(s, 0.75, 6.63, 1.15, 0.4, "Art. 41", fill=VERM, size=11)
    txt(s, 2.15, 6.65, 10.5, 0.35, "freedom to profess, practice, and propagate religion",
        size=13.5, color=WHITE, bold=True)
    return {}


def s09(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    headline(s, "Census 2022: The Religious Mosaic in Numbers", size=26, w=11.18)
    badge(s, 9)
    # donut chart (native, editable data)
    data = CategoryChartData()
    data.categories = ["Muslim", "Hindu", "Buddhist", "Christian", "Other"]
    data.add_series("Shares", (91.04, 7.95, 0.61, 0.30, 0.12))
    gf = s.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(0.55), Inches(1.9),
                            Inches(5.4), Inches(3.0), data)
    chart = gf.chart
    chart.has_legend = False
    chart.has_title = False
    plot = chart.plots[0]
    ser = plot.series[0]
    cols = [GREEN, VERM, RGBColor(0xE0, 0x9B, 0x1D), RGBColor(0x3D, 0x6F, 0xB5), GREY]
    for i, c in enumerate(cols):
        pt = ser.points[i]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = c
        pt.format.line.color.rgb = IVORY
        pt.format.line.width = Pt(1.25)
    txt(s, 2.15, 3.12, 2.2, 0.6, "169.8M", size=22, color=INK, bold=True, font=HEAD_FONT, align=PP_ALIGN.CENTER)
    txt(s, 0.65, 1.95, 3.5, 0.3, "RELIGION SHARES", size=13, color=INK, bold=True, font=HEAD_FONT, spc=80)
    # legend list (exact values, single occurrence)
    legend = [
        ("Muslim", "91.04%", GREEN), ("Hindu", "7.95%", VERM),
        ("Buddhist", "0.61%", RGBColor(0xE0, 0x9B, 0x1D)),
        ("Christian", "0.30%", RGBColor(0x3D, 0x6F, 0xB5)),
        ("Other", "0.12%", GREY),
    ]
    ly = 5.15
    lx = 1.0
    for name, val, c in legend:
        dot = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(lx), Inches(ly + 0.04), Inches(0.16), Inches(0.16))
        dot.fill.solid(); dot.fill.fore_color.rgb = c; dot.line.fill.background(); dot.shadow.inherit = False
        txt(s, lx + 0.24, ly, 2.3, 0.3, f"{name}  {val}", size=12, color=INK, bold=True)
        ly += 0.27
    # trend chart (native, editable)
    d2 = CategoryChartData()
    d2.categories = ["1941", "1951", "1974", "2011", "2022"]
    d2.add_series("Hindu share", (28, 22, 13.5, 8.54, 7.95))
    gf2 = s.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Inches(6.4), Inches(1.9),
                             Inches(6.4), Inches(2.6), d2)
    ch2 = gf2.chart
    ch2.has_legend = False
    ch2.has_title = False
    ch2.font.size = Pt(11)
    ch2.font.name = BODY
    ser2 = ch2.plots[0].series[0]
    ser2.format.line.color.rgb = GREEN
    ser2.format.line.width = Pt(2.25)
    ser2.smooth = False
    ser2.marker.style = XL_MARKER_STYLE.CIRCLE
    ser2.marker.size = 8
    ser2.marker.format.fill.solid()
    ser2.marker.format.fill.fore_color.rgb = VERM
    ser2.marker.format.line.color.rgb = VERM
    plot2 = ch2.plots[0]
    plot2.has_data_labels = True
    dl = plot2.data_labels
    dl.number_format = '0.##"%"'
    dl.number_format_is_linked = False
    dl.font.size = Pt(11)
    dl.font.bold = True
    dl.font.color.rgb = INK
    try:
        ch2.value_axis.visible = False
        ch2.value_axis.has_major_gridlines = False
    except Exception:
        pass
    txt(s, 6.5, 1.92, 6.0, 0.3, "HINDU SHARE OF POPULATION", size=13, color=INK, bold=True, font=HEAD_FONT, spc=80)
    # photo grid 2x2
    grid = [
        ("mosque_baitul_mukarram.jpg", "Mosque"), ("temple_dhakeshwari.webp", "Temple"),
        ("church_holy_cross.jpg", "Church"), ("buddhist_golden_temple.webp", "Golden Temple"),
    ]
    gx, gy = 6.5, 4.65
    for i, (fn, cap) in enumerate(grid):
        x = gx + (i % 2) * 3.2
        y = gy + (i // 2) * 1.15
        photo(s, FIGS, fn, x, y, 2.6, 0.85, frame=False)
        txt(s, x + 2.7, y + 0.2, 0.5, 0.5, cap.upper(), size=10, color=INK, bold=True, spc=60, line=1.0)
    rect(s, 0.55, H - 0.55, 12.25, 0.55, fill=GREEN)
    txt(s, 0.8, H - 0.44, 12, 0.35, "ARTICLE 41 — FREEDOM OF RELIGION IN PRACTICE",
        size=13, color=WHITE, bold=True, align=PP_ALIGN.CENTER, spc=120)
    return {}


def s10(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    headline(s, "Ethnic Peoples and Their Arts", y=0.35, size=29, w=9.5)
    badge(s, 10)
    photo(s, FIGS, "wangala_garo_festival.jpg", 0.6, 1.5, 12.1, 3.0, "Wangala — Garo harvest festival")
    eyebrow(s, 0.6, 5.15, "Diversity in the hills and plains")
    chips = ["Chakma", "Marma", "Tripuri", "Santal", "Oraon", "Garo", "Rakhaine", "Tanchangya"]
    x = 0.6
    for i, c in enumerate(chips):
        w_ = 1.42
        chip(s, x, 5.5, w_, 0.44, c, fill=(GREEN if i % 2 == 0 else CARD),
             color=(WHITE if i % 2 == 0 else INK), size=12)
        x += w_ + 0.1
    txt(s, 0.6, 6.15, 5.6, 0.35, "~54 peoples · 35+ languages", size=13.5, color=INK, bold=True)
    txt(s, 0.6, 6.55, 12, 0.6,
        "Wangala — the Garo harvest festival of thanks · Weaving, drum, and dance traditions carried across generations",
        size=13, color=INK)
    motif_tr_safe(s, VERM)
    return {}


def s11(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    headline(s, "Chittagong Hill Tracts: Accord and After", size=27, w=10.2, split_at=" Tracts:")
    badge(s, 11)
    checks = [
        "Accord signed 2 December 1997 after two decades of conflict",
        "Ministry of CHT Affairs",
        "CHT Regional Council · 3 Hill District Councils · Land Commission",
        "Land and recognition still contested: “ethnic minority” or “indigenous”?",
    ]
    y = 1.8
    for c in checks:
        b = rect(s, 6.3, y, 0.42, 0.42, fill=GREEN, rounded=True, radius=0.25)
        shape_text(b, "✓", size=16, color=WHITE, bold=True)
        txt(s, 6.9, y + 0.03, 5.9, 0.7, c, size=13, color=INK, bold=True, line=1.0)
        y += 0.8
    rect(s, 6.3, 5.25, 6.5, 1.7, fill=None, line=VERM, lw=1.75, rounded=True, radius=0.05)
    tab = rect(s, 6.45, 5.1, 1.7, 0.4, fill=VERM, rounded=True, radius=0.3)
    shape_text(tab, "UNFINISHED", size=12, color=WHITE, bold=True, spc=60)
    txt(s, 6.55, 5.6, 6.05, 0.4, "27 years without Regional or District Council elections",
        size=13, color=INK, bold=True)
    rect(s, 6.55, 6.05, 6.0, 0.02, fill=VERM)
    txt(s, 6.55, 6.15, 6.05, 0.7,
        "Conflict toll ~8,140 killed, injured, or missing (Dey 2024) · Kroc peace-score 49/100",
        size=13, color=INK, bold=True, line=1.0)
    motif_tr_safe(s, GREEN)
    return {"p1": ("cht_rangamati_lake.jpg", 0.6, 1.8, 5.3, 4.9, "Kaptai Lake, Rangamati")}


def s12(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    headline(s, "The Rohingya Crisis", size=32, w=7)
    badge(s, 12)
    tiles = [
        ("1.1–1.2M", "Rohingya hosted in Bangladesh"),
        ("33 camps", "Cox’s Bazar + Bhasan Char"),
        ("~700,000", "fled Myanmar after 27 August 2017;\n~150,000 new arrivals since early 2024"),
        ("$710.5M", "UNHCR 2026 appeal — 26% below the previous year"),
    ]
    y = 1.75
    for fig, sub in tiles:
        card(s, 0.6, y, 5.5, 1.15)
        txt(s, 0.85, y + 0.12, 5.0, 0.5, fig, size=24, color=GREEN, bold=True, font=HEAD_FONT)
        txt(s, 0.85, y + 0.62, 5.05, 0.5, sub, size=11.5, color=INK, line=1.0)
        y += 1.3
    rect(s, 0.0, H - 0.62, W, 0.62, fill=VERM)
    txt(s, 0.4, H - 0.5, 12.6, 0.4,
        "2017 mass flight → 2024 new arrivals → 2025 deadliest sea year (~900 dead or missing) → "
        "2025 UNGA High-Level Conference (30 Sep)",
        size=12, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    motif_tr_safe(s, GREEN)
    return {"p1": ("rohingya_camp_aerial.jpg", 6.5, 1.75, 6.3, 4.4, "Kutupalong camp, Cox’s Bazar")}


def s13(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    headline(s, "Urban Minorities: Biharis and the Politics of Belonging", size=25, w=11.4, split_at=" Minorities:")
    badge(s, 13)
    blocks = [
        ("I", "300,000–400,000", "Urdu-speaking Bangladeshis across 116 camps since 1971"),
        ("II", "Geneva Camp", "Mohammadpur, Dhaka — the largest camp (~21,700 residents)"),
        ("III", "2008", "Supreme Court ruling: stateless no more — “Urdu-speaking Bangladeshis”"),
        ("IV", "CITIZENSHIP", "won in court; urban exclusion persists"),
    ]
    positions = [(0.6, 2.0), (6.9, 2.0), (0.6, 4.2), (6.9, 4.2)]
    for (num, fig, sub), (x, y) in zip(blocks, positions):
        card(s, x, y, 5.85, 1.95, line=INK, lw=1.25)
        txt(s, x + 0.25, y + 0.2, 1.0, 0.7, num, size=30, color=GREEN, bold=True, font=HEAD_FONT)
        txt(s, x + 1.3, y + 0.22, 4.4, 0.6, fig, size=22, color=INK, bold=True, font=HEAD_FONT, caps=True)
        txt(s, x + 1.3, y + 0.95, 4.4, 0.9, sub, size=12.5, color=INK, line=1.05)
    rect(s, 0.0, H - 0.7, W, 0.7, fill=VERM)
    txt(s, 0.4, H - 0.56, 12.6, 0.45, "BELONGING, EARNED LATE", size=20, color=WHITE, bold=True,
        font=HEAD_FONT, align=PP_ALIGN.CENTER, spc=200)
    motif_tr_safe(s, GREEN)
    return {}


def s14(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    headline(s, "Festivals and Intangible Heritage", y=0.35, size=28, w=9.8)
    badge(s, 14)
    photo(s, FIGS, "mangal_shobhajatra.jpg", 0.6, 1.45, 12.1, 2.5, "Mangal Shobhajatra — New Year procession")
    photo(s, FIGS, "jamdani_weaving.webp", 0.6, 4.5, 3.5, 2.3, "Jamdani weaving")
    unesco = [
        "Mangal Shobhajatra — UNESCO 2016, born 1989",
        "Baul songs — UNESCO 2005   ·   Jamdani — UNESCO 2008",
    ]
    y = 4.55
    for u in unesco:
        chip(s, 4.5, y, 8.2, 0.5, u, fill=GREEN, size=12.5, rounded=True)
        y += 0.62
    x = 4.5
    for c in ["Pahela Baishakh", "Eid cycle", "Wangala"]:
        w_ = 2.0 if c != "Pahela Baishakh" else 2.5
        chip(s, x, y, w_, 0.48, c, fill=CARD, color=INK, size=12)
        x += w_ + 0.15
    y += 0.6
    chip(s, 4.5, y, 5.6, 0.48, "DURGA PUJA — NATIONAL HOLIDAY", fill=VERM, size=12.5)
    txt(s, 4.5, y + 0.6, 8.2, 0.4, "Festivals are civic spaces where pluralism becomes national.",
        size=13, color=INK, bold=True)
    motif_tr_safe(s, VERM)
    return {}


def s15(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    txt(s, 0.55, 0.35, 11.4, 1.5,
        "2024–2026: UPHEAVAL, ELECTION,\nAND THE MINORITY QUESTION",
        size=32, color=INK, bold=True, font=HEAD_FONT, line=0.98)
    rules(s, 0.55, 1.7, W - 2.2)
    badge(s, 15)
    band_y = 3.9
    rect(s, 0.55, band_y, 12.25, 0.35, fill=GREEN, rounded=True, radius=0.5)
    cards_ = [
        ("AUG 2024", "Uprising topples the government; Yunus-led interim takes office"),
        ("12 FEB 2026", "BNP 209/297 seats (49.97%) · Jamaat 68 (31.76%) · turnout 59.44%"),
        ("17 FEB 2026", "Tarique Rahman sworn in as Prime Minister"),
        ("OCT 2026", "Minority rights “on citizenship, not religion”; security at ~33,000 Durga Puja venues"),
    ]
    x = 0.55
    for i, (head, body) in enumerate(cards_):
        cx = x + i * 3.12
        d = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(cx + 1.35), Inches(band_y - 0.02), Inches(0.4), Inches(0.4))
        d.fill.solid(); d.fill.fore_color.rgb = VERM; d.line.color.rgb = IVORY
        d.line.width = Pt(2); d.shadow.inherit = False
        card(s, cx, 2.1, 2.9, 1.6)
        txt(s, cx + 0.15, 2.25, 2.6, 0.35, head, size=14, color=INK, bold=True, font=HEAD_FONT)
        txt(s, cx + 0.15, 2.63, 2.6, 1.0, body, size=11, color=INK, line=1.0)
    rect(s, 2.4, 4.75, 8.5, 0.85, fill=None, line=VERM, lw=1.75, rounded=True, radius=0.3)
    txt(s, 2.6, 4.9, 8.1, 0.6,
        "ASK: 138 violent incidents against religious minorities, Jan–Nov 2024 (mostly 5–8 Aug) · rival counts contested",
        size=12.5, color=INK, bold=True, align=PP_ALIGN.CENTER, line=1.0)
    rect(s, 0, H - 0.4, W, 0.4, fill=GREEN)
    for i in range(26):
        d = rect(s, 0.25 + i * 0.5, H - 0.3, 0.14, 0.14, fill=IVORY)
        d.rotation = 45
    motif_tr_safe(s, GREEN)
    return {}


def s16(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    headline(s, "Conclusion: Identity as Negotiation", y=0.5, size=30, w=11.18)
    badge(s, 16)
    rows = [
        ("UNITY", "language and the memory of 1971 hold the nation together"),
        ("PLURALISM", "faiths, peoples, and languages make the nation plural"),
        ("CONSTITUTION", "secularism removed and restored: Art. 12 as promise"),
        ("BAROMETER", "how minorities fare tells us how democracy fares"),
    ]
    y = 2.3
    for lead, rest in rows:
        tb = s.shapes.add_textbox(Inches(0.7), Inches(y), Inches(12.0), Inches(0.7))
        tf = tb.text_frame; tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        r1 = p.add_run(); r1.text = lead + " — "
        r1.font.size = Pt(24); r1.font.bold = True; r1.font.name = HEAD_FONT
        r1.font.color.rgb = VERM
        r2 = p.add_run(); r2.text = rest
        r2.font.size = Pt(20); r2.font.bold = True; r2.font.name = BODY
        r2.font.color.rgb = INK
        y += 1.0
    rect(s, 0.7, 6.7, 12.0, 0.02, fill=GREEN)
    txt(s, 0.7, 6.85, 12, 0.3, "Cultural Diversity and National Identity in Bangladesh · Seminar",
        size=11, color=GREEN, bold=True, spc=200, caps=True, align=PP_ALIGN.CENTER)
    motif_tr_safe(s, GREEN)
    return {}


def s17(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    headline(s, "References and Sources", y=0.42, size=32, w=10)
    badge(s, 17)
    refs = [
        ("Demography", "Bangladesh Bureau of Statistics, Population Census 2022 (final)"),
        ("Rohingya", "UNHCR emergency updates & 2026 appeal; UNGA High-Level Conference on Myanmar (30 Sep 2025)"),
        ("Heritage", "UNESCO ICH: Mangal Shobhajatra (01091), Baul songs, Jamdani"),
        ("Law & politics", "The Daily Star; AJCL 69(3); Asian Survey (2017); ASK reports 2024–25; Feb 2026 gazette results"),
        ("CHT", "UNPO; ANU; Indigenous Navigator; Dey (2024); Kroc Institute"),
        ("Photos", "Wikimedia Commons; per-figure sources: assets/figures/SOURCES.md"),
    ]
    pos = [(0.6, 1.8), (6.9, 1.8), (0.6, 3.35), (6.9, 3.35), (0.6, 4.9), (6.9, 4.9)]
    for (cat, body), (x, y) in zip(refs, pos):
        card(s, x, y, 5.85, 1.35)
        txt(s, x + 0.2, y + 0.15, 5.4, 0.3, cat.upper(), size=12, color=GREEN, bold=True, spc=150)
        txt(s, x + 0.2, y + 0.48, 5.45, 0.8, body, size=12, color=INK, line=1.05)
    txt(s, 0.6, 6.55, 12.1, 0.35,
        "Cultural Diversity and National Identity in Bangladesh · Masters Seminar",
        size=12, color=INK, align=PP_ALIGN.CENTER, spc=60)
    motif_tr_safe(s, GREEN)
    return {}


def main():
    global FIGS
    ap = argparse.ArgumentParser()
    ap.add_argument("deck_dir")
    ap.add_argument("out")
    args = ap.parse_args()
    deck = Path(args.deck_dir)
    FIGS = deck / "assets" / "figures"

    prs = Presentation()
    prs.slide_width = Inches(W)
    prs.slide_height = Inches(H)

    builders = [s01, s02, s03, s04, s05, s06, s07, s08, s09, s10, s11, s12, s13, s14, s15, s16, s17]
    photo_specs = []
    for b in builders:
        spec = b(prs) or {}
        photo_specs.append(spec)

    # post-pass: insert photos (so they layer above backgrounds but captions stay readable)
    for slide, spec in zip(prs.slides, photo_specs):
        for key, val in spec.items():
            if key == "strip":
                fn, x, y, w, h = val
                photo(slide, FIGS, fn, x, y, w, h, frame=False)
            elif key == "p2b":
                fn, x, y, w, h, cap = val
                photo(slide, FIGS, fn, x, y, w, h, cap)
            else:
                fn, x, y, w, h, cap = val
                photo(slide, FIGS, fn, x, y, w, h, cap)

    # notes
    notes = load_notes(deck)
    n_set = 0
    for i, slide in enumerate(prs.slides, 1):
        if i in notes:
            slide.notes_slide.notes_text_frame.text = notes[i]
            n_set += 1

    prs.save(args.out)
    print(f"saved {args.out} | slides={len(prs.slides._sldIdLst)} | notes={n_set}/17")


if __name__ == "__main__":
    main()
