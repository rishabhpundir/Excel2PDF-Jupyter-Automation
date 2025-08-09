#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel -> PDF (one page per row), simple "Key: Value" list.
- Page size / orientation is copied from the FIRST PAGE of template.pdf
- NO command-line font argument: script auto-picks an Arabic-capable font
- Draws keys with a Latin font and values with mixed-script support:
    * Arabic segments are shaped (arabic-reshaper + python-bidi) and drawn with an Arabic font
    * Non‑Arabic segments are drawn with a Latin font
    * Segments are concatenated on the same line with correct widths
- Very simple wrapping to page width (left-aligned lines for simplicity)

Run:
    python main2.py --excel "sample rows.xlsx" --template template.pdf --out new.pdf
"""

import argparse
import os
from pathlib import Path
from typing import Optional, List, Tuple

import pandas as pd
import fitz  # PyMuPDF (to read template page size)

# ReportLab for writing a new PDF
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import black

# Arabic helpers
try:
    import arabic_reshaper
    from bidi.algorithm import get_display
except Exception:
    arabic_reshaper = None
    get_display = None
    
# ONLY columns that are required
template_keys = [
    "region", "city", "hay", "pca",
    "Population",
    "Total Premisis (Census)",
    "Remaining Premisis",
    "Female %",
    "HH Avg Size",
    "Population Density",
    "ss ARPU",
    "ss Usage",
    "Income Level (Based on ARPU)",
    "Last deployment year",
    "Major City?",
    "Visual Remark",
    "NPV",
    "IRR",
    "Forecasted Uptake",
    "Actual Uptake %",
    "Coverage",
    "Wrong Classification?",
    "# of Data centers",
    "Mobile Towers",
    "Fiberized Towers",
    "Remaining Towers",
    "ISP Cost",
    "OSP Cost",
    "STC Category",
    "Mobily Category",
    "Dawiyat Category",
    "TLS Category",
]

# ---------------- Arabic detection / shaping ----------------
def is_arabic_char(ch: str) -> bool:
    return (
        ("\u0600" <= ch <= "\u06FF") or
        ("\u0750" <= ch <= "\u077F") or
        ("\u08A0" <= ch <= "\u08FF") or
        ("\uFB50" <= ch <= "\uFDFF") or
        ("\uFE70" <= ch <= "\uFEFF")
    )

def is_arabic_text(s: str) -> bool:
    if not isinstance(s, str):
        return False
    return any(is_arabic_char(ch) for ch in s)

def shape_arabic(s: str) -> str:
    if not isinstance(s, str) or not s:
        return s
    if arabic_reshaper and get_display and is_arabic_text(s):
        return get_display(arabic_reshaper.reshape(s))
    return s

# ---------------- Fonts ----------------
LATIN_FONT = "Helvetica"           # built-in (good Latin)
AR_FONT_NAME = "CustomArabicFont"  # we will register an Arabic-capable font here

FONT_SEARCH_PATHS = [
    # current folder first (drop a copy of the ttf/ttc here for best reliability)
    ".",
    # macOS
    "/System/Library/Fonts",
    "/System/Library/Fonts/Supplemental",
    "/Library/Fonts",
    os.path.expanduser("~/Library/Fonts"),
    # Linux
    "/usr/share/fonts/truetype",
    "/usr/share/fonts",
]
# Prefer fonts that include BOTH Arabic and Latin glyphs.
AR_CANDIDATES = [
    "NotoNaskhArabic-Regular.ttf",
    "NotoKufiArabic-Regular.ttf",
    "Amiri-Regular.ttf",
    "GeezaPro.ttc",     # macOS
    "AlBayan.ttc",      # macOS
]

def find_arabic_font() -> Optional[Path]:
    for base in FONT_SEARCH_PATHS:
        try:
            for cand in AR_CANDIDATES:
                p = Path(base) / cand if base != "." else Path(cand)
                if p.exists():
                    return p
        except Exception:
            continue
    return None

def register_arabic_font() -> str:
    """
    Try to find and register an Arabic-capable font.
    Returns the font name to use for Arabic runs. Falls back to LATIN_FONT if none found.
    """
    font_path = find_arabic_font()
    if font_path:
        try:
            pdfmetrics.registerFont(TTFont(AR_FONT_NAME, str(font_path)))
            return AR_FONT_NAME
        except Exception:
            pass
    # Fallback (won't shape Arabic correctly if glyphs missing)
    return LATIN_FONT

# ---------------- Layout helpers ----------------
def text_width(text: str, font_name: str, font_size: float) -> float:
    return pdfmetrics.stringWidth(text, font_name, font_size)

def segment_by_script(s: str) -> List[Tuple[bool, str]]:
    """
    Split string into runs of (is_arabic, substring).
    """
    if not s:
        return []
    runs: List[Tuple[bool, str]] = []
    cur_is_ar = is_arabic_char(s[0])
    cur = [s[0]]
    for ch in s[1:]:
        ia = is_arabic_char(ch)
        if ia == cur_is_ar:
            cur.append(ch)
        else:
            runs.append((cur_is_ar, "".join(cur)))
            cur = [ch]
            cur_is_ar = ia
    runs.append((cur_is_ar, "".join(cur)))
    return runs

def draw_mixed_line(
    cnv: rl_canvas.Canvas,
    x_left: float,
    y_baseline: float,
    max_width: float,
    latin_font: str,
    ar_font: str,
    kv_text: str,
    kv_size: float,
    leading: float
) -> float:
    """
    Draw a possibly mixed-script line with simple wrapping:
      - Segment into Arabic / non-Arabic runs.
      - Shape Arabic runs.
      - Place runs left-to-right, switching fonts per run.
    Returns the next y-baseline.
    """
    # Wrap manually: build lines that fit width
    lines: List[List[Tuple[str, str]]] = []  # list of lines; each line is list of (font_name, segment_text)
    current_line: List[Tuple[str, str]] = []
    current_width = 0.0

    # Split by spaces to allow wrapping at spaces, but keep script segmentation inside tokens
    words = kv_text.split(" ")
    for wi, w in enumerate(words):
        w_runs = segment_by_script(w)
        # shape / set font per run and compute total word width (including a space before it if needed)
        word_segments: List[Tuple[str, str]] = []
        word_width = 0.0

        # space prefix if not first word of line
        space_text = (" " if (current_line or wi > 0) else "")
        if space_text:
            word_segments.append((latin_font, space_text))
            word_width += text_width(space_text, latin_font, kv_size)

        for is_ar, seg in w_runs:
            seg2 = shape_arabic(seg) if is_ar else seg
            font = ar_font if is_ar else latin_font
            word_segments.append((font, seg2))
            word_width += text_width(seg2, font, kv_size)

        # If this word doesn't fit, break line (unless it's the first on the line)
        if current_line and current_width + word_width > max_width:
            lines.append(current_line)
            current_line = []
            current_width = 0.0
            # Add the word (without a leading space on a new line)
            # remove the space segment if present at start
            if word_segments and word_segments[0][1] == " ":
                word_segments = word_segments[1:]
                word_width = sum(text_width(seg, fnt, kv_size) for fnt, seg in word_segments)
        current_line.extend(word_segments)
        current_width += word_width

    if current_line:
        lines.append(current_line)

    # Draw lines
    y = y_baseline
    for line in lines:
        x = x_left
        for font, seg in line:
            cnv.setFont(font, kv_size)
            cnv.drawString(x, y, seg)
            x += text_width(seg, font, kv_size)
        y -= leading
    return y

# ---------------- Main ----------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", required=True, help="Path to input .xlsx")
    ap.add_argument("--template", required=True, help="Path to template.pdf (only page size/orientation used)")
    ap.add_argument("--out", default="output_simple.pdf", help="Output PDF path")
    args = ap.parse_args()

    excel_path = Path(args.excel)
    template_path = Path(args.template)
    out_path = Path(args.out)

    # Read first page size from template
    tpl = fitz.open(str(template_path))
    if tpl.page_count < 1:
        raise SystemExit("template.pdf has no pages.")
    p0 = tpl[0]
    page_w, page_h = float(p0.rect.width), float(p0.rect.height)
    tpl.close()

    # Read Excel
    df = pd.read_excel(excel_path)
    df.columns = [str(c).strip() for c in df.columns]

    # case-insensitive mapping to actual Excel headers
    lc_to_orig = {c.lower(): c for c in df.columns}
    keep_cols = [lc_to_orig[k.lower()] for k in template_keys if k.lower() in lc_to_orig]

    # filter & reorder
    df = df[keep_cols]

    # Fonts: hard-coded auto-detection (no CLI arg)
    ar_font = register_arabic_font()  # Arabic-capable or fallback to Helvetica

    # Create output
    cnv = rl_canvas.Canvas(str(out_path), pagesize=(page_w, page_h))
    cnv.setFillColor(black)

    # Layout
    margin_l, margin_r = 40, 40
    margin_t, margin_b = 40, 40
    kv_font_size = 12
    leading = 18
    max_width = page_w - margin_l - margin_r

    for _, row in df.iterrows():
        y = page_h - margin_t
        for col in df.columns:
            raw_val = row[col]
            val = "" if pd.isna(raw_val) else str(raw_val)

            # Build "Key: Value" but keep fonts separate:
            #   - key + ": " goes in Latin font
            #   - value is segmented and drawn mixed-font
            key_prefix = f"{col}: "

            # Draw key prefix first (wrap-aware: we combine as a single string and let the mixed renderer handle it)
            # To keep it simple, we pass the full "key: value" to the mixed renderer,
            # but we force the key portion to be Latin by temporarily marking non-Arabic.
            kv_full = key_prefix + val
            print(f"+++{kv_full}+++")
            # Mixed draw
            y = draw_mixed_line(
                cnv=cnv,
                x_left=margin_l,
                y_baseline=y,
                max_width=max_width,
                latin_font=LATIN_FONT,
                ar_font=ar_font,
                kv_text=kv_full,
                kv_size=kv_font_size,
                leading=leading
            )

            # Page break safety
            if y < margin_b + leading:
                cnv.showPage()
                cnv.setFillColor(black)
                y = page_h - margin_t

        cnv.showPage()

    cnv.save()
    print(f"OK: wrote {out_path}")

if __name__ == "__main__":
    main()
