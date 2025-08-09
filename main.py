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

import os
import argparse
from pathlib import Path
from typing import Optional, List, Tuple

import fitz
import pandas as pd

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
    
# ONLY list columns that are required
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

title_fields = ["region", "city", "hay", "pca"]

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
    df_full = df.copy()

    # Drop title fields from the key-value printing DataFrame
    df = df[[c for c in df.columns if c.lower() not in [t.lower() for t in title_fields]]]


    # Fonts: hard-coded auto-detection (no CLI arg)
    ar_font = register_arabic_font()  # Arabic-capable or fallback to Helvetica

    # Create output
    cnv = rl_canvas.Canvas(str(out_path), pagesize=(page_w, page_h))
    cnv.setFillColor(black)

    # Layout constants
    margin_l, margin_r = 40, 40
    margin_t, margin_b = 40, 40
    kv_font_size = 12
    leading = 18
    gap_between_pairs = 6

    # Calculate two column widths and positions
    col_gap = 40
    usable_width = page_w - margin_l - margin_r - col_gap
    col_width = usable_width / 2
    left_x = margin_l
    right_x = margin_l + col_width + col_gap

    for idx, row in df.iterrows():
        # Build title from the original full DF so we still have region/city/hay/pca
        row_full = df_full.iloc[idx]
        title_text = f"{row_full['Region']} – {row_full['City']} – {row_full['Hay']} – {row_full['PCA']}"

        # Use mixed-script renderer for title
        title_font_size = 14
        y_title_top = page_h - margin_t
        y_after_title = draw_mixed_line(
            cnv=cnv,
            x_left=margin_l,
            y_baseline=y_title_top,
            max_width=page_w - margin_l - margin_r,
            latin_font=LATIN_FONT,
            ar_font=ar_font,
            kv_text=title_text,
            kv_size=title_font_size,
            leading=title_font_size + 6  # line height for title
        )

        # Add spacing before key-value columns start
        title_to_columns_gap = 20
        y_left = y_after_title - title_to_columns_gap
        y_right = y_after_title - title_to_columns_gap

        col_toggle = "left"

        for col in df.columns:
            raw_val = row[col]
            val = "" if pd.isna(raw_val) else str(raw_val)
            kv_full = f"{col}: {val}"

            if col_toggle == "left":
                y_left = draw_mixed_line(
                    cnv=cnv,
                    x_left=left_x,
                    y_baseline=y_left,
                    max_width=col_width,
                    latin_font=LATIN_FONT,
                    ar_font=ar_font,
                    kv_text=kv_full,
                    kv_size=kv_font_size,
                    leading=leading
                )
                y_left -= gap_between_pairs
                col_toggle = "right"
            else:
                y_right = draw_mixed_line(
                    cnv=cnv,
                    x_left=right_x,
                    y_baseline=y_right,
                    max_width=col_width,
                    latin_font=LATIN_FONT,
                    ar_font=ar_font,
                    kv_text=kv_full,
                    kv_size=kv_font_size,
                    leading=leading
                )
                y_right -= gap_between_pairs
                col_toggle = "left"

        cnv.showPage()

    cnv.save()
    print(f"PDF output saved to --> {out_path}.")

if __name__ == "__main__":
    main()

# python main.py --excel "sample rows.xlsx" --template template.pdf --out output.pdf
