Excel -> PDF (one page per row), simple "Key: Value" list.
- Page size / orientation is A4 landscape
- NO command-line font argument: script auto-picks an Arabic-capable font
- Draws keys with a Latin font and values with mixed-script support:
    * Arabic segments are shaped (arabic-reshaper + python-bidi) and drawn with an Arabic font
    * Non‑Arabic segments are drawn with a Latin font
    * Segments are concatenated on the same line with correct widths
- Very simple wrapping to page width (left-aligned lines for simplicity)

Run:
    python excel2pdf.py --excel "sample rows.xlsx" --output output.pdf