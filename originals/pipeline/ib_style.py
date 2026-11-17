"""
ib_style.py — Investment-banking / equity-research Excel styling engine.

Encodes the house style distilled from two reference models (a WST-style DCF and a
broker operating model). Import the helpers and build models with them so that every
workbook Claude produces shares one consistent, professional IB look.

Usage:
    import sys; sys.path.insert(0, 'scripts')
    from ib_style import (IB, apply_workbook_defaults, title_block, section_band,
        header_years, label, input_cell, formula_cell, link_cell, subrow_pct,
        total_row, comment, autofit)

Color code (load-bearing convention):
    blue text   0000FF  -> hardcoded input / assumption
    black text  000000  -> formula on the same sheet
    green text  007D00  -> link to another sheet
    red text    FF0000  -> link to another file
    yellow fill FFF2CC  -> key assumption / cell the user should fill in
"""

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


class IB:
    """Named constants: colors, number formats, and the base font size."""

    # --- fonts ---
    FONT = "Arial"
    SIZE = 9          # dense operating models; use 10 for lighter valuation pages
    SIZE_TITLE = 11

    # --- font colors (ARGB, openpyxl style) ---
    BLUE = "FF0000FF"     # hardcoded input
    BLACK = "FF000000"    # same-sheet formula
    GREEN = "FF007D00"    # cross-sheet link
    RED = "FFFF0000"      # cross-file link
    WHITE = "FFFFFFFF"    # text on dark bands

    # --- fills ---
    YELLOW = "FFFFF2CC"       # editable assumption
    BAND_BLACK = "FF000000"   # WST section band
    BAND_BLUE = "FF0018A8"    # broker primary band
    BAND_BLUE_PALE = "FFB0DAF6"  # company/projected estimates
    BAND_GREY = "FFA8BCC2"    # historical
    TOTAL_SHADE = "FFF1F5F7"  # subtle shade behind total rows

    # --- number formats ---
    NUM = "#,##0_);(#,##0);-"                    # integers, () negatives, - zeros
    NUM1 = "#,##0.0_);(#,##0.0);-"               # one decimal
    NUM2 = "#,##0.00_);(#,##0.00);-"             # two decimals
    PCT = "0.0%_);(0.0%);-"                      # percent, stored as fraction
    PCT0 = "0%_);(0%);-"
    MULT = "0.0x"                                # valuation multiple
    CUR = '"$"#,##0_);("$"#,##0);-'
    CUR1 = '"$"#,##0.0_);("$"#,##0.0);-'
    YEAR_A = '0"A"'                              # actual year: 2020A
    YEAR_E = "General"                           # estimate year: pass the string "2022E"

    THIN = Side(style="thin", color="FF000000")


def _font(bold=False, italic=False, color=IB.BLACK, size=None):
    return Font(name=IB.FONT, size=size or IB.SIZE, bold=bold, italic=italic, color=color)


def apply_workbook_defaults(ws, spacer_col_a=True, label_width=25.5, data_width=10.1):
    """Turn gridlines off, set the column-A spacer and standard column widths."""
    ws.sheet_view.showGridLines = False
    if spacer_col_a:
        ws.column_dimensions["A"].width = 0.9
    ws.column_dimensions["B"].width = label_width
    for i in range(3, 12):  # C..K default data width
        ws.column_dimensions[get_column_letter(i)].width = data_width


def title_block(ws, title, units=None, subtitle=None, row=1, col=1,
                dark=False, company=None):
    """
    Write the top-of-sheet title block.

    dark=True renders the WST cover/valuation style: black fill, white text,
    optional company name line above the title.
    """
    r = row
    if dark:
        # company / title / units all on a black band with white text
        if company:
            c = ws.cell(row=r, column=col, value=company)
            c.font = _font(bold=True, color=IB.WHITE, size=IB.SIZE_TITLE + 3)
            c.fill = PatternFill("solid", fgColor=IB.BAND_BLACK)
            r += 1
        c = ws.cell(row=r, column=col, value=title)
        c.font = _font(bold=True, color=IB.WHITE, size=IB.SIZE_TITLE + 1)
        c.fill = PatternFill("solid", fgColor=IB.BAND_BLACK)
        r += 1
        if units:
            c = ws.cell(row=r, column=col, value=units)
            c.font = _font(italic=True, color=IB.WHITE, size=IB.SIZE - 1)
            c.fill = PatternFill("solid", fgColor=IB.BAND_BLACK)
        return r
    # light style
    c = ws.cell(row=r, column=col, value=title)
    c.font = _font(bold=True, size=IB.SIZE)
    r += 1
    if units:
        c = ws.cell(row=r, column=col, value=units)
        c.font = _font(italic=True)
        r += 1
    if subtitle:
        c = ws.cell(row=r, column=col, value=subtitle)
        c.font = _font(italic=True)
    return r


def section_band(ws, row, start_col, end_col, texts, style="blue"):
    """
    Draw a period-type / section band spanning columns via centerContinuous
    (NOT merged cells). `texts` is a dict {col_index: label} placed at the
    leftmost cell of each sub-span; all cells in the range get the fill.

    style: 'black' | 'blue' | 'grey' | 'pale' | 'hist' (historical grey +
    projected pale — pass texts spanning both and it will color per your fills).
    """
    palette = {
        "black": (IB.BAND_BLACK, IB.WHITE),
        "blue": (IB.BAND_BLUE, IB.WHITE),
        "grey": (IB.BAND_GREY, IB.BLACK),
        "pale": (IB.BAND_BLUE_PALE, IB.BLACK),
    }
    fill_rgb, font_rgb = palette.get(style, palette["blue"])
    for col in range(start_col, end_col + 1):
        c = ws.cell(row=row, column=col)
        c.fill = PatternFill("solid", fgColor=fill_rgb)
        c.alignment = Alignment(horizontal="centerContinuous")
        c.font = _font(color=font_rgb)
        if col in texts:
            c.value = texts[col]


def header_years(ws, row, start_col, years, band_fill=IB.BAND_GREY,
                 est_from_col=None, fmt=None):
    """
    Write a bold year header row. `years` is a list of values (ints like 2020,
    or strings like '2022E'). Ints get the 'A' actual format; strings are shown
    verbatim (use for estimates like '2022E'). Optional band_fill shades the row.
    """
    for i, y in enumerate(years):
        col = start_col + i
        c = ws.cell(row=row, column=col, value=y)
        c.font = _font(bold=True, color=IB.BLUE if isinstance(y, int) else IB.BLACK)
        c.alignment = Alignment(horizontal="center")
        if band_fill:
            c.fill = PatternFill("solid", fgColor=band_fill)
        c.number_format = fmt or (IB.YEAR_A if isinstance(y, int) else IB.YEAR_E)


def label(ws, row, text, col=2, bold=False, italic=False, indent=0):
    c = ws.cell(row=row, column=col, value=text)
    c.font = _font(bold=bold, italic=italic)
    if indent:
        c.alignment = Alignment(indent=indent)
    return c


def _write(ws, row, col, value, color, bold, italic, numfmt, fill):
    c = ws.cell(row=row, column=col, value=value)
    c.font = _font(bold=bold, italic=italic, color=color)
    c.number_format = numfmt
    c.alignment = Alignment(horizontal="right")
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    return c


def input_cell(ws, row, col, value, numfmt=IB.NUM, bold=False, highlight=False):
    """Hardcoded input -> BLUE text. highlight=True adds the yellow fill-in fill."""
    return _write(ws, row, col, value, IB.BLUE, bold, False, numfmt,
                  IB.YELLOW if highlight else None)


def formula_cell(ws, row, col, formula, numfmt=IB.NUM, bold=False, fill=None):
    """Same-sheet formula -> BLACK text. Pass the '=' formula string."""
    return _write(ws, row, col, formula, IB.BLACK, bold, False, numfmt, fill)


def link_cell(ws, row, col, formula, numfmt=IB.NUM, bold=False, fill=None):
    """Cross-sheet link -> GREEN text. Remember to quote sheet names with spaces."""
    return _write(ws, row, col, formula, IB.GREEN, bold, False, numfmt, fill)


def subrow_pct(ws, row, col, formula_or_value, numfmt=IB.PCT):
    """Italic %-growth / %-margin subrow. Accepts a formula string or a fraction."""
    c = ws.cell(row=row, column=col, value=formula_or_value)
    c.font = _font(italic=True)
    c.number_format = numfmt
    c.alignment = Alignment(horizontal="right")
    return c


def total_row(ws, row, label_text, formulas, start_col=3, label_col=2,
              numfmt=IB.NUM, shade=True):
    """
    Bold total/subtotal row with subtle shading. `formulas` is a list aligned to
    columns starting at start_col; each item may be a formula string or a number.
    """
    lc = ws.cell(row=row, column=label_col, value=label_text)
    lc.font = _font(bold=True)
    if shade:
        lc.fill = PatternFill("solid", fgColor=IB.TOTAL_SHADE)
    for i, f in enumerate(formulas):
        col = start_col + i
        c = ws.cell(row=row, column=col, value=f)
        c.font = _font(bold=True)
        c.number_format = numfmt
        c.alignment = Alignment(horizontal="right")
        if shade:
            c.fill = PatternFill("solid", fgColor=IB.TOTAL_SHADE)


def comment(ws, row, text, col=11, header=False):
    """Write into the far 'Key comments' column (default col K)."""
    c = ws.cell(row=row, column=col, value=text)
    if header:
        c.font = _font(bold=True)
        c.border = Border(bottom=IB.THIN)
    else:
        c.font = _font()
    ws.column_dimensions[get_column_letter(col)].width = 65
    return c


def source_row(ws, row, source_text, col=1, fillin=False):
    """
    Write a source-citation line on a RAW-DATA sheet, e.g.
        '출처: 모두싸인(주) 개별 감사보고서 FY2024 (DART 전자공시)'
        'Source: Company 10-K, FY2024, p.45 [SEC EDGAR URL]'
    Italic grey. If the exact source is unknown and the user must supply it,
    set fillin=True to render it as a blue/yellow fill-in placeholder.
    """
    c = ws.cell(row=row, column=col, value=source_text)
    if fillin:
        c.font = _font(italic=True, color=IB.BLUE)
        c.fill = PatternFill("solid", fgColor=IB.YELLOW)
    else:
        c.font = _font(italic=True, color="FF595959")
    return c


def recolor_by_type(ws, first_row=1, last_row=None, first_col=1, last_col=None,
                    input_fmt=None):
    """
    Enforce the color code across a range by inspecting each cell's *content*:
        number            -> BLUE  (raw / hardcoded input)
        '=...!...'         -> GREEN (link to another sheet)
        '=[...]...'        -> RED   (link to another file)
        '=...' (else)      -> BLACK (same-sheet formula)
        text               -> left unchanged (labels, headers)
    This is the fastest way to make an inherited workbook obey the convention.
    Preserves bold/italic/size and only rewrites the font *color*; optionally
    stamps `input_fmt` (e.g. IB.NUM) onto numeric cells whose format is General.
    Percent cells (number_format containing '%') keep their format.
    """
    last_row = last_row or ws.max_row
    last_col = last_col or ws.max_column
    for r in range(first_row, last_row + 1):
        for cc in range(first_col, last_col + 1):
            cell = ws.cell(row=r, column=cc)
            v = cell.value
            if v is None:
                continue
            if isinstance(v, str) and v.startswith("="):
                if "]" in v.split("!")[0]:          # [1]Sheet! -> external file
                    color = IB.RED
                elif "!" in v:                       # Sheet! -> cross-sheet link
                    color = IB.GREEN
                else:
                    color = IB.BLACK                 # same-sheet formula
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                color = IB.BLUE                      # raw input
                if input_fmt and cell.number_format in ("General", "0", "0.00"):
                    cell.number_format = input_fmt
            else:
                continue                             # text label: leave alone
            f = cell.font
            cell.font = Font(name=f.name or IB.FONT, size=f.size or IB.SIZE,
                             bold=f.bold, italic=f.italic, color=color)


def set_font_everywhere(ws, name, size=None):
    """Force one consistent font family across every populated cell (keeps
    bold/italic/color/size). Use a CJK-capable face (e.g. '맑은 고딕',
    'Malgun Gothic', 'Noto Sans CJK KR') when the content is Korean —
    Arial cannot render Hangul cleanly."""
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            f = cell.font
            cell.font = Font(name=name, size=size or f.size or IB.SIZE,
                             bold=f.bold, italic=f.italic, color=f.color)


def autofit(ws, min_width=8.5, max_width=70, sample_rows=200):
    """Approximate autofit that respects the col-A spacer and label column."""
    for col_cells in ws.columns:
        col_letter = col_cells[0].column_letter
        if col_letter == "A":
            continue
        longest = 0
        for cell in col_cells[:sample_rows]:
            v = cell.value
            if v is None:
                continue
            s = str(v)
            if s.startswith("="):  # don't size to formula text
                s = "0000000"
            longest = max(longest, len(s))
        width = max(min_width, min(max_width, longest + 2))
        # never shrink an explicitly-set wide comment column
        cur = ws.column_dimensions[col_letter].width
        if cur and cur > width:
            continue
        ws.column_dimensions[col_letter].width = width
