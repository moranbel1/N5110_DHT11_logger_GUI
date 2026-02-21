"""
make_smart_excel.py
Reads Test.xlsx (personal-finance workbook) and produces Test_smart.xlsx with:
  - Dashboard sheet (key metrics + portfolio pie chart)
  - Enhanced copies of every original sheet (formatting, auto-filter, freeze)
  - Monthly-spending line chart (עו"ש sheet)
  - Loan-payment bar chart (הלוואות sheet)
  - Pension growth bar chart (פנסייה sheet)
  - Sorting-guide sheet

Usage:
    python3 make_smart_excel.py [input.xlsx] [output.xlsx]
"""

import sys
import copy
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, GradientFill
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.chart import LineChart, BarChart, PieChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# ── Colour palette ────────────────────────────────────────────────────────────
DARK_BLUE   = "1F4E79"
MED_BLUE    = "2E75B6"
LIGHT_BLUE  = "DEEAF1"
ORANGE      = "ED7D31"
GREEN       = "70AD47"
YELLOW      = "FFF2CC"
RED         = "FF6347"
WHITE       = "FFFFFF"
GRAY_LIGHT  = "F2F2F2"
GRAY_MED    = "BFBFBF"
GOLD        = "FFD700"

def _side(style="thin", color=GRAY_MED):
    return Side(border_style=style, color=color)

def _border(style="thin"):
    s = _side(style)
    return Border(left=s, right=s, top=s, bottom=s)

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _font(bold=False, color="000000", size=11, italic=False):
    return Font(bold=bold, color=color, size=size, name="Calibri", italic=italic)

def _hdr(ws, row, col, value, bg=DARK_BLUE, fg=WHITE, size=11):
    c = ws.cell(row=row, column=col, value=value)
    c.font      = _font(bold=True, color=fg, size=size)
    c.fill      = _fill(bg)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border    = _border()
    return c

def _val(ws, row, col, value, bg=WHITE, bold=False, align="center", num_fmt=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font      = _font(bold=bold, size=10)
    c.fill      = _fill(bg)
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    c.border    = _border()
    if num_fmt:
        c.number_format = num_fmt
    return c

# ── helpers ───────────────────────────────────────────────────────────────────

def fmt_num(v):
    """Return ₪ formatted string or original."""
    try:
        return f"₪{float(v):,.0f}"
    except (TypeError, ValueError):
        return str(v) if v is not None else ""

def safe_float(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None

# ═══════════════════════════════════════════════════════════════════════════════
# DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════════

def build_dashboard(wb_out, wb_in):
    ws = wb_out.create_sheet("📊 דשבורד", 0)
    ws.sheet_view.rightToLeft = True

    # ── Title ──
    ws.merge_cells("A1:H1")
    t = ws.cell(row=1, column=1, value="📋 סיכום פיננסי אישי – מורן ויערה")
    t.font      = _font(bold=True, color=WHITE, size=16)
    t.fill      = _fill(DARK_BLUE)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 40

    ws.merge_cells("A2:H2")
    sub = ws.cell(row=2, column=1, value=f"עודכן לאחרונה: ינואר 2025")
    sub.font      = _font(italic=True, color=DARK_BLUE, size=10)
    sub.fill      = _fill(LIGHT_BLUE)
    sub.alignment = Alignment(horizontal="center")
    ws.row_dimensions[2].height = 18

    # ── KPI cards ──
    kpis = [
        ("פנסיה מורן (כלל)",     "₪588,339",   MED_BLUE,  WHITE),
        ("פנסיה יערה (מנורה)",   "₪728,571",   MED_BLUE,  WHITE),
        ("קרן השתלמות מורן",     "₪220,168",   GREEN,     WHITE),
        ("קרן השתלמות יערה",     "₪238,895",   GREEN,     WHITE),
        ("הלוואה יערה (נותרו)",  "₪~38,000",   RED,       WHITE),
        ("הוצאות חודשיות ממוצע", "₪15,437",    ORANGE,    WHITE),
        ("חיסכון חודשי (תחזית)", "₪9,361",     GREEN,     WHITE),
        ("סה\"כ ביטוחים/חודש",   "₪888",        GOLD,      "000000"),
    ]

    row = 4
    ws.merge_cells(f"A{row}:H{row}")
    h = ws.cell(row=row, column=1, value="מדדי מפתח")
    h.font      = _font(bold=True, color=WHITE, size=12)
    h.fill      = _fill(DARK_BLUE)
    h.alignment = Alignment(horizontal="center")
    ws.row_dimensions[row].height = 22

    row = 5
    for i, (label, value, bg, fg) in enumerate(kpis):
        col = i + 1
        # Label cell
        lc = ws.cell(row=row, column=col, value=label)
        lc.font      = _font(bold=True, color=fg, size=9)
        lc.fill      = _fill(bg)
        lc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        lc.border    = _border()
        ws.row_dimensions[row].height = 30
        # Value cell
        vc = ws.cell(row=row+1, column=col, value=value)
        vc.font      = _font(bold=True, color=bg, size=12)
        vc.fill      = _fill(GRAY_LIGHT)
        vc.alignment = Alignment(horizontal="center", vertical="center")
        vc.border    = _border()
        ws.row_dimensions[row+1].height = 28

    # ── Portfolio allocation table ──
    row = 9
    ws.merge_cells(f"A{row}:D{row}")
    h2 = ws.cell(row=row, column=1, value="הרכב תיק ההשקעות")
    h2.font      = _font(bold=True, color=WHITE, size=12)
    h2.fill      = _fill(MED_BLUE)
    h2.alignment = Alignment(horizontal="center")
    ws.row_dimensions[row].height = 22

    portfolio = [
        ("פנסיה מורן",          570000),
        ("קרן השתלמות מורן",    220000),
        ("קרן השתלמות לא פעיל", 52400),
        ("תיק מסחר",            24000),
        ("קרן ביטחון",          108000),
        ("עו\"ש פנוי",           20000),
    ]
    total_p = sum(v for _, v in portfolio)

    _hdr(ws, row+1, 1, "נכס",      bg=MED_BLUE)
    _hdr(ws, row+1, 2, "שווי (₪)",  bg=MED_BLUE)
    _hdr(ws, row+1, 3, "אחוז",     bg=MED_BLUE)
    _hdr(ws, row+1, 4, "ערבה",     bg=MED_BLUE)
    ws.row_dimensions[row+1].height = 22

    fills = [LIGHT_BLUE, WHITE]
    for j, (name, val) in enumerate(portfolio, start=row+2):
        bg = fills[j % 2]
        _val(ws, j, 1, name,                       bg=bg, align="right")
        _val(ws, j, 2, val,                         bg=bg, num_fmt='#,##0 ₪')
        pct = val / total_p if total_p else 0
        _val(ws, j, 3, pct,                         bg=bg, num_fmt='0.0%')
        # simple bar using block characters
        blocks = int(pct * 20)
        _val(ws, j, 4, "█" * blocks,                bg=bg)
        ws.row_dimensions[j].height = 18

    # Total row
    total_row = row + 2 + len(portfolio)
    _val(ws, total_row, 1, "סה\"כ", bg=DARK_BLUE, bold=True).font = _font(bold=True, color=WHITE)
    _val(ws, total_row, 2, total_p, bg=DARK_BLUE, num_fmt='#,##0 ₪').font = _font(bold=True, color=WHITE)
    _val(ws, total_row, 3, 1.0,    bg=DARK_BLUE, num_fmt='0%').font = _font(bold=True, color=WHITE)
    ws.cell(total_row, 2).fill = _fill(DARK_BLUE)
    ws.cell(total_row, 3).fill = _fill(DARK_BLUE)

    # ── Pie chart of portfolio ──
    # Write helper data for chart (hidden columns)
    chart_start_col = 10
    ws.cell(row=9,  column=chart_start_col, value="נכס")
    ws.cell(row=9,  column=chart_start_col+1, value="שווי")
    for j, (name, val) in enumerate(portfolio, start=10):
        ws.cell(row=j, column=chart_start_col,   value=name)
        ws.cell(row=j, column=chart_start_col+1, value=val)
        ws.column_dimensions[get_column_letter(chart_start_col)].width = 0.1
        ws.column_dimensions[get_column_letter(chart_start_col+1)].width = 0.1

    pie = PieChart()
    pie.title  = "הרכב תיק ההשקעות"
    pie.style  = 10
    pie.height = 14
    pie.width  = 20

    data_ref = Reference(ws, min_col=chart_start_col+1, min_row=9,
                         max_row=9+len(portfolio))
    cats_ref = Reference(ws, min_col=chart_start_col,   min_row=10,
                         max_row=9+len(portfolio))
    pie.add_data(data_ref, titles_from_data=True)
    pie.set_categories(cats_ref)
    ws.add_chart(pie, "E4")

    # ── Column widths ──
    for col_letter, width in zip("ABCDEFGH", [22, 14, 8, 12, 14, 14, 14, 14]):
        ws.column_dimensions[col_letter].width = width


# ═══════════════════════════════════════════════════════════════════════════════
# MONTHLY SPENDING CHART (עו"ש)
# ═══════════════════════════════════════════════════════════════════════════════

def build_spending_chart(wb_out, wb_in):
    src = wb_in['עו"ש']
    ws  = wb_out.create_sheet('גרף הוצאות')
    ws.sheet_view.rightToLeft = True

    # Pull monthly data: col B=date, C=moran, D=yaara, E=total
    rows_data = []
    for row in src.iter_rows(min_row=3, max_row=src.max_row, values_only=True):
        date_val = row[1]  # column B (0-indexed = 1)
        moran    = safe_float(row[2])
        yaara    = safe_float(row[3])
        total    = safe_float(row[4])
        if date_val and (moran is not None or yaara is not None):
            if isinstance(date_val, datetime.datetime):
                lbl = date_val.strftime("%m/%Y")
            else:
                lbl = str(date_val)
            rows_data.append((lbl, moran or 0, yaara or 0, total or 0))

    if not rows_data:
        return

    n = len(rows_data) + 1

    # Header row
    for col, hdr in enumerate(["תאריך", "מורן (₪)", "יערה (₪)", "סה\"כ (₪)"], start=1):
        _hdr(ws, 1, col, hdr, bg=MED_BLUE)
    for i, (lbl, m, y, t) in enumerate(rows_data, start=2):
        bg = LIGHT_BLUE if i % 2 == 0 else WHITE
        ws.cell(i, 1, lbl).alignment = Alignment(horizontal="center")
        _val(ws, i, 2, m, bg=bg, num_fmt='#,##0')
        _val(ws, i, 3, y, bg=bg, num_fmt='#,##0')
        _val(ws, i, 4, t, bg=bg, num_fmt='#,##0')

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14

    # ── Line chart: Moran + Yaara ──
    chart = LineChart()
    chart.title  = "הוצאות חודשיות – מורן, יערה וסה\"כ"
    chart.style  = 10
    chart.height = 14
    chart.width  = 26
    chart.y_axis.title = "₪"
    chart.x_axis.title = "חודש"
    chart.smooth = True

    data_ref = Reference(ws, min_col=2, min_row=1, max_col=4, max_row=n)
    cats_ref = Reference(ws, min_col=1, min_row=2, max_row=n)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    colors = [ORANGE, MED_BLUE, GREEN]
    for i, color in enumerate(colors):
        chart.series[i].graphicalProperties.line.solidFill = color
        chart.series[i].graphicalProperties.line.width = 20000

    ws.add_chart(chart, "F2")

    # ── Bar chart: total per month ──
    bar = BarChart()
    bar.title   = "סה\"כ הוצאות חודשי"
    bar.style   = 10
    bar.type    = "col"
    bar.height  = 14
    bar.width   = 26
    bar.y_axis.title = "₪"
    bar.x_axis.title = "חודש"

    bar_data = Reference(ws, min_col=4, min_row=1, max_row=n)
    bar.add_data(bar_data, titles_from_data=True)
    bar.set_categories(cats_ref)
    bar.series[0].graphicalProperties.solidFill = MED_BLUE

    ws.add_chart(bar, "F30")

    # ── Conditional formatting on totals ──
    total_range = f"D2:D{n}"
    ws.conditional_formatting.add(
        total_range,
        ColorScaleRule(
            start_type="min",  start_color="63B3FF",
            mid_type="percentile", mid_value=50, mid_color="FFFF00",
            end_type="max",    end_color="FF0000",
        )
    )


# ═══════════════════════════════════════════════════════════════════════════════
# LOAN PAYMENT CHART (הלוואות)
# ═══════════════════════════════════════════════════════════════════════════════

def build_loan_chart(wb_out, wb_in):
    src = wb_in["הלוואות"]
    ws  = wb_out.create_sheet("גרף הלוואה")
    ws.sheet_view.rightToLeft = True

    rows_data = []
    for row in src.iter_rows(min_row=7, max_row=src.max_row, values_only=True):
        date_val = row[1]
        payment  = safe_float(row[2])
        if date_val and isinstance(date_val, datetime.datetime):
            lbl = date_val.strftime("%m/%Y")
            rows_data.append((lbl, payment or 0, payment is not None))

    if not rows_data:
        return

    # Write data
    _hdr(ws, 1, 1, "חודש",        bg=MED_BLUE)
    _hdr(ws, 1, 2, "תשלום (₪)",   bg=MED_BLUE)
    _hdr(ws, 1, 3, "שולם?",       bg=MED_BLUE)

    for i, (lbl, amt, paid) in enumerate(rows_data, start=2):
        bg = LIGHT_BLUE if paid else GRAY_LIGHT
        _val(ws, i, 1, lbl, bg=bg)
        _val(ws, i, 2, amt if amt else "", bg=bg, num_fmt='#,##0')
        _val(ws, i, 3, "✅" if paid else "⏳", bg=bg)

    n = len(rows_data) + 1
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 10

    # Bar chart: payments over time
    bar = BarChart()
    bar.title   = "היסטוריית תשלומי הלוואה יערה"
    bar.style   = 10
    bar.type    = "col"
    bar.height  = 14
    bar.width   = 28
    bar.y_axis.title = "תשלום (₪)"
    bar.x_axis.title = "חודש"

    data_ref = Reference(ws, min_col=2, min_row=1, max_row=n)
    cats_ref = Reference(ws, min_col=1, min_row=2, max_row=n)
    bar.add_data(data_ref, titles_from_data=True)
    bar.set_categories(cats_ref)
    bar.series[0].graphicalProperties.solidFill = MED_BLUE

    ws.add_chart(bar, "E2")

    # Info box
    info_row = 2
    info_col = 5 + 14  # after chart
    loan_info = [
        ("סכום הלוואה", "₪66,252"),
        ("ריבית",       "P − 0.5%"),
        ("החזר חודשי",  "₪1,269"),
        ("תאריך סיום",  "אוקטובר 2028"),
    ]
    ws.merge_cells(f"S1:T1")
    h = ws.cell(1, 19, "פרטי ההלוואה")
    h.font = _font(bold=True, color=WHITE, size=12)
    h.fill = _fill(DARK_BLUE)
    h.alignment = Alignment(horizontal="center")
    for j, (k, v) in enumerate(loan_info, start=2):
        ws.cell(j, 19, k).font = _font(bold=True)
        ws.cell(j, 20, v).font = _font()
        for c in [19, 20]:
            ws.cell(j, c).border = _border()
            ws.cell(j, c).fill   = _fill(LIGHT_BLUE if j % 2 == 0 else WHITE)


# ═══════════════════════════════════════════════════════════════════════════════
# PENSION GROWTH CHART
# ═══════════════════════════════════════════════════════════════════════════════

def build_pension_chart(wb_out, wb_in):
    ws = wb_out.create_sheet("גרף פנסיה")
    ws.sheet_view.rightToLeft = True

    # Dates and values extracted from pension sheets
    moran_data = [
        ("ינואר 22",  305174),
        ("מאי 22",    316245),
        ("אוקטובר 22", None),
        ("ינואר 23",  345916),
        ("יוני 23",   None),
        ("יולי 24",   533665),
        ("נובמבר 24", 563121),
        ("מאי 25",    568209),
    ]
    yaara_data = [
        ("הנוכחי",    693564),
    ]

    # Extended moran (clean only non-None)
    moran_clean = [(d, v) for d, v in moran_data if v is not None]

    _hdr(ws, 1, 1, "תאריך",              bg=MED_BLUE)
    _hdr(ws, 1, 2, "פנסיה מורן (₪)",      bg=MED_BLUE)

    for i, (d, v) in enumerate(moran_clean, start=2):
        bg = LIGHT_BLUE if i % 2 == 0 else WHITE
        _val(ws, i, 1, d, bg=bg)
        _val(ws, i, 2, v, bg=bg, num_fmt='#,##0')

    n = len(moran_clean) + 1
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 20

    bar = BarChart()
    bar.title   = "צמיחת הפנסיה – מורן (כלל / הפניקס)"
    bar.style   = 10
    bar.type    = "col"
    bar.height  = 14
    bar.width   = 22
    bar.y_axis.title  = "₪"
    bar.x_axis.title  = "תאריך"

    data_ref = Reference(ws, min_col=2, min_row=1, max_row=n)
    cats_ref = Reference(ws, min_col=1, min_row=2, max_row=n)
    bar.add_data(data_ref, titles_from_data=True)
    bar.set_categories(cats_ref)
    bar.series[0].graphicalProperties.solidFill = MED_BLUE

    ws.add_chart(bar, "D2")

    # Yaara table
    r = n + 3
    ws.merge_cells(f"A{r}:B{r}")
    h = ws.cell(r, 1, "פנסיה יערה (מנורה)")
    h.font = _font(bold=True, color=WHITE)
    h.fill = _fill(GREEN)
    h.alignment = Alignment(horizontal="center")
    yaara_rows = [
        ("פנסייה מקיפה",   693564),
        ("פנסייה משלימה",   35007),
        ("קרן השתלמות 1",   94240),
        ("קרן השתלמות 2",  144656),
    ]
    for j, (name, val) in enumerate(yaara_rows, start=r+1):
        bg = LIGHT_BLUE if j % 2 == 0 else WHITE
        _val(ws, j, 1, name, bg=bg, align="right")
        _val(ws, j, 2, val,  bg=bg, num_fmt='#,##0 ₪')


# ═══════════════════════════════════════════════════════════════════════════════
# ENHANCED COPY OF ORIGINAL SHEETS
# ═══════════════════════════════════════════════════════════════════════════════

def enhance_sheet(ws_out, ws_in, sheet_name):
    """Copy all cells with improved formatting."""
    # Collect merged cell ranges from source
    merged_ranges = {str(m) for m in ws_in.merged_cells.ranges}
    # Build set of non-top-left merged cells to skip
    merged_skip = set()
    for merge in ws_in.merged_cells.ranges:
        cells = list(ws_in[str(merge)])
        for mrow in cells[1:]:          # skip first row of merged range (top-left)
            for mc in mrow:
                merged_skip.add((mc.row, mc.column))

    # Copy dimensions
    for col in ws_in.column_dimensions:
        ws_out.column_dimensions[col].width = ws_in.column_dimensions[col].width or 15

    # Copy merged cells
    for merge in ws_in.merged_cells.ranges:
        try:
            ws_out.merge_cells(str(merge))
        except Exception:
            pass

    header_done = False
    for r_idx, row in enumerate(ws_in.iter_rows(), start=1):
        for cell in row:
            # Skip non-top-left merged cells
            if (cell.row, cell.column) in merged_skip:
                continue
            dest = ws_out.cell(row=cell.row, column=cell.column)
            try:
                dest.value = cell.value
            except AttributeError:
                continue

            # Style: first non-empty row = header style
            if not header_done and cell.value is not None:
                dest.font      = _font(bold=True, color=WHITE, size=10)
                dest.fill      = _fill(DARK_BLUE)
                dest.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                dest.border    = _border()
            else:
                # Alternate row colors for rows with data
                if any(c.value is not None for c in row):
                    bg = LIGHT_BLUE if r_idx % 2 == 0 else WHITE
                    dest.fill   = _fill(bg)
                    dest.border = _border()
                    dest.font   = _font(size=9)
                    dest.alignment = Alignment(wrap_text=True, vertical="center")

                    # Numbers: right-align, auto-format
                    if isinstance(cell.value, (int, float)) and cell.value != 0:
                        dest.alignment = Alignment(horizontal="right", vertical="center")
                        if abs(cell.value) >= 1000:
                            dest.number_format = '#,##0'

                    # Dates
                    if isinstance(cell.value, datetime.datetime):
                        dest.number_format = "DD/MM/YYYY"
                        dest.alignment = Alignment(horizontal="center", vertical="center")

        # Mark first non-empty row as done
        if not header_done and any(c.value is not None for c in row):
            header_done = True
        ws_out.row_dimensions[r_idx].height = max(
            ws_in.row_dimensions[r_idx].height or 15, 15
        )

    # Freeze top row
    ws_out.freeze_panes = "A2"


# ═══════════════════════════════════════════════════════════════════════════════
# SORTING GUIDE
# ═══════════════════════════════════════════════════════════════════════════════

def build_sort_guide(wb_out):
    ws = wb_out.create_sheet("💡 מדריך מיון")
    ws.sheet_view.rightToLeft = True

    ws.merge_cells("A1:E1")
    t = ws.cell(1, 1, "מדריך מיון ופילטור – קובץ פיננסי אישי")
    t.font      = _font(bold=True, color=WHITE, size=14)
    t.fill      = _fill(DARK_BLUE)
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    guides = [
        # (sheet, sort_col, direction, why)
        ('עו"ש',        "עמודת סה\"כ (E)",         "גבוה→נמוך",   "לזיהוי חודשים יקרים"),
        ('עו"ש',        "עמודת תאריך (B)",           "ישן→חדש",     "סדר כרונולוגי"),
        ('עו"ש',        "עמודת מורן (C)",            "גבוה→נמוך",   "להשוואת הוצאות אישיות"),
        ('הלוואות',     "עמודת תאריך (B)",           "ישן→חדש",     "מעקב לפי זמן"),
        ('הלוואות',     "עמודת תשלום (C)",           "גבוה→נמוך",   "חודשי התשלום הגבוהים"),
        ('ביטוח רפואי', "עמודת פרמיה (D)",           "גבוה→נמוך",   "לזיהוי הביטוחים היקרים"),
        ('ביטוח רפואי', "עמודת תאריך סיום (G)",      "קרוב→רחוק",   "ביטוחים הפגים הכי מהר"),
        ('מחשבון דמי ניהול', "עמודת יתרה סופית (E)", "גבוה→נמוך",   "לבחירת האלטרנטיבה הטובה"),
        ('פנסייה',      "עמודת צבירה",               "גבוה→נמוך",   "להשוואת גדלי הצבירות"),
    ]

    headers = ["גיליון", "מיין לפי", "כיוון", "מטרה", "כיצד?"]
    for c, h in enumerate(headers, start=1):
        _hdr(ws, 2, c, h, bg=MED_BLUE)
    ws.row_dimensions[2].height = 22

    fills = [LIGHT_BLUE, WHITE]
    for i, (sheet, col, direction, why) in enumerate(guides, start=3):
        bg = fills[i % 2]
        _val(ws, i, 1, sheet,      bg=bg)
        _val(ws, i, 2, col,        bg=bg)
        _val(ws, i, 3, direction,  bg=bg)
        _val(ws, i, 4, why,        bg=bg)
        _val(ws, i, 5, "נתונים ← מיון ← בחר עמודה", bg=bg)
        ws.row_dimensions[i].height = 20

    # Tip box
    tip_row = len(guides) + 5
    ws.merge_cells(f"A{tip_row}:E{tip_row}")
    tip = ws.cell(tip_row, 1,
        "💡 טיפ: השתמש ב-AutoFilter (Ctrl+Shift+L) על כל טבלה כדי לפלטר לפי קריטריונים מרובים")
    tip.font      = _font(bold=True, color=DARK_BLUE, size=11)
    tip.fill      = _fill(YELLOW)
    tip.alignment = Alignment(horizontal="center", vertical="center")
    tip.border    = _border("medium")
    ws.row_dimensions[tip_row].height = 28

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 28
    ws.column_dimensions["E"].width = 30


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def make_smart_excel(input_path, output_path):
    print(f"Reading: {input_path}")
    wb_in = openpyxl.load_workbook(input_path, data_only=True)
    print(f"Sheets found: {wb_in.sheetnames}")

    wb_out = openpyxl.Workbook()
    # Remove default blank sheet
    if "Sheet" in wb_out.sheetnames:
        del wb_out["Sheet"]

    # 1. Dashboard
    print("Building dashboard...")
    build_dashboard(wb_out, wb_in)

    # 2. Spending chart
    print("Building spending chart...")
    build_spending_chart(wb_out, wb_in)

    # 3. Loan chart
    print("Building loan chart...")
    build_loan_chart(wb_out, wb_in)

    # 4. Pension chart
    print("Building pension chart...")
    build_pension_chart(wb_out, wb_in)

    # 5. Enhanced copies of original sheets
    for name in wb_in.sheetnames:
        print(f"Enhancing sheet: {name}")
        ws_out = wb_out.create_sheet(name)
        ws_out.sheet_view.rightToLeft = True
        enhance_sheet(ws_out, wb_in[name], name)

    # 6. Sorting guide
    print("Building sort guide...")
    build_sort_guide(wb_out)

    wb_out.save(output_path)
    print(f"\n✅ Smart Excel saved to: {output_path}")
    print("Sheets created:")
    for name in wb_out.sheetnames:
        print(f"  • {name}")


if __name__ == "__main__":
    input_file  = sys.argv[1] if len(sys.argv) >= 2 else "Test.xlsx"
    output_file = sys.argv[2] if len(sys.argv) >= 3 else "Test_smart.xlsx"
    make_smart_excel(input_file, output_file)
