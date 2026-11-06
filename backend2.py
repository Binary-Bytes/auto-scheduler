"""
Pure backend logic for the Night School Weekday Schedule Auto-Filler.

Unlike the room-schedule tool, there is no pre-built template grid to match
against here: each weekday ("Monday", "Tuesday", ...) gets its own sheet,
and every course gets its own ROW within the sheet(s) for the weekday(s) it
meets on. Weeks are single columns (one column per calendar occurrence of
that weekday), grouped under a month header, and a course's active weeks
are shaded light gray rather than filled with text.
"""
import re
from datetime import datetime, timedelta

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

WEEKDAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
CODE_TO_WEEKDAY = {'M': 0, 'T': 1, 'W': 2, 'R': 3, 'F': 4, 'S': 5}  # anything else -> 6 (Sunday)

HEADER_COLS = ["Course", "Course #", "Room", "Instructor", "Start", "End", "Time", "NOTE"]
FIRST_WEEK_COL = 9   # column I - first week-of-term column, matching the real file's layout
TITLE_ROW = 1
HEADER_ROW = 2
DATE_ROW = 3
FIRST_COURSE_ROW = 4


# ----------------------------------------------------------------------------
# THEMING  (every look-and-feel choice lives in a Theme object that is passed
# explicitly into the builders - no module-level state, so concurrent Streamlit
# sessions can never affect each other's colours)
# ----------------------------------------------------------------------------
WEEK_COL_WIDTH = 5.0   # near-square columns for the week-date grid (fits "S/E")

DEFAULT_ACCENT = "2F5597"

# name -> accent colour. The default keeps its hand-tuned palette (see make_theme).
PRESET_ACCENTS = {
    "Classic Blue (default)": DEFAULT_ACCENT,
    "Forest Green": "2E6B4A",
    "Teal": "1F7A7A",
    "Burgundy": "8B2635",
    "Royal Purple": "5B3F8C",
    "Burnt Orange": "B45309",
    "Slate Grey": "475569",
}
FONT_CHOICES = ["Calibri", "Arial", "Cambria", "Tahoma"]
BORDER_CHOICES = ["Standard", "Bold"]
DENSITY_CHOICES = ["Compact", "Standard", "Spacious"]
_BODY_HEIGHTS = {"Compact": 27, "Standard": 30, "Spacious": 38}

# Hand-tuned palette that reproduces the original look exactly
_DEFAULT_PALETTE = {
    "dark": "1F3864", "accent": "2F5597", "month": "D9E2F3",
    "date": "EEF2F9", "shade": "BDD0EC", "zebra": "F5F7FB",
}


def _hex_to_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _rgb_to_hex(rgb):
    return "".join(f"{max(0, min(255, round(v))):02X}" for v in rgb)


def _mix(hex_color, other_rgb, t):
    """Blend hex_color toward other_rgb by fraction t (0 = unchanged, 1 = fully other)."""
    r, g, b = _hex_to_rgb(hex_color)
    return _rgb_to_hex(tuple(c + (o - c) * t for c, o in zip((r, g, b), other_rgb)))


def _luminance(hex_color):
    def lin(c):
        c /= 255
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = (lin(v) for v in _hex_to_rgb(hex_color))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def _contrast_with_white(hex_color):
    return 1.05 / (_luminance(hex_color) + 0.05)


def normalize_accent(accent):
    """Validates a hex colour; darkens it until white text on it is readable (WCAG AA)."""
    h = (accent or DEFAULT_ACCENT).strip().lstrip("#")
    if len(h) != 6 or any(ch not in "0123456789abcdefABCDEF" for ch in h):
        h = DEFAULT_ACCENT
    h = h.upper()
    guard = 0
    while _contrast_with_white(h) < 4.5 and guard < 40:
        h = _mix(h, (0, 0, 0), 0.08)
        guard += 1
    return h


def palette_for_accent(accent):
    accent = normalize_accent(accent)
    if accent == DEFAULT_ACCENT:
        return dict(_DEFAULT_PALETTE)
    white, black = (255, 255, 255), (0, 0, 0)
    return {
        "accent": accent,
        "dark": _mix(accent, black, 0.33),
        "month": _mix(accent, white, 0.82),
        "date": _mix(accent, white, 0.92),
        "shade": _mix(accent, white, 0.70),
        "zebra": _mix(accent, white, 0.96),
    }


class Theme:
    """All fonts / fills / borders / sizes used by the workbook builder."""

    def __init__(self, accent=DEFAULT_ACCENT, font="Calibri", border="Standard",
                 density="Standard", zebra=True):
        self.palette = p = palette_for_accent(accent)
        self.zebra_on = bool(zebra)
        font = font if font in FONT_CHOICES else "Calibri"
        ink = "262626"

        self.title_font = Font(name=font, bold=True, size=16, color="FFFFFF")
        self.header_font = Font(name=font, bold=True, size=12, color="FFFFFF")
        self.month_font = Font(name=font, bold=True, size=11, color=p["dark"])
        self.date_font = Font(name=font, bold=True, size=10, color=p["dark"])
        body_size = 10 if density == "Compact" else 11   # keeps two-line course names fully visible
        self.body_font = Font(name=font, size=body_size, color=ink)
        self.course_font = Font(name=font, size=body_size, bold=True, color=ink)
        self.mark_font = Font(name=font, size=10, bold=True, color="FFFFFF")
        self.legend_font = Font(name=font, italic=True, size=10, color="595959")

        solid = lambda c: PatternFill(patternType="solid", fgColor=c)
        self.title_fill = solid(p["dark"])
        self.header_fill = solid(p["accent"])
        self.month_fill = solid(p["month"])
        self.date_fill = solid(p["date"])
        self.zebra_fill = solid(p["zebra"])
        self.mark_fill = solid(p["accent"])
        self.shade_fill = solid(p["shade"])

        self.thin_side = Side(style="thin", color="A6A6A6")
        self.outer_side = Side(style="thick" if border == "Bold" else "medium", color=p["dark"])

        self.title_row_height = 32
        self.header_row_height = 24
        self.date_row_height = 18
        self.body_row_height = _BODY_HEIGHTS.get(density, 30)


DEFAULT_THEME = Theme()

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)


def weekday_for_code(ch):
    return CODE_TO_WEEKDAY.get(ch.upper(), 6)


def weekdays_for_days_string(days_str):
    """Returns the set of weekday indices (0=Mon..6=Sun) a Days code string covers."""
    return {weekday_for_code(ch) for ch in days_str if not ch.isspace()}


# ----------------------------------------------------------------------------
# DATA SHEET LOADING (same conventions as the room-schedule tool)
# ----------------------------------------------------------------------------
def parse_data_date(value, base_year, fall_start_month=7):
    if isinstance(value, datetime):
        return value.date()
    if hasattr(value, 'year') and hasattr(value, 'month'):
        return value
    s = str(value).strip()
    m = re.match(r'^(\d{1,2})[-\s]([A-Za-z]{3,})$', s)
    if not m:
        raise ValueError(f"Could not parse date value: {value!r}")
    day = int(m.group(1))
    mon_str = m.group(2)[:3].title()
    month = datetime.strptime(mon_str, "%b").month
    year = base_year if month >= fall_start_month else base_year + 1
    return datetime(year, month, day).date()


def load_data_rows(ws_data):
    headers = {}
    for c in range(1, ws_data.max_column + 1):
        v = ws_data.cell(row=1, column=c).value
        if v:
            headers[str(v).strip().lower()] = c

    required = ['starts', 'ends', 'course', 'name', 'days', 'instructor']
    for req in required:
        if req not in headers:
            raise ValueError(f"Data sheet is missing required column: '{req}'")

    first_start_raw = None
    for r in range(2, ws_data.max_row + 1):
        v = ws_data.cell(row=r, column=headers['starts']).value
        if v is not None:
            first_start_raw = v
            break
    if first_start_raw is None:
        raise ValueError("Could not find any Start dates in the Data sheet.")
    if hasattr(first_start_raw, 'year'):
        base_year = first_start_raw.year
    else:
        base_year = datetime.now().year

    rows = []
    for r in range(2, ws_data.max_row + 1):
        name_val = ws_data.cell(row=r, column=headers['name']).value
        if name_val is None or str(name_val).strip() == "":
            continue

        start_raw = ws_data.cell(row=r, column=headers['starts']).value
        end_raw = ws_data.cell(row=r, column=headers['ends']).value
        start_date = parse_data_date(start_raw, base_year)
        end_date = parse_data_date(end_raw, base_year)

        course_raw = str(ws_data.cell(row=r, column=headers['course']).value or "").strip()
        days_raw = str(ws_data.cell(row=r, column=headers['days']).value or "").strip()
        instructor = str(ws_data.cell(row=r, column=headers['instructor']).value or "").strip()
        name = str(name_val).strip()

        notes = ""
        if 'notes' in headers:
            notes = str(ws_data.cell(row=r, column=headers['notes']).value or "").strip()

        rows.append({
            'row': r,
            'start': start_date,
            'end': end_date,
            'course_raw': course_raw,
            'name': name,
            'days': days_raw,
            'instructor': instructor,
            'notes': notes,
        })
    return rows


def required_weekdays(data_rows):
    """Which of the 7 weekdays have at least one course meeting on them."""
    needed = set()
    for course in data_rows:
        needed |= weekdays_for_days_string(course['days'])
    return sorted(needed)  # 0=Mon..6=Sun


# ----------------------------------------------------------------------------
# YEAR OVERRIDE
# ----------------------------------------------------------------------------
def detect_start_year(data_rows):
    """Year of the earliest course start date found in the Data sheet."""
    return min(c['start'] for c in data_rows).year


def _shift_date_years(d, delta):
    try:
        return d.replace(year=d.year + delta)
    except ValueError:  # Feb 29 landing on a non-leap year
        return d.replace(year=d.year + delta, day=28)


def shift_rows_to_year(data_rows, target_year):
    """
    Moves every course's start/end date so the EARLIEST start falls in
    `target_year`. All dates move by the same number of years, so terms that
    cross New Year (e.g. Sep -> Jan) keep their shape. Returns new row dicts.
    """
    delta = target_year - detect_start_year(data_rows)
    if delta == 0:
        return data_rows
    return [
        {**c,
         'start': _shift_date_years(c['start'], delta),
         'end': _shift_date_years(c['end'], delta)}
        for c in data_rows
    ]


def suggest_term_label(data_rows):
    """FALL for terms starting Jul-Dec, otherwise SPRING (Jan-Jun)."""
    return "FALL" if min(c['start'] for c in data_rows).month >= 7 else "SPRING"


# ----------------------------------------------------------------------------
# WEEK-COLUMN CALCULATION (shared term range across all weekday sheets)
# ----------------------------------------------------------------------------
def term_date_range(data_rows):
    """
    Returns the expanded date range stretching from the 1st day of the earliest 
    month to the final day of the latest month across all courses.
    """
    if not data_rows:
        return None, None
    min_start = min(c['start'] for c in data_rows)
    max_end = max(c['end'] for c in data_rows)
    
    # Expand to the 1st day of the earliest start month
    start_bound = min_start.replace(day=1)
    
    # Expand to the last day of the latest end month
    if max_end.month == 12:
        end_bound = max_end.replace(year=max_end.year + 1, month=1, day=1) - timedelta(days=1)
    else:
        end_bound = max_end.replace(month=max_end.month + 1, day=1) - timedelta(days=1)
        
    return start_bound, end_bound


def week_dates_for_weekday(weekday_idx, term_min_start, term_max_end):
    if term_min_start is None or term_max_end is None:
        return []

    # first occurrence of this weekday on/after term_min_start boundary
    delta = (weekday_idx - term_min_start.weekday()) % 7
    d = term_min_start + timedelta(days=delta)
    dates = []
    while d <= term_max_end:
        dates.append(d)
        d += timedelta(days=7)
    return dates


def group_dates_by_month(dates):
    """Returns list of (month_label, [dates]) preserving order, for the header row."""
    groups = []
    for d in dates:
        label = d.strftime("%B").upper()
        if groups and groups[-1][0] == label:
            groups[-1][1].append(d)
        else:
            groups.append((label, [d]))
    return groups



# ----------------------------------------------------------------------------
# WORKBOOK CONSTRUCTION
# ----------------------------------------------------------------------------
def build_sheet_for_weekday(wb, weekday_idx, dates, school_name="Night School", term_label="FALL",
                            theme=DEFAULT_THEME):
    weekday_name = WEEKDAY_NAMES[weekday_idx]
    ws = wb.create_sheet(title=weekday_name)
    ws.sheet_properties.tabColor = theme.palette["dark"]

    year_label = dates[0].year if dates else datetime.now().year
    last_col = FIRST_WEEK_COL + max(len(dates) - 1, 0)
    full_last_col = max(last_col, 8)

    # Title band
    title_cell = ws.cell(row=TITLE_ROW, column=1,
                         value=f"{school_name} \u2013 Room Schedule \u2013 {term_label} {year_label} \u2013 {weekday_name}")
    ws.merge_cells(start_row=TITLE_ROW, start_column=1, end_row=TITLE_ROW, end_column=full_last_col)
    for c in range(1, full_last_col + 1):
        ws.cell(row=TITLE_ROW, column=c).fill = theme.title_fill
    title_cell.font = theme.title_font
    title_cell.alignment = CENTER
    ws.row_dimensions[TITLE_ROW].height = theme.title_row_height

    # Fixed columns: header text spans the month row + date row
    for i, label in enumerate(HEADER_COLS, start=1):
        cell = ws.cell(row=HEADER_ROW, column=i, value=label)
        ws.cell(row=DATE_ROW, column=i).fill = theme.header_fill
        cell.font = theme.header_font
        cell.alignment = CENTER
        cell.fill = theme.header_fill
        ws.merge_cells(start_row=HEADER_ROW, start_column=i, end_row=DATE_ROW, end_column=i)

    # Month header + date row
    col = FIRST_WEEK_COL
    for month_label, month_dates in group_dates_by_month(dates):
        start_col = col
        for d in month_dates:
            dcell = ws.cell(row=DATE_ROW, column=col, value=d.day)
            dcell.font = theme.date_font
            dcell.alignment = CENTER
            dcell.fill = theme.date_fill
            ws.cell(row=HEADER_ROW, column=col).fill = theme.month_fill
            col += 1
        end_col = col - 1
        cell = ws.cell(row=HEADER_ROW, column=start_col, value=month_label)
        cell.font = theme.month_font
        cell.alignment = CENTER
        cell.fill = theme.month_fill
        if end_col > start_col:
            ws.merge_cells(start_row=HEADER_ROW, start_column=start_col, end_row=HEADER_ROW, end_column=end_col)

    ws.row_dimensions[HEADER_ROW].height = theme.header_row_height
    ws.row_dimensions[DATE_ROW].height = theme.date_row_height
    return ws


def first_and_last_class_dates(course):
    """
    The course's real first and last class days: the earliest / latest date inside
    [start, end] that falls on one of the weekdays the course meets (so a Mon-Thu
    course starting on Mon 4 Jan has its first class on 4 Jan only, not on every
    weekday sheet).
    """
    wds = weekdays_for_days_string(course['days'])
    first = last = None
    d = course['start']
    while d <= course['end']:
        if d.weekday() in wds:
            if first is None:
                first = d
            last = d
        d += timedelta(days=1)
    return first, last


def add_course_row(ws, row_idx, course, room, time_str, dates_for_weekday, weekday_idx, theme=DEFAULT_THEME):
    """
    Writes one course row. Week cells inside the course's date range are shaded;
    "S" marks the course's very first class day and "E" its very last (or "S/E"
    if both are the same day) - each appears on only one weekday sheet.
    """
    days_suffix = f" ({course['days']})" if course['days'] else ""

    values = [
        course['name'] + days_suffix,
        course['course_raw'],
        room,
        course['instructor'],
        datetime(course['start'].year, course['start'].month, course['start'].day),
        datetime(course['end'].year, course['end'].month, course['end'].day),
        time_str,
        course['notes'],
    ]
    zebra = theme.zebra_fill if (theme.zebra_on and (row_idx - FIRST_COURSE_ROW) % 2 == 1) else None
    for col_idx, val in enumerate(values, start=1):
        cell = ws.cell(row=row_idx, column=col_idx, value=val if val != "" else None)
        cell.font = theme.course_font if col_idx == 1 else theme.body_font
        cell.alignment = LEFT_WRAP if col_idx in (1, 8) else CENTER
        if zebra:
            cell.fill = zebra

    ws.cell(row=row_idx, column=5).number_format = "d-mmm-yyyy"
    ws.cell(row=row_idx, column=6).number_format = "d-mmm-yyyy"
    ws.row_dimensions[row_idx].height = theme.body_row_height

    active = [i for i, d in enumerate(dates_for_weekday) if course['start'] <= d <= course['end']]
    for i in range(len(dates_for_weekday)):
        wcell = ws.cell(row=row_idx, column=FIRST_WEEK_COL + i)
        if i in active:
            wcell.fill = theme.shade_fill
        elif zebra:
            wcell.fill = zebra

    first_day, last_day = first_and_last_class_dates(course)
    for i, d in enumerate(dates_for_weekday):
        is_first = d == first_day
        is_last = d == last_day
        if not (is_first or is_last):
            continue
        text = "S/E" if (is_first and is_last) else ("S" if is_first else "E")
        mcell = ws.cell(row=row_idx, column=FIRST_WEEK_COL + i, value=text)
        mcell.font = theme.mark_font
        mcell.fill = theme.mark_fill
        mcell.alignment = CENTER
    return len(active)


def add_legend(ws, row_idx, theme=DEFAULT_THEME):
    cell = ws.cell(row=row_idx, column=1,
                   value="S = first class day     E = last class day     Shaded = class running that week")
    cell.font = theme.legend_font
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)


def apply_table_borders(ws, last_row, last_col, month_end_cols, theme=DEFAULT_THEME):
    """Thin grey grid inside, heavier themed frame + dividers (header base, fixed/week split, months)."""
    THIN, OUTER = theme.thin_side, theme.outer_side
    for r in range(HEADER_ROW, last_row + 1):
        for c in range(1, last_col + 1):
            left = OUTER if c == 1 else THIN
            right = OUTER if (c == last_col or c == FIRST_WEEK_COL - 1 or c in month_end_cols) else THIN
            top = OUTER if r == HEADER_ROW else THIN
            bottom = OUTER if (r == DATE_ROW or r == last_row) else THIN
            ws.cell(row=r, column=c).border = Border(left=left, right=right, top=top, bottom=bottom)


def enable_wrap_text_everywhere(wb):
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                existing = cell.alignment
                cell.alignment = Alignment(
                    horizontal=existing.horizontal,
                    vertical=existing.vertical or "center",
                    wrap_text=True if cell.row != ws.max_row else existing.wrap_text,
                    text_rotation=existing.text_rotation,
                    indent=existing.indent,
                )


def style_columns(ws, last_week_col):
    fixed_widths = {1: 30, 2: 12, 3: 9, 4: 16, 5: 13, 6: 13, 7: 12, 8: 24}
    for col, width in fixed_widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width
    for col in range(FIRST_WEEK_COL, last_week_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = WEEK_COL_WIDTH


def setup_print_and_view(ws):
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = ws.cell(row=FIRST_COURSE_ROW, column=2)   # keep headers + course name visible
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_title_rows = f"{TITLE_ROW}:{DATE_ROW}"
    ws.page_margins.left = ws.page_margins.right = 0.4


def build_workbook(data_rows, assignments, school_name="Night School", term_label="FALL", theme=None):
    theme = theme or DEFAULT_THEME
    wb = Workbook()
    wb.remove(wb.active)

    needed_weekdays = required_weekdays(data_rows)
    placed_counts = {}

    term_min_start, term_max_end = term_date_range(data_rows)

    for weekday_idx in needed_weekdays:
        dates = week_dates_for_weekday(weekday_idx, term_min_start, term_max_end)
        ws = build_sheet_for_weekday(wb, weekday_idx, dates, school_name=school_name,
                                     term_label=term_label, theme=theme)

        row_idx = FIRST_COURSE_ROW
        for course in data_rows:
            if weekday_idx not in weekdays_for_days_string(course['days']):
                continue
            assignment = assignments.get(course['row'], {})
            room = assignment.get('room', '') or ''
            time_str = assignment.get('time', '') or ''
            placed = add_course_row(ws, row_idx, course, room, time_str, dates, weekday_idx, theme)
            placed_counts[(course['row'], weekday_idx)] = placed
            row_idx += 1

        last_table_row = row_idx - 1
        last_week_col = FIRST_WEEK_COL + max(len(dates) - 1, 0)
        last_col = max(last_week_col, 8)

        # Column index of the last week of each month (except the final month) -> heavier divider
        month_end_cols, curr_col = set(), FIRST_WEEK_COL
        for _, month_dates in group_dates_by_month(dates)[:-1]:
            curr_col += len(month_dates)
            month_end_cols.add(curr_col - 1)

        apply_table_borders(ws, last_table_row, last_col, month_end_cols, theme)
        add_legend(ws, last_table_row + 2, theme)
        style_columns(ws, last_week_col)
        setup_print_and_view(ws)

    enable_wrap_text_everywhere(wb)
    return wb, placed_counts
