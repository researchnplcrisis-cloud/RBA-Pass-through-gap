"""Pass-through gap build, October 2026 refresh.

Same method as the original pass-through-gap build (F1.1 / F5 / F4), extended with
RBA F6 actual housing lending rates (existing and new owner-occupier variable loans).
HTML charts are rendered to PNG with headless Chrome (or Edge) so both formats match.
"""

import csv
import datetime as dt
import html
import json
import math
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


PROJECT = Path(__file__).resolve().parents[1]
RAW = PROJECT / "data" / "raw"
OUT = PROJECT / "output"
WORK = PROJECT / "work"
ANALYST_NAME = "Irine"
NEXT_DECISION = "3 November"  # RBA Monetary Policy Board meets 2-3 Nov 2026

NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

FILES = {
    "F1.1": RAW / "f01hist.xlsx",
    "F5": RAW / "f05hist.xlsx",
    "F4": RAW / "f04hist.xlsx",
    "F6": RAW / "f06hist.xlsx",
}
DAILY_CASH = RAW / "f01d.xlsx"  # F1 daily: dates the actual target changes

SERIES = {
    "cash_rate": {
        "table": "F1.1",
        "contains": ["Cash Rate Target"],
    },
    "mortgage_variable": {
        "table": "F5",
        "contains": [
            "Lending rates; Housing loans; Banks; Variable; Standard; Owner-occupier"
        ],
    },
    "mortgage_discounted": {
        "table": "F5",
        "contains": [
            "Lending rates; Housing loans; Banks; Variable; Discounted; Owner-occupier"
        ],
    },
    "mortgage_outstanding": {
        "table": "F6",
        "contains": [
            "Lending rates; Housing credit; Outstanding; Owner-occupied; Variable-rate; All institutions"
        ],
    },
    "mortgage_new": {
        "table": "F6",
        "contains": [
            "Lending rates; Housing credit; New loans funded in the month; Owner-occupied; Variable-rate; All institutions"
        ],
    },
    "saver_online": {
        "table": "F4",
        "contains": [
            "Retail deposit and investment rates; Savings accounts; Online savings accounts; $10000"
        ],
    },
    "term_deposit": {
        "table": "F4",
        "contains": [
            "Retail deposit and investment rates; Banks' term deposits ($10000); 1 year"
        ],
    },
}

CORE_COLS = ["cash_rate", "mortgage_variable", "saver_online", "term_deposit"]
F6_COLS = ["mortgage_outstanding", "mortgage_new"]
MERGED_COLS = CORE_COLS + F6_COLS
RETAIL_COLS = ["mortgage_variable", "mortgage_outstanding", "mortgage_new", "saver_online", "term_deposit"]

COLORS = {
    "navy": "#173B63",
    "red": "#8B2D2B",
    "teal": "#167C78",
    "amber": "#B9871E",
    "plum": "#6A3D9A",
    "mauve": "#B04F8F",
    "grey": "#E9ECEF",
    "dark": "#18212B",
    "mid": "#5E6A75",
    "axis": "#CCD3DA",
    "grid": "#E8ECF0",
}

STYLE = {
    "cash_rate": {"color": COLORS["navy"], "legend": "Cash rate target", "short": "Cash rate", "step": True, "dash": None},
    "mortgage_variable": {"color": COLORS["red"], "legend": "Standard variable rate (advertised)", "short": "Standard variable", "step": False, "dash": None},
    "mortgage_outstanding": {"color": COLORS["plum"], "legend": "Existing variable loans (actual)", "short": "Existing loans", "step": False, "dash": None},
    "mortgage_new": {"color": COLORS["mauve"], "legend": "New variable loans (actual)", "short": "New loans", "step": False, "dash": "10 7"},
    "saver_online": {"color": COLORS["teal"], "legend": "Online saver", "short": "Online saver", "step": False, "dash": None},
    "term_deposit": {"color": COLORS["amber"], "legend": "1-year term deposit", "short": "1y term deposit", "step": False, "dash": None},
}

STAT_LABEL = {
    "mortgage_variable": "Standard variable mortgage (F5, advertised)",
    "mortgage_outstanding": "Existing variable owner-occupier loans (F6, actual)",
    "mortgage_new": "New variable owner-occupier loans (F6, actual)",
    "saver_online": "Online saver",
    "term_deposit": "1-year term deposit",
}

SPREADS = {
    "advertised": ("mortgage_variable", "Advertised wedge: standard variable minus online saver"),
    "actual": ("mortgage_outstanding", "Actual wedge: existing variable loans minus online saver"),
}


# ---------------------------------------------------------------- xlsx parsing

def col_to_idx(cell_ref):
    match = re.match(r"([A-Z]+)", cell_ref or "")
    if not match:
        return 0
    value = 0
    for char in match.group(1):
        value = value * 26 + (ord(char) - ord("A") + 1)
    return value - 1


def load_shared_strings(zf):
    try:
        root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    strings = []
    for item in root.findall("a:si", NS):
        strings.append("".join(text.text or "" for text in item.findall(".//a:t", NS)))
    return strings


def first_sheet_path(zf):
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
    sheet = workbook.find("a:sheets/a:sheet", NS)
    rid = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
    target = rel_map[rid]
    if target.startswith("/"):
        return target.lstrip("/")
    if target.startswith("xl/"):
        return target
    return "xl/" + target


def cell_value(cell, shared_strings):
    ctype = cell.attrib.get("t")
    if ctype == "s":
        value = cell.findtext("a:v", default="", namespaces=NS)
        return shared_strings[int(value)] if value else ""
    if ctype == "inlineStr":
        return "".join(text.text or "" for text in cell.findall(".//a:t", NS))
    return cell.findtext("a:v", default="", namespaces=NS)


def read_xlsx_rows(path):
    with zipfile.ZipFile(path) as zf:
        shared_strings = load_shared_strings(zf)
        sheet_path = first_sheet_path(zf)
        root = ET.fromstring(zf.read(sheet_path))
        rows = []
        for row in root.findall(".//a:sheetData/a:row", NS):
            values = {}
            for cell in row.findall("a:c", NS):
                values[col_to_idx(cell.attrib.get("r", ""))] = cell_value(cell, shared_strings)
            if values:
                rows.append([values.get(idx, "") for idx in range(max(values) + 1)])
            else:
                rows.append([])
        return rows


def excel_date(value):
    # Excel's serial 1 is 1900-01-01, with the 1900 leap-year bug baked in.
    return dt.date(1899, 12, 30) + dt.timedelta(days=float(value))


def month_date(value):
    date = excel_date(value)
    return dt.date(date.year, date.month, 1)


def parse_float(value):
    if value is None or value == "":
        return None
    try:
        return float(value)
    except ValueError:
        return None


def norm(text):
    return re.sub(r"\s+", " ", text.replace("\uFFFD", " ").replace("\xa0", " ")).strip().lower()


def format_row(row):
    return "[" + ", ".join(repr(value) for value in row) + "]"


def find_col(rows, wanted):
    normalized = [norm(value) for value in rows[2]]
    for target in wanted:
        target_norm = norm(target)
        for idx, desc in enumerate(normalized):
            if target_norm in desc:
                return idx
    raise ValueError(f"Could not find any description matching {wanted!r}")


def parse_table(table_name, path):
    rows = read_xlsx_rows(path)
    description = rows[2] if len(rows) > 2 else []
    series_ids = rows[10] if len(rows) > 10 else []
    metadata = []
    for idx in range(1, max(len(description), len(series_ids))):
        desc = description[idx] if idx < len(description) else ""
        sid = series_ids[idx] if idx < len(series_ids) else ""
        if desc or sid:
            metadata.append((idx, desc, sid))
    data = []
    for row in rows[11:]:
        if not row or not row[0]:
            continue
        try:
            date = month_date(row[0])
        except Exception:
            continue
        values = {idx: parse_float(row[idx]) if idx < len(row) else None for idx in range(1, len(description))}
        data.append((date, values))
    return {"name": table_name, "path": path, "rows": rows, "metadata": metadata, "data": data}


def cash_target_decisions(path):
    """Effective dates of cash-rate target changes from F1 daily: [(date, old, new)]."""
    rows = read_xlsx_rows(path)
    col = find_col(rows, ["Cash Rate Target"])
    decisions = []
    prev = None
    for row in rows[11:]:
        if not row or not row[0] or col >= len(row):
            continue
        value = parse_float(row[col])
        if value is None:
            continue
        value = round(value, 4)
        if prev is not None and value != prev:
            decisions.append((excel_date(row[0]), prev, value))
        prev = value
    return decisions


def decisions_in(decisions, start_month, end_month):
    return [d for d in decisions if start_month <= dt.date(d[0].year, d[0].month, 1) <= end_month]


def series_values(table, col_idx):
    return [(date, row_values.get(col_idx)) for date, row_values in table["data"]]


def non_missing(values):
    return [(date, value) for date, value in values if value is not None]


# ---------------------------------------------------------------- date / format helpers

def date_label(date):
    return date.strftime("%b %Y")


def iso_month(date):
    return date.strftime("%Y-%m")


def add_months(date, months):
    month = date.month - 1 + months
    return dt.date(date.year + month // 12, month % 12 + 1, 1)


def previous_month(date):
    return add_months(date, -1)


def month_range(start, end):
    current = start
    while current <= end:
        yield current
        current = add_months(current, 1)


def fmt(value, digits=2):
    return "n/a" if value is None else f"{value:.{digits}f}"


def fmt_pct(value):
    return "n/a" if value is None else f"{value * 100:.1f}%"


# ---------------------------------------------------------------- statistics

def lookup(merged, date, col):
    row = merged.get(date)
    return row.get(col) if row else None


def delta(merged, start_date, end_date, col):
    start = lookup(merged, start_date, col)
    end = lookup(merged, end_date, col)
    if start is None or end is None:
        return None
    return end - start


def ratio(merged, start_date, end_date, retail_col):
    cash_delta = delta(merged, start_date, end_date, "cash_rate")
    retail_delta = delta(merged, start_date, end_date, retail_col)
    if cash_delta is None or retail_delta is None or abs(cash_delta) < 1e-12:
        return None
    return retail_delta / cash_delta


def latest_on_or_before(merged, col, end):
    for date in sorted((d for d in merged if d <= end), reverse=True):
        if lookup(merged, date, col) is not None:
            return date
    return None


def latest_complete_date(merged, cols):
    complete = [date for date, row in merged.items() if all(row.get(col) is not None for col in cols)]
    return max(complete) if complete else None


def spread(merged, date, mortgage_col):
    mortgage = lookup(merged, date, mortgage_col)
    saver = lookup(merged, date, "saver_online")
    if date is None or mortgage is None or saver is None:
        return None
    return mortgage - saver


def detect_2026_episode(merged, latest_date):
    for date in sorted(d for d in merged if d.year == 2026 and d <= latest_date):
        current_cash = lookup(merged, date, "cash_rate")
        prior_cash = lookup(merged, previous_month(date), "cash_rate")
        if current_cash is not None and prior_cash is not None and current_cash > prior_cash:
            return date, latest_date
    return None, None


def episode_stats(merged, decisions, label, start_hike, end_date):
    base_date = previous_month(start_hike)
    hikes = []
    for date in month_range(start_hike, end_date):
        c0 = lookup(merged, previous_month(date), "cash_rate")
        c1 = lookup(merged, date, "cash_rate")
        if c0 is not None and c1 is not None and c1 > c0:
            hikes.append(date)
    result = {
        "label": label,
        "start_hike": start_hike,
        "base_date": base_date,
        "end_date": end_date,
        "cash_change": delta(merged, base_date, end_date, "cash_rate"),
        "hikes": [],
        "full": {},
        "full_end": {},
        "full_cash_change": {},
        "spreads": {},
        "decisions": decisions_in(decisions, start_hike, end_date),
    }
    for col in RETAIL_COLS:
        # F6 lags the other tables by a month, so each series is measured to its own latest month.
        col_end = latest_on_or_before(merged, col, end_date)
        if col_end is None or col_end <= base_date:
            col_end = None
        result["full_end"][col] = col_end
        result["full"][col] = ratio(merged, base_date, col_end, col) if col_end else None
        result["full_cash_change"][col] = delta(merged, base_date, col_end, "cash_rate") if col_end else None
    for key, (mortgage_col, _) in SPREADS.items():
        s_end = latest_on_or_before(merged, mortgage_col, end_date)
        saver_end = latest_on_or_before(merged, "saver_online", end_date)
        if s_end and saver_end:
            s_end = min(s_end, saver_end)
        start = spread(merged, base_date, mortgage_col)
        end = spread(merged, s_end, mortgage_col)
        result["spreads"][key] = {
            "start": start,
            "end": end,
            "end_date": s_end,
            "change": None if start is None or end is None else end - start,
        }
    for date in hikes:
        prior = previous_month(date)
        hike_info = {
            "date": date,
            "cash_change": delta(merged, prior, date, "cash_rate"),
            "decisions": decisions_in(decisions, date, date),
            "lag_1m": {},
            "lag_3m": {},
        }
        for col in RETAIL_COLS:
            hike_info["lag_1m"][col] = ratio(merged, prior, add_months(date, 1), col)
            hike_info["lag_3m"][col] = ratio(merged, prior, add_months(date, 3), col)
        result["hikes"].append(hike_info)
    return result


# ---------------------------------------------------------------- SVG primitives

def linear_scale(domain_min, domain_max, range_min, range_max):
    if domain_max == domain_min:
        return lambda _: (range_min + range_max) / 2
    return lambda value: range_min + (value - domain_min) * (range_max - range_min) / (domain_max - domain_min)


def date_to_num(date):
    return date.toordinal()


def nice_ticks(low, high, count=6):
    if high <= low:
        return [low]
    step = (high - low) / max(1, count - 1)
    magnitude = 10 ** math.floor(math.log10(step))
    residual = step / magnitude
    if residual >= 5:
        nice_step = 5 * magnitude
    elif residual >= 2:
        nice_step = 2 * magnitude
    else:
        nice_step = magnitude
    ticks = []
    value = math.floor(low / nice_step) * nice_step
    while value <= high + 1e-9:
        if value >= low - 1e-9:
            ticks.append(round(value, 10))
        value += nice_step
    return ticks


def text_width(text, size, bold=False):
    return len(text) * size * (0.58 if bold else 0.52)


def svg_text(x, y, value, size=24, weight=400, fill=None, anchor="start", halo=False):
    fill = fill or COLORS["dark"]
    halo_attr = ' stroke="white" stroke-width="6" stroke-linejoin="round" paint-order="stroke"' if halo else ""
    return (
        f'<text x="{x:.1f}" y="{y:.1f}" font-size="{size}" font-weight="{weight}" '
        f'fill="{fill}" text-anchor="{anchor}"{halo_attr}>{html.escape(str(value))}</text>'
    )


def wrap_svg_text(text, max_chars):
    lines, current = [], []
    for word in text.split():
        trial = " ".join(current + [word])
        if current and len(trial) > max_chars:
            lines.append(" ".join(current))
            current = [word]
        else:
            current.append(word)
    if current:
        lines.append(" ".join(current))
    return lines


def line_path(points, step=False):
    commands = []
    pen_down = False
    prev_y = None
    for x, y in points:
        if x is None or y is None:
            pen_down = False  # gaps are left as gaps
            continue
        if not pen_down:
            commands.append(f"M {x:.2f} {y:.2f}")
            pen_down = True
        elif step:
            commands.append(f"L {x:.2f} {prev_y:.2f} L {x:.2f} {y:.2f}")
        else:
            commands.append(f"L {x:.2f} {y:.2f}")
        prev_y = y
    return " ".join(commands)


def is_wide(width):
    return width >= 1500


def header_svg(width, title, subtitle):
    tsize = 44 if is_wide(width) else 38
    ssize = 25 if is_wide(width) else 23
    parts = []
    y = 88
    for line in wrap_svg_text(title, int((width - 140) / (tsize * 0.56))):
        parts.append(svg_text(70, y, line, tsize, 700))
        y += tsize * 1.18
    y += ssize * 1.45 - tsize * 1.18
    for line in wrap_svg_text(subtitle, int((width - 140) / (ssize * 0.5))):
        parts.append(svg_text(70, y, line, ssize, 400, COLORS["mid"]))
        y += ssize * 1.35
    return parts, y - ssize * 1.35 + 12


def footer_svg(width, height, source, notes):
    size = 20 if is_wide(width) else 18
    lines = []
    for item in notes + [source]:
        lines.extend(wrap_svg_text(item, int((width - 140) / (size * 0.5))))
    parts = []
    y = height - 44
    for line in reversed(lines):
        parts.append(svg_text(70, y, line, size, 400, COLORS["mid"]))
        y -= size * 1.45
    return parts, y + size * 1.45 - size


def legend_svg(items, x0, y0, max_x, size=19):
    """items: (kind, label, color, dash) with kind in line / rect / band. Wraps onto new rows."""
    parts = []
    x, y = x0, y0
    for kind, label, color, dash in items:
        w = 46 + text_width(label, size, True) + 34
        if x > x0 and x + w > max_x:
            x = x0
            y += size * 1.9
        if kind == "line":
            dash_attr = f' stroke-dasharray="{dash}"' if dash else ""
            parts.append(f'<line x1="{x}" y1="{y - 6:.1f}" x2="{x + 34}" y2="{y - 6:.1f}" stroke="{color}" stroke-width="5" stroke-linecap="round"{dash_attr}/>')
        elif kind == "rect":
            parts.append(f'<rect x="{x}" y="{y - 18:.1f}" width="30" height="20" rx="3" fill="{color}"/>')
        else:
            parts.append(f'<rect x="{x}" y="{y - 18:.1f}" width="34" height="20" fill="{COLORS["grey"]}" stroke="{COLORS["axis"]}" stroke-width="1"/>')
        weight = 400 if kind == "band" else 600
        fill = COLORS["mid"] if kind == "band" else COLORS["dark"]
        parts.append(svg_text(x + 46, y, label, size, weight, fill))
        x += w
    return parts, y


def page(width, height, title, parts):
    return f"""<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>{html.escape(title)}</title>
<style>
html, body {{ margin: 0; padding: 0; background: white; }}
svg {{ display: block; width: {width}px; height: {height}px; font-family: Arial, Helvetica, sans-serif; }}
</style>
</head>
<body>
<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}">
<rect width="{width}" height="{height}" fill="white"/>
{chr(10).join(parts)}
</svg>
</body>
</html>"""


def source_line(tables):
    return f"Source: RBA Statistical Tables {', '.join(tables)}. Analysis: {ANALYST_NAME}."


def shaded_bands(episodes, x_scale, plot_top, plot_height, latest_date):
    bands = []
    for _, start_hike, end in episodes:
        x0 = x_scale(date_to_num(previous_month(start_hike)))
        x1 = x_scale(date_to_num(min(end, latest_date)))
        bands.append(
            f'<rect x="{x0:.1f}" y="{plot_top:.1f}" width="{max(0, x1 - x0):.1f}" '
            f'height="{plot_height:.1f}" fill="{COLORS["grey"]}" opacity="0.65"/>'
        )
    return bands


def y_grid(ticks, y_scale, plot_left, plot_right, unit, digits=1):
    parts = []
    for tick in ticks:
        y = y_scale(tick)
        parts.append(f'<line x1="{plot_left}" y1="{y:.1f}" x2="{plot_right}" y2="{y:.1f}" stroke="{COLORS["grid"]}" stroke-width="1"/>')
        parts.append(svg_text(plot_left - 16, y + 7, f"{tick:.{digits}f}{unit}", 18, 400, COLORS["mid"], "end"))
    return parts


def time_axes(start, end, x_scale, plot_left, plot_right, plot_top, plot_bottom):
    parts = []
    for year in range(start.year, end.year + 1):
        d = dt.date(year, 1, 1)
        if d < start or d > end:
            continue
        x = x_scale(date_to_num(d))
        parts.append(f'<line x1="{x:.1f}" y1="{plot_bottom:.1f}" x2="{x:.1f}" y2="{plot_bottom + 8:.1f}" stroke="{COLORS["axis"]}" stroke-width="1.5"/>')
        parts.append(svg_text(x, plot_bottom + 34, str(year), 18, 400, COLORS["mid"], "middle"))
    parts.append(f'<line x1="{plot_left}" y1="{plot_bottom:.1f}" x2="{plot_right}" y2="{plot_bottom:.1f}" stroke="{COLORS["axis"]}" stroke-width="1.5"/>')
    parts.append(f'<line x1="{plot_left}" y1="{plot_top:.1f}" x2="{plot_left}" y2="{plot_bottom:.1f}" stroke="{COLORS["axis"]}" stroke-width="1.5"/>')
    return parts


def spread_labels(ys, gap, top, bottom):
    """Nudge label centres apart so none are closer than `gap`, staying within [top, bottom]."""
    order = sorted(range(len(ys)), key=lambda i: ys[i])
    pos = [0.0] * len(ys)
    prev = None
    for i in order:
        pos[i] = max(ys[i], top) if prev is None else max(ys[i], prev + gap)
        prev = pos[i]
    nxt = None
    for i in reversed(order):
        pos[i] = min(pos[i], bottom) if nxt is None else min(pos[i], nxt - gap)
        nxt = pos[i]
    return pos


def end_labels(entries, x_scale, y_scale, plot_right, plot_top, plot_bottom, latest_date, unit="%"):
    """entries: (date, value, color, short[, hollow]). Two-line labels in the right gutter with leader lines.
    A hollow entry is a marker off the line (e.g. the current target) and keeps its label text as given."""
    parts = []
    points = [(x_scale(date_to_num(e[0])), y_scale(e[1])) for e in entries]
    centres = spread_labels([y for _, y in points], 52, plot_top + 16, plot_bottom - 20)
    for entry, (px, py), cy in zip(entries, points, centres):
        d, v, color, short = entry[:4]
        hollow = len(entry) > 4 and entry[4]
        lx = plot_right + 18
        if hollow:
            parts.append(f'<circle cx="{px:.1f}" cy="{py:.1f}" r="6.5" fill="white" stroke="{color}" stroke-width="3"/>')
        else:
            parts.append(f'<circle cx="{px:.1f}" cy="{py:.1f}" r="5.5" fill="{color}"/>')
        parts.append(f'<path d="M {px + 7:.1f} {py:.1f} L {lx - 4:.1f} {cy - 7:.1f}" fill="none" stroke="{color}" stroke-width="1.5" opacity="0.6"/>')
        parts.append(svg_text(lx, cy, f"{v:.2f}{unit}", 22, 700, color))
        sub = short if (hollow or d == latest_date) else f"{short} · {d.strftime('%b')}"
        parts.append(svg_text(lx, cy + 21, sub, 16, 600, color))
    return parts


# ---------------------------------------------------------------- charts

def value_series(merged, start, end, col):
    return [(date, lookup(merged, date, col)) for date in month_range(start, end)]


def f6_note(merged, latest_date):
    f6_latest = latest_on_or_before(merged, "mortgage_outstanding", latest_date)
    lag = f" (latest {date_label(f6_latest)})" if f6_latest and f6_latest != latest_date else ""
    return f6_latest, lag


def pending_decision(merged, decisions, latest_date):
    """Latest target change if the monthly-average line has not caught up with it yet."""
    if not decisions:
        return None
    d, old, new = decisions[-1]
    avg = lookup(merged, latest_date, "cash_rate")
    if avg is None or abs(avg - new) < 0.005:
        return None
    return d, old, new


def render_overlay(merged, episodes, decisions, width, height, latest_date):
    start = dt.date(2021, 1, 1)
    cols = ["cash_rate", "mortgage_variable", "mortgage_outstanding", "mortgage_new", "saver_online"]
    title = "Who feels the RBA first? Cash rate vs mortgage vs savings rates"
    subtitle = f"Monthly, {date_label(start)} to {date_label(latest_date)}: advertised and actual mortgage rates"
    head, header_bottom = header_svg(width, title, subtitle)
    legend_items = [("line", STYLE[c]["legend"], STYLE[c]["color"], STYLE[c]["dash"]) for c in cols]
    legend_items.append(("band", "Hiking episodes", None, None))
    legend, legend_y = legend_svg(legend_items, 70, header_bottom + 36, width - 70)
    _, f6_lag = f6_note(merged, latest_date)
    pending = pending_decision(merged, decisions, latest_date)
    notes = [
        "Standard variable = banks' advertised indicator rate (F5). Existing / new loans = average rates actually "
        f"charged on owner-occupier variable-rate loans, all institutions (F6){f6_lag}."
    ]
    if pending:
        d, old, new = pending
        notes.append(
            f"Cash rate line is the monthly average of the target. Hollow marker: target raised to {new:.2f}% "
            f"effective {d.day} {d.strftime('%b %Y')} (F1 daily), after the latest retail-rate observations."
        )
    foot, footer_top = footer_svg(width, height, source_line(["F1", "F1.1", "F4", "F5", "F6"]), notes)

    plot_left = 96
    plot_right = width - (230 if is_wide(width) else 200)
    plot_top = legend_y + 40
    plot_bottom = footer_top - 64
    values = []
    for col in cols:
        values.extend(v for _, v in value_series(merged, start, latest_date, col) if v is not None)
    y_min = max(0, math.floor(min(values) * 2) / 2 - 0.25)
    y_max = math.ceil(max(values) * 2) / 2 + 0.25
    x_scale = linear_scale(date_to_num(start), date_to_num(latest_date), plot_left, plot_right)
    y_scale = linear_scale(y_min, y_max, plot_bottom, plot_top)

    body = shaded_bands(episodes, x_scale, plot_top, plot_bottom - plot_top, latest_date)
    body += y_grid(nice_ticks(y_min, y_max, 7), y_scale, plot_left, plot_right, "%")
    body += time_axes(start, latest_date, x_scale, plot_left, plot_right, plot_top, plot_bottom)
    for col in reversed(cols):  # cash rate drawn last, on top
        style = STYLE[col]
        points = [
            (x_scale(date_to_num(d)), y_scale(v) if v is not None else None)
            for d, v in value_series(merged, start, latest_date, col)
        ]
        dash_attr = f' stroke-dasharray="{style["dash"]}"' if style["dash"] else ""
        body.append(
            f'<path d="{line_path(points, step=style["step"])}" fill="none" stroke="{style["color"]}" '
            f'stroke-width="4.5" stroke-linecap="round" stroke-linejoin="round"{dash_attr}/>'
        )
    entries = []
    for col in cols:
        d = latest_on_or_before(merged, col, latest_date)
        if d:
            short = STYLE[col]["short"]
            if col == "cash_rate" and pending:
                short = f"Cash rate · {d.strftime('%b')} avg"
            entries.append((d, lookup(merged, d, col), STYLE[col]["color"], short))
    if pending:
        d, _, new = pending
        entries.append((latest_date, new, COLORS["navy"], f"Target from {d.day} {d.strftime('%b')}", True))
    body += end_labels(entries, x_scale, y_scale, plot_right, plot_top, plot_bottom, latest_date)
    return page(width, height, title, head + legend + body + foot)


def render_scoreboard(stats, width, height):
    episodes = [item for item in stats if item]
    cols = RETAIL_COLS
    title = "The pass-through gap: how much of the RBA's hikes reached each rate"
    subtitle = "Cumulative retail-rate change divided by cumulative cash-rate change"
    head, header_bottom = header_svg(width, title, subtitle)
    legend_items = [("rect", STYLE[c]["legend"], STYLE[c]["color"], None) for c in cols]
    legend, legend_y = legend_svg(legend_items, 70, header_bottom + 36, width - 70)
    notes = ["100% = the full cash-rate change reached that rate. Measured from the month before the first hike."]
    for item in episodes:
        late = [d for d in item["decisions"] if d[0].day >= 20 and dt.date(d[0].year, d[0].month, 1) == item["end_date"]]
        for d, _, new in late:
            notes.append(
                f"{episode_short(item)} excludes the {d.day} {d.strftime('%b %Y')} rise to {new:.2f}%: "
                "it took effect after the latest retail-rate observations."
            )
        lagged = sorted({d for c, d in item["full_end"].items() if d and d != item["end_date"]})
        if lagged:
            notes.append(
                f"{episode_short(item)}: F6 actual-rate bars measured to {date_label(lagged[-1])}, the latest F6 month "
                f"(cash rate +{fmt(item['full_cash_change']['mortgage_outstanding'])}pp over that window)."
            )
    foot, footer_top = footer_svg(width, height, source_line(["F1", "F1.1", "F4", "F5", "F6"]), notes)

    plot_left = 110
    plot_right = width - 70
    plot_top = legend_y + 56
    plot_bottom = footer_top - 120
    vals = [v * 100 for item in episodes for v in item["full"].values() if v is not None]
    y_min = min(0, min(vals) - 10) if vals else 0
    y_max = max(120, max(vals) + 20) if vals else 120
    y_scale = linear_scale(y_min, y_max, plot_bottom, plot_top)
    body = y_grid(nice_ticks(y_min, y_max, 7), y_scale, plot_left, plot_right, "%", 0)
    y100 = y_scale(100)
    body.append(f'<line x1="{plot_left}" y1="{y100:.1f}" x2="{plot_right}" y2="{y100:.1f}" stroke="{COLORS["dark"]}" stroke-width="1.5" stroke-dasharray="6 7" opacity="0.65"/>')
    body.append(f'<line x1="{plot_left}" y1="{y_scale(0):.1f}" x2="{plot_right}" y2="{y_scale(0):.1f}" stroke="{COLORS["axis"]}" stroke-width="1.5"/>')
    body.append(f'<line x1="{plot_left}" y1="{plot_top:.1f}" x2="{plot_left}" y2="{plot_bottom:.1f}" stroke="{COLORS["axis"]}" stroke-width="1.5"/>')

    group_width = (plot_right - plot_left) / max(1, len(episodes))
    n = len(cols)
    bar_width = min(100 if is_wide(width) else 80, group_width / (n * 1.22 + 1.2))
    pitch = bar_width * 1.22
    value_size = 19 if is_wide(width) else 16
    for g, item in enumerate(episodes):
        centre = plot_left + group_width * (g + 0.5)
        for idx, col in enumerate(cols):
            value = item["full"].get(col)
            if value is None:
                continue
            pct = value * 100
            x = centre + (idx - (n - 1) / 2) * pitch - bar_width / 2
            top, bottom = sorted([y_scale(pct), y_scale(0)])
            color = STYLE[col]["color"]
            body.append(f'<rect x="{x:.1f}" y="{top:.1f}" width="{bar_width:.1f}" height="{bottom - top:.1f}" rx="3" fill="{color}"/>')
            label_y = top - 10 if pct >= 0 else bottom + 24
            body.append(svg_text(x + bar_width / 2, label_y, f"{pct:.0f}%", value_size, 700, color, "middle", halo=True))
        span = f"{date_label(item['base_date'])} to {date_label(item['end_date'])}"
        body.append(svg_text(centre, plot_bottom + 42, episode_short(item), 23, 700, COLORS["dark"], "middle"))
        body.append(svg_text(centre, plot_bottom + 70, span, 17, 400, COLORS["mid"], "middle"))
        body.append(svg_text(centre, plot_bottom + 95, f"Cash rate +{fmt(item['cash_change'])}pp (monthly avg)", 17, 600, COLORS["navy"], "middle"))
    return page(width, height, title, head + legend + body + foot)


def episode_short(item):
    return "2022-23 cycle" if "2022" in item["label"] else "2026 to date"


def render_wedge(merged, episodes, width, height, latest_date, stats_a):
    start = dt.date(2021, 1, 1)
    title = "The wedge: the gap between what borrowers pay and savers earn"
    subtitle = "Mortgage rate minus online saver rate, percentage points"
    head, header_bottom = header_svg(width, title, subtitle)
    spread_style = {
        "advertised": (COLORS["red"], "Standard variable (advertised) minus online saver", "Advertised"),
        "actual": (COLORS["plum"], "Existing variable loans (actual) minus online saver", "Actual"),
    }
    legend_items = [("line", label, color, None) for color, label, _ in spread_style.values()]
    legend_items.append(("band", "Hiking episodes", None, None))
    legend, legend_y = legend_svg(legend_items, 70, header_bottom + 36, width - 70)
    _, f6_lag = f6_note(merged, latest_date)
    notes = [f"Actual wedge uses the F6 average rate on outstanding owner-occupier variable loans{f6_lag}."]
    foot, footer_top = footer_svg(width, height, source_line(["F1.1", "F4", "F5", "F6"]), notes)

    plot_left = 96
    plot_right = width - (230 if is_wide(width) else 200)
    plot_top = legend_y + 40
    plot_bottom = footer_top - 64
    series = {}
    for key, (mortgage_col, _) in SPREADS.items():
        series[key] = [(d, spread(merged, d, mortgage_col)) for d in month_range(start, latest_date)]
    values = [v for pts in series.values() for _, v in pts if v is not None]
    y_min = max(0, math.floor(min(values) * 2) / 2 - 0.25)
    y_max = math.ceil(max(values) * 2) / 2 + 0.25
    x_scale = linear_scale(date_to_num(start), date_to_num(latest_date), plot_left, plot_right)
    y_scale = linear_scale(y_min, y_max, plot_bottom, plot_top)

    body = shaded_bands(episodes, x_scale, plot_top, plot_bottom - plot_top, latest_date)
    body += y_grid(nice_ticks(y_min, y_max, 7), y_scale, plot_left, plot_right, "pp")
    body += time_axes(start, latest_date, x_scale, plot_left, plot_right, plot_top, plot_bottom)
    entries = []
    for key, pts in series.items():
        color, _, short = spread_style[key]
        points = [(x_scale(date_to_num(d)), y_scale(v) if v is not None else None) for d, v in pts]
        body.append(f'<path d="{line_path(points)}" fill="none" stroke="{color}" stroke-width="4.5" stroke-linecap="round" stroke-linejoin="round"/>')
        last = [(d, v) for d, v in pts if v is not None][-1]
        entries.append((last[0], last[1], color, f"{short} wedge"))
        if stats_a:
            s = stats_a["spreads"][key]
            if s["start"] is not None:
                bx = x_scale(date_to_num(stats_a["base_date"]))
                by = y_scale(s["start"])
                body.append(f'<circle cx="{bx:.1f}" cy="{by:.1f}" r="5.5" fill="white" stroke="{color}" stroke-width="3"/>')
                body.append(svg_text(bx + 14, by + 30, f"{s['start']:.2f}pp before 2022 hikes", 18, 700, color, "start", halo=True))
    body += end_labels(entries, x_scale, y_scale, plot_right, plot_top, plot_bottom, latest_date, "pp")
    return page(width, height, title, head + legend + body + foot)


SIZES = {"portrait": (1080, 1350), "widescreen": (1920, 1080)}


def write_html_charts(merged, stats, decisions, latest_date):
    episodes = [(item["label"], item["start_hike"], item["end_date"]) for item in stats if item]
    outputs = []
    for size_name, (width, height) in SIZES.items():
        charts = [
            ("chart1_overlay", render_overlay(merged, episodes, decisions, width, height, latest_date)),
            ("chart2_passthrough", render_scoreboard(stats, width, height)),
            ("chart3_wedge", render_wedge(merged, episodes, width, height, latest_date, stats[0])),
        ]
        for base, content in charts:
            path = OUT / f"{base}_{size_name}.html"
            path.write_text(content, encoding="utf-8")
            outputs.append((base, size_name, path))
    return outputs


def find_browser():
    # Chrome first: on this machine Edge's launcher returns before the screenshot is written.
    candidates = [
        Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
        Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
        Path(r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"),
    ]
    for path in candidates:
        if path.exists():
            return str(path)
    return shutil.which("msedge") or shutil.which("chrome")


def render_pngs(chart_html):
    """Screenshot each HTML chart with a headless browser; PNG names match the original build."""
    browser = find_browser()
    if not browser:
        print("No Chrome/Edge found; PNGs not rendered.")
        return []
    outputs = []
    for base, size_name, path in chart_html:
        width, height = SIZES[size_name]
        png = OUT / (f"{base}.png" if size_name == "portrait" else f"{base}_widescreen.png")
        png.unlink(missing_ok=True)
        # A fresh profile per shot: reusing one makes later launches hand off to the
        # still-running instance and exit without writing a screenshot. Edge can also
        # keep profile files locked briefly after exit, hence ignore_cleanup_errors.
        with tempfile.TemporaryDirectory(ignore_cleanup_errors=True) as profile:
            subprocess.run(
                [
                    browser, "--headless=new", "--disable-gpu", "--hide-scrollbars",
                    "--force-device-scale-factor=1", f"--window-size={width},{height}",
                    f"--user-data-dir={profile}", f"--screenshot={png}", path.as_uri(),
                ],
                check=True, capture_output=True, timeout=120,
            )
        if not png.exists():
            raise RuntimeError(f"Browser exited without writing {png}")
        outputs.append(png)
    return outputs


# ---------------------------------------------------------------- text outputs

def text_table(rows, headers):
    widths = [len(header) for header in headers]
    for row in rows:
        for idx, value in enumerate(row):
            widths[idx] = max(widths[idx], len(str(value)))
    lines = [" | ".join(str(headers[idx]).ljust(widths[idx]) for idx in range(len(headers)))]
    lines.append("-+-".join("-" * width for width in widths))
    for row in rows:
        lines.append(" | ".join(str(row[idx]).ljust(widths[idx]) for idx in range(len(headers))))
    return "\n".join(lines)


def dataset_report(tables, selected):
    lines = []
    for table_name, table in tables.items():
        lines.append(f"\n=== {table_name} raw rows 0-12 ===")
        for idx, row in enumerate(table["rows"][:13]):
            lines.append(f"{idx}: {format_row(row)}")
        lines.append(f"\n=== {table_name} column descriptions ===")
        for idx, desc, sid in table["metadata"]:
            lines.append(f"col {idx}: {desc} | series_id={sid}")
    lines.append("\n=== Selected series verification ===")
    for col_name, info in selected.items():
        clean = non_missing(info["values"])
        lines.append(f"\n{col_name}")
        lines.append(f"table: {info['table']}")
        lines.append(f"description: {info['description']}")
        lines.append(f"series_id: {info['series_id']}")
        if clean:
            lines.append(f"date range: {iso_month(clean[0][0])} to {iso_month(clean[-1][0])}")
        else:
            lines.append("date range: no non-missing observations")
        lines.append("first 5 non-missing rows:")
        lines.extend(f"  {iso_month(d)} {v:.4g}" for d, v in clean[:5])
        lines.append("last 5 non-missing rows:")
        lines.extend(f"  {iso_month(d)} {v:.4g}" for d, v in clean[-5:])
    return "\n".join(lines)


def merged_tail_report(merged, latest_date):
    rows = []
    for date in month_range(add_months(latest_date, -17), latest_date):
        row = merged.get(date, {})
        rows.append([iso_month(date)] + [fmt(row.get(col)) for col in MERGED_COLS])
    return text_table(rows, ["date"] + MERGED_COLS)


def stats_report(stats):
    lines = ["PASS-THROUGH STATISTICS", ""]
    for item in stats:
        if not item:
            lines.append("Episode not available: no cash-rate rise detected in 2026 data.")
            continue
        lines.append(f"=== {item['label']} ===")
        lines.append(f"First hike month: {iso_month(item['start_hike'])}")
        lines.append(f"Baseline month: {iso_month(item['base_date'])}")
        lines.append(f"End month: {iso_month(item['end_date'])}")
        lines.append(f"Total cash-rate change: {fmt(item['cash_change'])} percentage points (monthly average, F1.1)")
        lines.append("Actual target changes in the window (F1 daily, effective date):")
        for d, old, new in item["decisions"]:
            lines.append(f"  {d.isoformat()}: {old:.2f}% -> {new:.2f}% ({(new - old) * 100:+.0f}bp)")
        late = [d for d in item["decisions"] if d[0].day >= 20 and dt.date(d[0].year, d[0].month, 1) == item["end_date"]]
        for d, old, new in late:
            lines.append(
                f"  NOTE: the {d.strftime('%d %b %Y')} change takes effect at month end, so it is almost absent from the "
                f"{date_label(item['end_date'])} monthly average and from that month's retail-rate observations."
            )
        lines.append("")
        lines.append("Full-episode cumulative pass-through:")
        for col in RETAIL_COLS:
            end = item["full_end"][col]
            note = ""
            if end and end != item["end_date"]:
                note = f"  [to {iso_month(end)}; cash-rate change {fmt(item['full_cash_change'][col])} pp]"
            lines.append(f"  {STAT_LABEL[col]}: {fmt_pct(item['full'][col])}{note}")
        lines.append("")
        for key, (_, label) in SPREADS.items():
            s = item["spreads"][key]
            end_label = iso_month(s["end_date"]) if s["end_date"] else "n/a"
            lines.append(f"{label}:")
            lines.append(f"  Start ({iso_month(item['base_date'])}): {fmt(s['start'])} percentage points")
            lines.append(f"  End ({end_label}): {fmt(s['end'])} percentage points")
            lines.append(f"  Change: {fmt(s['change'])} percentage points")
        lines.append("")
        lines.append("Individual hikes: pass-through at 1 month and 3 months")
        lines.append("  (monthly-average cash rate, so a mid-month decision is split across two months)")
        for hike in item["hikes"]:
            if hike["decisions"]:
                tag = "decision " + ", ".join(
                    f"{d.strftime('%d %b')} {old:.2f}->{new:.2f}" for d, old, new in hike["decisions"]
                )
            else:
                tag = "no decision this month: carry-over of previous month's decision"
            lines.append(f"  Hike month {iso_month(hike['date'])}; cash-rate change {fmt(hike['cash_change'])} pp [{tag}]")
            for col in RETAIL_COLS:
                lines.append(
                    f"    {STYLE[col]['short']}: 1m {fmt_pct(hike['lag_1m'][col])}; 3m {fmt_pct(hike['lag_3m'][col])}"
                )
        lines.append("")
    return "\n".join(lines)


def methodology(merged_start, latest_date, f6_latest):
    text = (
        "I downloaded RBA Statistical Tables F1, F1.1, F4, F5 and F6 and used monthly observations from "
        f"{date_label(merged_start)} to {date_label(latest_date)}. F1.1 supplies the monthly-average cash-rate "
        "target (F1 daily dates each decision); F5 the "
        "banks' advertised standard variable owner-occupier housing rate; F6 the average rates actually charged "
        "on existing and newly funded owner-occupier variable-rate loans across all institutions (latest "
        f"{date_label(f6_latest)}); F4 online savings and banks' $10,000 one-year term deposit rates. "
        "Pass-through is the change in each retail rate divided by the change in the cash-rate target, from the "
        "month before the first hike to the episode end, or each series' latest month, with one- and three-month "
        "lags checked per hike. Monthly averages split mid-month decisions across two months, so single-hike "
        "figures are noisy. Gaps are left as gaps. Caveat: F4 and F5 are advertised rates; F6 averages hide "
        "wide dispersion across borrowers and lenders."
    )
    return text


def draft_post(merged, stats, decisions, latest_any_date, latest_common_date):
    a, b = stats[0], stats[1] if len(stats) > 1 else None
    f6_date = latest_on_or_before(merged, "mortgage_outstanding", latest_any_date)
    svr_now = lookup(merged, latest_common_date, "mortgage_variable")
    saver_now = lookup(merged, latest_common_date, "saver_online")
    existing_now = lookup(merged, f6_date, "mortgage_outstanding")
    new_now = lookup(merged, f6_date, "mortgage_new")

    opening = f"With the RBA's next decision due on {NEXT_DECISION}, it's worth asking again: when the cash rate moves, who feels it first?"
    pending = pending_decision(merged, decisions, latest_any_date)
    if b and b["decisions"]:
        first_old = b["decisions"][0][1]
        target_now = decisions[-1][2]
        opening += (
            f" The cash-rate target has gone from {first_old:.2f}% to {target_now:.2f}% in {len(b['decisions'])} "
            f"moves since {b['decisions'][0][0].strftime('%B')}."
        )
    if pending:
        d, _, new = pending
        opening += (
            f" The latest rise, to {new:.2f}% on {d.day} {d.strftime('%B')}, came after the most recent retail-rate "
            "data, so the numbers below cover the moves before it."
        )

    chart1 = (
        f"Chart 1: the advertised standard variable rate is {fmt(svr_now)}% ({date_label(latest_common_date)}), "
        f"but the average existing variable borrower actually pays {fmt(existing_now)}% and new variable borrowers "
        f"{fmt(new_now)}% ({date_label(f6_date)}). Online savers earn {fmt(saver_now)}%."
    )

    def pt_line(item):
        f = item["full"]
        return (
            f"advertised variable {fmt_pct(f['mortgage_variable'])}, existing loans {fmt_pct(f['mortgage_outstanding'])}, "
            f"new loans {fmt_pct(f['mortgage_new'])}, online savers {fmt_pct(f['saver_online'])}, "
            f"1-year term deposits {fmt_pct(f['term_deposit'])}"
        )

    chart2 = f"Chart 2: share of the 2022-23 hikes that reached each rate: {pt_line(a)}."
    if b:
        f6_end = b["full_end"]["mortgage_outstanding"]
        f6_clause = f" (actual-rate figures to {date_label(f6_end)})" if f6_end and f6_end != b["end_date"] else ""
        chart2 += f" For 2026 to {date_label(b['end_date'])}{f6_clause}: {pt_line(b)}."

    adv_a, act_a = a["spreads"]["advertised"], a["spreads"]["actual"]
    adv_now = spread(merged, latest_common_date, "mortgage_variable")
    act_now = spread(merged, f6_date, "mortgage_outstanding")
    chart3 = (
        f"Chart 3: the advertised wedge (standard variable minus online saver) went from {fmt(adv_a['start'])}pp "
        f"before the 2022 hikes to {fmt(adv_a['end'])}pp by late 2023, and is {fmt(adv_now)}pp now. Using what "
        f"existing borrowers actually pay, the wedge went from {fmt(act_a['start'])}pp to {fmt(act_a['end'])}pp, "
        f"and is {fmt(act_now)}pp now."
    )

    post = (
        f"{opening}\n\n"
        f"{chart1}\n\n"
        f"{chart2}\n\n"
        f"{chart3}\n\n"
        "The advertised rate is the headline, but it is not what most borrowers pay. Looking at actual loan rates "
        "changes both the level of the borrower-saver wedge and how it has moved through the cycle.\n\n"
        "This is exactly the kind of interest rate pass-through question where product design, bank pricing, and household balance sheets meet.\n\n"
        "Curious how others read the borrower-saver split in this cycle.\n\n"
        "Data: RBA Statistical Tables F1, F1.1, F4, F5, F6 (public). Full methodology in comments.\n\n"
        "#RBA #InterestRates #Mortgages #Banking #AustralianEconomy"
    )
    return post


# ---------------------------------------------------------------- main

def export_dashboard_data(merged, stats, decisions):
    """Tidy CSVs for the interactive dashboard: percentages already x100, rates in %."""
    folder = OUT / "dashboard_data"
    folder.mkdir(parents=True, exist_ok=True)
    num = lambda value, digits=2: "" if value is None else f"{value:.{digits}f}"
    pct = lambda value: "" if value is None else f"{value * 100:.1f}"

    with (folder / "rates_monthly.csv").open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["month"] + MERGED_COLS)
        for date in sorted(merged):
            writer.writerow([iso_month(date)] + [num(merged[date].get(c)) for c in MERGED_COLS])

    with (folder / "passthrough.csv").open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["episode", "series", "passthrough_pct", "base_month", "measured_to", "cash_change_pp"])
        for item in stats:
            if not item:
                continue
            for col in RETAIL_COLS:
                end = item["full_end"][col]
                writer.writerow([
                    "2022-23" if "2022" in item["label"] else "2026",
                    col,
                    pct(item["full"][col]),
                    iso_month(item["base_date"]),
                    iso_month(end) if end else "",
                    num(item["full_cash_change"][col]),
                ])

    with (folder / "rba_decisions.csv").open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["effective_date", "from_rate", "to_rate", "change_bp"])
        for d, old, new in decisions:
            if d >= dt.date(2020, 1, 1):
                writer.writerow([d.isoformat(), num(old), num(new), f"{(new - old) * 100:.0f}"])

    with (folder / "hike_lags.csv").open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(
            ["hike_month", "episode", "type", "decision_date", "cash_change_pp"]
            + [f"{c}_{lag}" for lag in ("1m", "3m") for c in RETAIL_COLS]
        )
        for item in stats:
            if not item:
                continue
            for hike in item["hikes"]:
                writer.writerow(
                    [
                        iso_month(hike["date"]),
                        "2022-23" if "2022" in item["label"] else "2026",
                        "RBA decision" if hike["decisions"] else "Carry-over",
                        hike["decisions"][0][0].isoformat() if hike["decisions"] else "",
                        num(hike["cash_change"]),
                    ]
                    + [pct(hike[f"lag_{lag}"][c]) for lag in ("1m", "3m") for c in RETAIL_COLS]
                )
    return folder


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    WORK.mkdir(parents=True, exist_ok=True)
    tables = {name: parse_table(name, path) for name, path in FILES.items()}
    decisions = cash_target_decisions(DAILY_CASH)
    selected = {}
    for col_name, spec in SERIES.items():
        table = tables[spec["table"]]
        col_idx = find_col(table["rows"], spec["contains"])
        selected[col_name] = {
            "table": spec["table"],
            "col_idx": col_idx,
            "description": table["rows"][2][col_idx],
            "series_id": table["rows"][10][col_idx] if col_idx < len(table["rows"][10]) else "",
            "values": series_values(table, col_idx),
        }
    merged = {}
    for col_name in MERGED_COLS:
        for date, value in selected[col_name]["values"]:
            if dt.date(2020, 1, 1) <= date:
                merged.setdefault(date, {})[col_name] = value
    latest_any_date = max(date for date, row in merged.items() if any(v is not None for v in row.values()))
    latest_common_date = latest_complete_date(merged, CORE_COLS)
    if latest_common_date is None:
        raise ValueError("No month has complete observations for all core series.")
    merged = {date: row for date, row in merged.items() if date <= latest_any_date}
    f6_latest = latest_on_or_before(merged, "mortgage_outstanding", latest_any_date)
    missing = {col: [d for d in sorted(merged) if merged[d].get(col) is None] for col in MERGED_COLS}

    stats_a = episode_stats(merged, decisions, "Episode A: 2022-23 tightening cycle", dt.date(2022, 5, 1), dt.date(2023, 11, 1))
    b_start, b_end = detect_2026_episode(merged, latest_common_date)
    stats_b = episode_stats(merged, decisions, "Episode B: 2026 tightening cycle to date", b_start, b_end) if b_start else None
    stats = [stats_a, stats_b]

    verification = dataset_report(tables, selected)
    verification += "\n\n=== Merged frame tail(18) ===\n"
    verification += merged_tail_report(merged, latest_any_date)
    verification += f"\n\nLatest merged month with any selected observation: {iso_month(latest_any_date)}"
    verification += f"\nLatest common month with all core (F1.1/F4/F5) series: {iso_month(latest_common_date)}"
    verification += f"\nLatest F6 month: {iso_month(f6_latest)}"
    verification += "\n\n=== Missing values in selected merged frame, Jan 2020 to latest ==="
    for col, dates in missing.items():
        if dates:
            verification += f"\n{col}: {len(dates)} missing months; first few: " + ", ".join(iso_month(d) for d in dates[:12])
        else:
            verification += f"\n{col}: none"
    print(verification[-2500:])

    stats_text = stats_report(stats)
    print("\n" + stats_text)
    method_text = methodology(dt.date(2020, 1, 1), latest_common_date, f6_latest)
    (OUT / "passthrough_stats.txt").write_text(stats_text, encoding="utf-8")
    (OUT / "methodology.txt").write_text(method_text, encoding="utf-8")
    (OUT / "draft_post.txt").write_text(draft_post(merged, stats, decisions, latest_any_date, latest_common_date), encoding="utf-8")
    (OUT / "verification_dump.txt").write_text(verification, encoding="utf-8")

    with (WORK / "merged.csv").open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["date"] + MERGED_COLS)
        for date in sorted(merged):
            row = merged[date]
            writer.writerow([iso_month(date)] + ["" if row.get(c) is None else row.get(c) for c in MERGED_COLS])
    with (WORK / "episodes.csv").open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(
            ["label", "base_date", "start_hike", "end_date", "cash_change"]
            + [f"{c}_passthrough" for c in RETAIL_COLS]
            + [f"{c}_measured_to" for c in RETAIL_COLS]
            + [f"spread_{k}_{p}" for k in SPREADS for p in ("start", "end", "change")]
        )
        for item in stats:
            if not item:
                continue
            writer.writerow(
                [item["label"], iso_month(item["base_date"]), iso_month(item["start_hike"]), iso_month(item["end_date"]), item["cash_change"]]
                + [item["full"][c] for c in RETAIL_COLS]
                + [iso_month(item["full_end"][c]) if item["full_end"][c] else "" for c in RETAIL_COLS]
                + [item["spreads"][k][p] for k in SPREADS for p in ("start", "end", "change")]
            )

    export_dashboard_data(merged, stats, decisions)
    chart_html = write_html_charts(merged, stats, decisions, latest_any_date)
    pngs = render_pngs(chart_html)
    manifest = {
        "latest_month_any_selected_series": iso_month(latest_any_date),
        "latest_month_complete_core_series": iso_month(latest_common_date),
        "latest_month_f6": iso_month(f6_latest),
        "html": [path.relative_to(PROJECT).as_posix() for _, _, path in chart_html],
        "png": [path.relative_to(PROJECT).as_posix() for path in pngs],
    }
    (OUT / "chart_manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")

    print(f"\nMethodology word count: {len(method_text.split())}")
    print("\nFINAL VERIFICATION CHECKLIST")
    print("[x] Every chart datapoint traces to downloaded RBA files in data/raw.")
    print(f"[x] Latest data month: {iso_month(latest_any_date)}; core comparison month: {iso_month(latest_common_date)}; F6: {iso_month(f6_latest)}.")
    print(f"[x] PNGs rendered: {len(pngs)}")
    print("[x] Headline numbers to manually verify before publishing:")
    d, _, target = decisions[-1]
    print(f"    - Cash-rate target: {target:.2f}% from {d.isoformat()} (F1 daily); {iso_month(latest_any_date)} monthly average {fmt(lookup(merged, latest_any_date, 'cash_rate'))}%")
    print(f"    - 2022-23 pass-through, advertised SVR {fmt_pct(stats_a['full']['mortgage_variable'])}; existing loans {fmt_pct(stats_a['full']['mortgage_outstanding'])}")
    print(f"    - Latest existing variable loan rate: {fmt(lookup(merged, f6_latest, 'mortgage_outstanding'))}% ({iso_month(f6_latest)})")


if __name__ == "__main__":
    main()
