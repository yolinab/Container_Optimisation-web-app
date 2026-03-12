"""app/utils/export_excel.py
Export container packing results to a formatted Excel report.

Sheets:
  1. Overview        — summary metrics + per-container fill table + bar chart
  2. Container Details — full breakdown per container (pallet rows + NP box zones)
  3. Packing Layout  — colour-coded cell-grid (top-down view of pallet arrangement)
  4. Recommendations — what to add to the order to fill free space
"""

from pathlib import Path
from typing import Dict, List, Any, Optional
import datetime

try:
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.chart import BarChart, Reference
    from openpyxl.chart.series import SeriesLabel
    _HAS_OPENPYXL = True
except ImportError:
    _HAS_OPENPYXL = False


# ═══════════════════════════════════════════════════════════════════════════
# Colour palette
# ═══════════════════════════════════════════════════════════════════════════

_HEADER_BG    = "1A3A5C"   # dark navy  — main headers
_SUBHDR_BG    = "2E6DA4"   # mid blue   — section headers
_ACCENT_BG    = "D6E4F0"   # pale blue  — accent rows / column headers
_ALT_ROW      = "EBF5FB"   # very light blue — alternating rows
_WHITE        = "FFFFFF"
_LIGHT_GRAY   = "F5F5F5"

# Block type colours by footprint prefix
_BLOCK_PAL: Dict[str, str] = {
    "115x115": "AED6F1",   # sky blue
    "115x108": "FAD7A0",   # peach
    "115x77":  "A9DFBF",   # mint green
    "77x77":   "D7BDE2",   # lavender
}
_NP_BOX_COLOR  = "FFF9C4"  # pale yellow      — NP box zones (placed)
_EMPTY_COLOR   = "EEEEEE"  # light grey        — unused / empty space
_REC_PAL_COLOR = "C8F7C5"  # light fresh green — recommended pallets
_REC_NP_COLOR  = "B2EBF2"  # light teal        — recommended NP boxes

# Fill-rate cell colours
_TL_GOOD = "27AE60"   # dark green — optimally filled (≥ 85 %)
_TL_OK   = "E67E22"   # amber      — not optimally filled (< 85 %)

# Footprint → pallet type code (shown in Container Details sheet)
_FOOTPRINT_TO_PALLET_TYPE: Dict[str, str] = {
    "115x115": "A2",
    "115x108": "A1",
    "115x77":  "C2",
    "77x77":   "D2",
}


# ═══════════════════════════════════════════════════════════════════════════
# Low-level style helpers
# ═══════════════════════════════════════════════════════════════════════════

def _fill(hex_color: str) -> "PatternFill":
    return PatternFill("solid", fgColor=hex_color)

def _font(bold=False, color="000000", size=10, italic=False) -> "Font":
    return Font(bold=bold, color=color, size=size, italic=italic, name="Calibri")

def _align(h="left", v="center", wrap=False) -> "Alignment":
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _thin_border() -> "Border":
    s = Side(style="thin", color="D0D0D0")
    return Border(left=s, right=s, top=s, bottom=s)

def _block_color(key: str) -> str:
    for prefix, color in _BLOCK_PAL.items():
        if key.startswith(prefix):
            return color
    return "CCCCCC"

def _has_recommendations(rec: dict) -> bool:
    """True when the recommendation engine found items that can still be added."""
    return (rec.get("total_pallets_to_add", 0) > 0
            or rec.get("total_np_boxes_to_add", 0) > 0)


def _vol_fill_pct(container: dict, W: int, Hdoor: int, L: int) -> float:
    """Volumetric fill %: actual goods volume / (L × W × Hdoor) × 100.

    - Pallet block: length × container_width × block_height  (full cross-section)
    - NP box zone:  sum of actual placed-box volumes (length × width × height × qty)
    """
    pallet_vol = sum(
        r["length_cm"] * W * r["height_cm"]
        for r in container.get("rows", [])
    )
    box_zone_vol = sum(
        p["length_cm"] * p["width_cm"] * p["height_cm"] * p["quantity"]
        for z in container.get("box_zones", [])
        for p in z.get("placed", [])
    )
    container_vol = float(L * W * Hdoor)
    if container_vol <= 0:
        return 0.0
    return round(100.0 * (pallet_vol + box_zone_vol) / container_vol, 1)


def _set_cell(ws, row: int, col: int, value=None, *,
              bold=False, fg="000000", bg=None, size=10, italic=False,
              h="left", v="center", wrap=False,
              border=False, num_fmt=None):
    """Write and style a single cell. Returns the cell."""
    cell = ws.cell(row=row, column=col, value=value)
    cell.font      = _font(bold=bold, color=fg, size=size, italic=italic)
    cell.alignment = _align(h=h, v=v, wrap=wrap)
    if bg:
        cell.fill = _fill(bg)
    if border:
        cell.border = _thin_border()
    if num_fmt:
        cell.number_format = num_fmt
    return cell


def _merge_write(ws, r1, c1, r2, c2, value="", *,
                 bold=False, fg="000000", bg=None, size=10, italic=False,
                 h="left", v="center"):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    cell = ws.cell(row=r1, column=c1, value=value)
    cell.font      = _font(bold=bold, color=fg, size=size, italic=italic)
    cell.alignment = _align(h=h, v=v)
    if bg:
        cell.fill = _fill(bg)
    return cell


def _section_title(ws, row: int, c1: int, c2: int, title: str, bg=_SUBHDR_BG):
    _merge_write(ws, row, c1, row, c2, title,
                 bold=True, fg=_WHITE, bg=bg, size=11, h="left")
    ws.row_dimensions[row].height = 22


def _col_header_row(ws, row: int, headers: List[str], col_start: int = 1,
                    widths: Optional[List[float]] = None):
    for i, h in enumerate(headers):
        c = col_start + i
        _set_cell(ws, row, c, h, bold=True, bg=_ACCENT_BG, fg=_HEADER_BG,
                  h="center", border=True)
        if widths and i < len(widths):
            ws.column_dimensions[get_column_letter(c)].width = widths[i]
    ws.row_dimensions[row].height = 18


# ═══════════════════════════════════════════════════════════════════════════
# Sheet 1 — Overview
# ═══════════════════════════════════════════════════════════════════════════

def _write_overview(ws, containers, recs, np_boxes, unplaced, config):
    ws.sheet_view.showGridLines = False

    L     = config.get("CONTAINER_LENGTH_CM", 1203)
    W     = config.get("CONTAINER_WIDTH_CM", 235)
    Hdoor = config.get("CONTAINER_DOOR_HEIGHT_CM", 250)
    Wm    = config.get("CONTAINER_MAX_WEIGHT_KG", 18000)

    total_pallets  = sum(c["loaded_value"]  for c in containers)
    total_weight   = sum(c["loaded_weight"] for c in containers)
    total_np_boxes = sum(
        sum(p["quantity"] for z in c.get("box_zones", []) for p in z["placed"])
        for c in containers
    )
    total_rec_pal  = sum(r.get("total_pallets_to_add", 0)  for r in recs)
    total_rec_np   = sum(r.get("total_np_boxes_to_add", 0) for r in recs)
    total_unplaced = sum(e["remaining_qty"] for e in (unplaced or []))

    avg_fill = (
        sum(_vol_fill_pct(c, W, Hdoor, L) for c in containers) / len(containers)
        if containers else 0.0
    )

    # ── Title block ─────────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 18
    ws.column_dimensions["F"].width = 18
    ws.column_dimensions["G"].width = 18
    ws.column_dimensions["H"].width = 18

    _merge_write(ws, 1, 1, 1, 8,
                 "🚢  CONTAINER PACKING REPORT",
                 bold=True, fg=_WHITE, bg=_HEADER_BG, size=16, h="center")
    ws.row_dimensions[1].height = 38

    run_time = datetime.datetime.now().strftime("%d %b %Y  %H:%M")
    _merge_write(ws, 2, 1, 2, 8,
                 f"Generated: {run_time}   |   Objective: {config.get('RECOMMEND_OBJECTIVE', '—')}",
                 fg="666666", size=9, h="center", bg=_LIGHT_GRAY)
    ws.row_dimensions[2].height = 16

    ws.row_dimensions[3].height = 8  # spacer

    # ── Key metrics ──────────────────────────────────────────────────────────
    _section_title(ws, 4, 1, 8, "  KEY METRICS")
    ws.row_dimensions[4].height = 22

    metrics = [
        ("Containers used",             len(containers),        None),
        ("Total pallets loaded",         total_pallets,          None),
        ("Total NP boxes loaded",        total_np_boxes,         None),
        ("Total weight loaded (kg)",     f"{total_weight:,.0f}", None),
        ("Avg container volume fill",    f"{avg_fill:.1f} %",    None),
        ("Recommended extra pallets",    total_rec_pal,          None),
        ("Recommended extra NP boxes",   total_rec_np,           None),
        ("Unplaced NP boxes",            total_unplaced,         None),
    ]

    for i, (label, value, _) in enumerate(metrics):
        r = 5 + i
        bg = _ALT_ROW if i % 2 == 0 else _WHITE
        _set_cell(ws, r, 1, label, bold=True, bg=bg, border=True)
        _set_cell(ws, r, 2, value, bg=bg, h="right", border=True)
        _merge_write(ws, r, 3, r, 8, bg=bg)
        ws.row_dimensions[r].height = 17

    ws.row_dimensions[13].height = 10  # spacer

    # ── Per-container summary table ──────────────────────────────────────────
    _section_title(ws, 14, 1, 8, "  CONTAINER SUMMARY")

    hdrs = ["Container", "Len Used (cm)", "Total (cm)", "Vol Fill %",
            "Pallets", "Weight (kg)", "NP Boxes", "Rec Pallets"]
    _col_header_row(ws, 15, hdrs, col_start=1,
                    widths=[12, 12, 12, 10, 10, 14, 12, 14])

    # Build rec lookup by container index
    rec_by_idx = {r["container_index"]: r for r in recs}

    chart_rows_start = 16  # data starts at row 16 for chart reference

    for i, c in enumerate(containers):
        r      = 16 + i
        idx    = c["container_index"]
        used   = c["used_length_cm"]
        leftov = c["leftover_cm"]
        pals   = c["loaded_value"]
        wt     = c["loaded_weight"]
        zones  = c.get("box_zones", [])
        boxes  = sum(p["quantity"] for z in zones for p in z["placed"])
        rec    = rec_by_idx.get(idx, {})
        r_pal  = rec.get("total_pallets_to_add", 0)
        fill_p = _vol_fill_pct(c, W, Hdoor, L)

        bg      = _ALT_ROW if i % 2 == 0 else _WHITE
        # Green ✓  = no recommendations exist (nothing more can be added).
        # Orange % = recommendations found (more pallets/boxes could be loaded).
        optimal = not _has_recommendations(rec)

        data = [idx, used, L, fill_p, pals, round(wt, 0), boxes, r_pal]
        for j, val in enumerate(data):
            col  = 1 + j
            cell = _set_cell(ws, r, col, val, bg=bg, h="center", border=True)
            if j == 3:   # Fill % cell — always show numeric %, green if well-filled
                well_filled = fill_p >= 85
                cell.number_format = '0.0"%"'
                cell.fill = _fill(_TL_GOOD if well_filled else _TL_OK)
                cell.font = _font(bold=True, color=_WHITE, size=10)
            elif j in (1, 2):
                cell.number_format = '#,##0'
            elif j == 5:
                cell.number_format = '#,##0'
        ws.row_dimensions[r].height = 17

    chart_rows_end = 16 + len(containers) - 1

    # ── Bar chart: fill % per container ─────────────────────────────────────
    chart_row_for_data = chart_rows_start
    try:
        chart = BarChart()
        chart.type         = "col"
        chart.grouping     = "clustered"
        chart.title        = "Container Volume Fill %"
        chart.y_axis.title = "Vol Fill %"
        chart.x_axis.title = "Container"
        chart.y_axis.scaling.min = 0
        chart.y_axis.scaling.max = 100
        chart.height = 12
        chart.width  = 18

        # Data: column D (fill %) rows 16..end
        data_ref = Reference(ws,
                             min_col=4, max_col=4,
                             min_row=chart_rows_start, max_row=chart_rows_end)
        cats_ref = Reference(ws,
                             min_col=1, max_col=1,
                             min_row=chart_rows_start, max_row=chart_rows_end)
        chart.add_data(data_ref, from_rows=False, titles_from_data=False)
        chart.set_categories(cats_ref)
        chart.series[0].title = SeriesLabel(v="Vol Fill %")
        chart.series[0].graphicalProperties.solidFill = "2E6DA4"

        anchor_row = chart_rows_end + 3
        ws.add_chart(chart, f"A{anchor_row}")
    except Exception:
        pass  # chart is optional


# ═══════════════════════════════════════════════════════════════════════════
# Sheet 2 — Container Details
# ═══════════════════════════════════════════════════════════════════════════

def _write_details(ws, containers, config):
    ws.sheet_view.showGridLines = False

    col_widths = [20, 10, 16, 12, 12, 14, 12, 14]
    for i, w in enumerate(col_widths):
        ws.column_dimensions[get_column_letter(i + 1)].width = w

    L     = config.get("CONTAINER_LENGTH_CM", 1203)
    W     = config.get("CONTAINER_WIDTH_CM", 235)
    Hdoor = config.get("CONTAINER_DOOR_HEIGHT_CM", 250)
    Wm    = config.get("CONTAINER_MAX_WEIGHT_KG", 18000)

    row = 1
    for c in containers:
        idx     = c["container_index"]
        used    = c["used_length_cm"]
        leftov  = c["leftover_cm"]
        wt      = c["loaded_weight"]
        pals    = c["loaded_value"]
        vol_p   = _vol_fill_pct(c, W, Hdoor, L)
        len_p   = round(100.0 * used / L, 1) if L else 0
        zones   = c.get("box_zones", [])
        n_boxes = sum(p["quantity"] for z in zones for p in z["placed"])

        # Container header
        _merge_write(ws, row, 1, row, 8,
                     f"  CONTAINER {idx}   —   Vol Fill: {vol_p}%  |  "
                     f"Len used: {used} / {L} cm  ({len_p}%)"
                     f"   |   Pallets: {pals}   Weight: {wt:,.0f} kg   NP Boxes: {n_boxes}",
                     bold=True, fg=_WHITE, bg=_HEADER_BG, size=11)
        ws.row_dimensions[row].height = 24
        row += 1

        # Pallet rows sub-section
        _merge_write(ws, row, 1, row, 8, "  Pallet Row Blocks",
                     bold=True, fg=_WHITE, bg=_SUBHDR_BG, size=10)
        ws.row_dimensions[row].height = 18
        row += 1

        _col_header_row(ws, row,
                        ["Block Type", "Pallet Type", "Footprint", "Length (cm)", "Height (cm)",
                         "Weight (kg)", "Pallets", "Y Start (cm)"],
                        col_start=1)
        row += 1

        for ri, rrow in enumerate(c.get("rows", [])):
            bg = _ALT_ROW if ri % 2 == 0 else _WHITE
            bk = rrow["block_type"]
            footprint_key = bk.split("|")[0] if "|" in bk else bk
            pallet_type   = _FOOTPRINT_TO_PALLET_TYPE.get(footprint_key, "—")
            vals = [
                bk,
                pallet_type,
                footprint_key,
                rrow["length_cm"],
                rrow["height_cm"],
                round(rrow["weight_kg"], 1),
                rrow["pallet_count"],
                rrow["y_start_cm"],
            ]
            for ci, v in enumerate(vals):
                _set_cell(ws, row, ci + 1, v,
                          bg=_block_color(bk) if ci in (0, 1) else bg,
                          border=True, h="center" if ci > 1 else "left")
            ws.row_dimensions[row].height = 16
            row += 1

        # NP box zones sub-section
        if zones:
            _merge_write(ws, row, 1, row, 8, "  NP Box Zones",
                         bold=True, fg=_WHITE, bg=_SUBHDR_BG, size=10)
            ws.row_dimensions[row].height = 18
            row += 1

            _col_header_row(ws, row,
                            ["Zone Type", "Box Label", "Length (cm)", "Height (cm)",
                             "Weight (kg)", "Quantity", "Y Start (cm)"],
                            col_start=1)
            row += 1

            for zi, zone in enumerate(zones):
                for pi, placed in enumerate(zone["placed"]):
                    bg = _ALT_ROW if (zi + pi) % 2 == 0 else _WHITE
                    vals = [
                        zone["zone_type"].upper(),
                        placed["label"],
                        placed["length_cm"],
                        placed["height_cm"],
                        round(placed["weight_kg_total"], 1),
                        placed["quantity"],
                        zone["y_start_cm"],
                    ]
                    for ci, v in enumerate(vals):
                        _set_cell(ws, row, ci + 1, v,
                                  bg=_NP_BOX_COLOR if ci in (0, 1) else bg,
                                  border=True, h="center" if ci > 1 else "left")
                    ws.row_dimensions[row].height = 16
                    row += 1

        ws.row_dimensions[row].height = 10  # spacer between containers
        row += 1


# ═══════════════════════════════════════════════════════════════════════════
# Sheet 3 — Packing Layout  (colour-coded cell grid)
# ═══════════════════════════════════════════════════════════════════════════

_CM_PER_COL    = 12      # 1 spreadsheet column = 12 cm
_LAYOUT_LABEL_COLS = 3   # columns reserved for container labels on the left

def _cm_to_col(cm: int) -> int:
    """Convert a cm position → 1-based column index (offset by label cols)."""
    return _LAYOUT_LABEL_COLS + 1 + (cm // _CM_PER_COL)

def _layout_col_count(L_cm: int) -> int:
    return (L_cm // _CM_PER_COL) + 1


def _color_layout_range(ws, row: int, y_start_cm: int, length_cm: int,
                        color: str, label: str = ""):
    """Colour a run of cells in the layout row and optionally write a label."""
    c_start = _cm_to_col(y_start_cm)
    c_end   = _cm_to_col(y_start_cm + length_cm - 1)
    c_end   = max(c_end, c_start)
    fill    = _fill(color)

    for c in range(c_start, c_end + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill

    # Write label in the first cell of the block (if wide enough)
    if label and (c_end - c_start) >= 2:
        lbl_cell = ws.cell(row=row, column=c_start, value=label)
        lbl_cell.font      = _font(bold=True, color="333333", size=7)
        lbl_cell.alignment = _align(h="left", v="center")


def _write_layout(ws, containers, recs, config):
    ws.sheet_view.showGridLines = False

    L     = config.get("CONTAINER_LENGTH_CM", 1203)
    W     = config.get("CONTAINER_WIDTH_CM", 235)
    Hdoor = config.get("CONTAINER_DOOR_HEIGHT_CM", 250)
    n_layout_cols = _layout_col_count(L)
    total_cols    = _LAYOUT_LABEL_COLS + n_layout_cols + 1

    # ── Legend / title ───────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 8

    # Narrow layout columns (1 col = 12 cm ≈ 1.8 char wide)
    for c in range(_LAYOUT_LABEL_COLS + 1, total_cols + 1):
        ws.column_dimensions[get_column_letter(c)].width = 1.8

    _merge_write(ws, 1, 1, 1, total_cols,
                 "  PACKING LAYOUT   (each cell ≈ 12 cm • colours show block type  •  green = recommended additions)",
                 bold=True, fg=_WHITE, bg=_HEADER_BG, size=12, h="left")
    ws.row_dimensions[1].height = 26

    # ── Colour legend ────────────────────────────────────────────────────────
    legend_items = [
        ("115×115 pallet", _BLOCK_PAL["115x115"]),
        ("115×108 pallet", _BLOCK_PAL["115x108"]),
        ("115×77  pallet", _BLOCK_PAL["115x77"]),
        ("77×77   pallet", _BLOCK_PAL["77x77"]),
        ("NP box zone",    _NP_BOX_COLOR),
        ("Empty space",    _EMPTY_COLOR),
        ("Rec. pallets",   _REC_PAL_COLOR),
        ("Rec. NP boxes",  _REC_NP_COLOR),
    ]
    ws.row_dimensions[2].height = 5  # spacer
    leg_row = 3
    ws.row_dimensions[leg_row].height = 16
    col = 1
    for label, color in legend_items:
        # Colour swatch cell
        swatch = ws.cell(row=leg_row, column=col, value="  " + label)
        swatch.fill      = _fill(color)
        swatch.font      = _font(size=8, color="333333")
        swatch.alignment = _align(h="left", v="center")
        swatch.border    = _thin_border()
        # Span 2 cols for readability
        try:
            ws.merge_cells(start_row=leg_row, start_column=col,
                           end_row=leg_row, end_column=col + 1)
        except Exception:
            pass
        col += 3  # jump to next legend item (leave 1 gap col)

    ws.row_dimensions[4].height = 5  # spacer

    # ── Ruler: cm markers ────────────────────────────────────────────────────
    ruler_row = 5
    ws.row_dimensions[ruler_row].height = 12
    # Label cols
    _set_cell(ws, ruler_row, 1, "Container", bold=True, size=8, h="center", bg=_ACCENT_BG)
    _set_cell(ws, ruler_row, 2, "Vol Fill %", bold=True, size=8, h="center", bg=_ACCENT_BG)
    _set_cell(ws, ruler_row, 3, "Layer",     bold=True, size=8, h="center", bg=_ACCENT_BG)

    # Tick marks every 100 cm
    for tick_cm in range(0, L + 1, 100):
        c = _cm_to_col(tick_cm)
        if c <= total_cols:
            cell = ws.cell(row=ruler_row, column=c, value=str(tick_cm))
            cell.font      = _font(size=7, color="555555", bold=True)
            cell.alignment = _align(h="center", v="center")
            cell.fill      = _fill(_ACCENT_BG)

    # ── Container rows ───────────────────────────────────────────────────────
    rec_by_idx = {r["container_index"]: r for r in recs}

    cur_row = 6
    ROW_H_PALLET = 24
    ROW_H_BOX    = 16
    ROW_H_REC    = 16
    ROW_H_SPACER = 4

    for c in containers:
        idx    = c["container_index"]
        used   = c["used_length_cm"]
        fill_p  = _vol_fill_pct(c, W, Hdoor, L)
        rec     = rec_by_idx.get(idx, {})
        tl      = _TL_GOOD if fill_p >= 85 else _TL_OK

        rows_data = c.get("rows", [])
        zones     = c.get("box_zones", [])

        # Determine which sub-rows to draw
        has_np_boxes = bool(zones)
        tail_pls     = rec.get("tail_placements", [])
        atop_pls     = rec.get("atop_placements", [])
        has_recs     = bool(tail_pls or atop_pls)

        # Determine row indices
        pallet_row = cur_row
        box_row    = cur_row + 1 if has_np_boxes or has_recs else None
        rec_row    = (cur_row + 2) if (has_np_boxes and has_recs) else \
                     (cur_row + 1) if (has_recs and not has_np_boxes) else None

        n_sub_rows = 1 + (1 if has_np_boxes else 0) + (1 if has_recs else 0)

        # ── Label column ─────────────────────────────────────────────────────
        # Container number (merged across all sub-rows)
        _merge_write(ws, pallet_row, 1,
                     pallet_row + n_sub_rows - 1, 1,
                     f"Cont. {idx}",
                     bold=True, fg=_WHITE, bg=_SUBHDR_BG, size=9, h="center", v="center")
        # Fill % / optimal marker (merged across sub-rows).
        # Always write numeric fill_p so the chart can read it; number format
        # shows "✓" for optimal containers, "XX.X%" for others.
        _merge_write(ws, pallet_row, 2,
                     pallet_row + n_sub_rows - 1, 2,
                     fill_p,
                     bold=True, fg=_WHITE, bg=tl,
                     size=9, h="center", v="center")
        fill_cell = ws.cell(row=pallet_row, column=2)
        fill_cell.fill = _fill(tl)
        fill_cell.number_format = '0.0"%"'

        # Layer labels in col 3
        _set_cell(ws, pallet_row, 3, "Pallets", size=7, h="center", bg=_LIGHT_GRAY)
        if has_np_boxes and box_row:
            _set_cell(ws, box_row, 3, "NP Boxes", size=7, h="center", bg=_LIGHT_GRAY)
        if has_recs and rec_row:
            _set_cell(ws, rec_row, 3, "Rec +", size=7, bold=True, h="center",
                      bg=_REC_PAL_COLOR)

        # Row heights
        ws.row_dimensions[pallet_row].height = ROW_H_PALLET
        if box_row:
            ws.row_dimensions[box_row].height = ROW_H_BOX
        if rec_row:
            ws.row_dimensions[rec_row].height = ROW_H_REC

        # ── Background: shade entire container length as "empty" ─────────────
        for layout_c in range(_LAYOUT_LABEL_COLS + 1,
                              _cm_to_col(L) + 1):
            for sub_r in range(n_sub_rows):
                cell = ws.cell(row=pallet_row + sub_r, column=layout_c)
                cell.fill = _fill(_EMPTY_COLOR)

        # ── Pallet row blocks ────────────────────────────────────────────────
        for rrow in rows_data:
            bk     = rrow["block_type"]
            y0     = rrow["y_start_cm"]
            length = rrow["length_cm"]
            color  = _block_color(bk)
            short  = bk.split("|")[0] if "|" in bk else bk
            # e.g. "115x115" → show "115×115 (8p)"
            label  = f"{short}({rrow['pallet_count']}p)"
            _color_layout_range(ws, pallet_row, y0, length, color, label)

        # ── NP box zones (on box_row) ────────────────────────────────────────
        if has_np_boxes and box_row:
            for zone in zones:
                y0 = zone["y_start_cm"]
                lz = zone["length_cm"]
                _color_layout_range(ws, box_row, y0, lz, _NP_BOX_COLOR,
                                    f"{sum(p['quantity'] for p in zone['placed'])}b")

        # ── Recommendations ──────────────────────────────────────────────────
        if has_recs and rec_row:
            for p in tail_pls + atop_pls:
                y0    = p.get("y_start_cm", 0)
                lp    = p.get("length_cm", 0)
                is_np = p.get("type") == "np_box"
                color = _REC_NP_COLOR if is_np else _REC_PAL_COLOR
                units = p.get("units_per_placement", p.get("pallets_per_block", 0))
                lbl   = f"+{units}{'b' if is_np else 'p'}"
                _color_layout_range(ws, rec_row, y0, lp, color, lbl)

        # Spacer row
        spacer_row = pallet_row + n_sub_rows
        ws.row_dimensions[spacer_row].height = ROW_H_SPACER
        cur_row = spacer_row + 1


# ═══════════════════════════════════════════════════════════════════════════
# Sheet 4 — Recommendations
# ═══════════════════════════════════════════════════════════════════════════

def _write_recommendations(ws, recs):
    ws.sheet_view.showGridLines = False

    col_widths = [10, 20, 38, 14, 14, 10, 14, 14, 14]
    for i, w in enumerate(col_widths):
        ws.column_dimensions[get_column_letter(i + 1)].width = w

    _merge_write(ws, 1, 1, 1, 9,
                 "  FILL RECOMMENDATIONS — additional items to order to fill free container space",
                 bold=True, fg=_WHITE, bg=_HEADER_BG, size=12)
    ws.row_dimensions[1].height = 28

    _merge_write(ws, 2, 1, 2, 9,
                 "  ✅ Tail Zone = unused length after last pallet block   "
                 "  ✅ Atop Zone = headroom above existing pallet rows",
                 fg="444444", size=9, bg=_LIGHT_GRAY)
    ws.row_dimensions[2].height = 16

    row = 4

    for rec in recs:
        idx      = rec["container_index"]
        before   = rec.get("leftover_before_cm", 0)
        after    = rec.get("leftover_after_cm", 0)
        rate     = rec.get("fill_rate_pct", 0.0)
        n_pal    = rec.get("total_pallets_to_add", 0)
        n_np     = rec.get("total_np_boxes_to_add", 0)
        ts       = rec.get("tail_summary",  {})
        as_      = rec.get("atop_summary",  {})

        has_any  = bool(rec.get("tail_placements") or rec.get("atop_placements"))

        # Container header
        fill_msg = (f"Tail: {before} cm → {after} cm  ({rate}% filled)"
                    if before > 0 else "Tail: already full")
        _merge_write(ws, row, 1, row, 9,
                     f"  Container {idx}   —   {fill_msg}   |   "
                     f"+{n_pal} pallets  +{n_np} NP boxes",
                     bold=True, fg=_WHITE, bg=_SUBHDR_BG, size=10)
        ws.row_dimensions[row].height = 20
        row += 1

        if not has_any:
            _merge_write(ws, row, 1, row, 9,
                         "    No additions possible — no items from current order fit remaining free zones",
                         fg="777777", italic=True, bg=_LIGHT_GRAY, size=9)
            ws.row_dimensions[row].height = 16
            row += 1
            ws.row_dimensions[row].height = 8
            row += 1
            continue

        # Column headers
        _col_header_row(ws, row,
                        ["Zone", "Type", "Product / Block", "Length (cm)", "Height (cm)",
                         "Count", "Units/Row", "Total Units", "Est. FOB"],
                        col_start=1)
        row += 1

        def _write_pallet_rows(blocks, zone_label):
            nonlocal row
            for ri, b in enumerate(blocks):
                bg = _ALT_ROW if ri % 2 == 0 else _WHITE
                fob_s = f"{b['est_value_fob']:,.2f}" if b.get("est_value_fob") else "—"
                # "115x77|>130  (L=77cm, H=144cm, 2 pal/block)"
                block_label = (
                    f"{b['block_type_key']}"
                    f"  —  {b['pallets_per_block']} pallets/block"
                )
                vals = [
                    zone_label,
                    "Pallet block",
                    block_label,
                    b["length_cm"],
                    b["height_cm"],
                    b["count"],
                    b["pallets_per_block"],
                    b["total_pallets"],
                    fob_s,
                ]
                for ci, v in enumerate(vals):
                    color = _block_color(b["block_type_key"]) if ci == 2 else \
                            _REC_PAL_COLOR if ci == 0 else bg
                    _set_cell(ws, row, ci + 1, v, bg=color, border=True,
                              h="center" if ci > 2 else "left", wrap=(ci == 2))
                ws.row_dimensions[row].height = 18
                row += 1

        def _write_np_rows(np_rows, zone_label):
            nonlocal row
            for ri, b in enumerate(np_rows):
                bg = _ALT_ROW if ri % 2 == 0 else _WHITE
                # Show product name + dimension key: "Bravo bowl... (35×55×26cm)"
                dim = b.get("dim", "")
                product_label = f"{b['label']}  [{dim}]" if dim else b["label"]
                vals = [
                    zone_label,
                    f"NP box  ({b['n_across']}× across width)",
                    product_label,
                    b["length_cm"],
                    b["height_cm"],
                    b["count"],
                    b["n_across"],
                    b["total_boxes"],
                    "—",
                ]
                for ci, v in enumerate(vals):
                    color = _NP_BOX_COLOR if ci == 2 else \
                            _REC_NP_COLOR if ci == 0 else bg
                    _set_cell(ws, row, ci + 1, v, bg=color, border=True,
                              h="center" if ci > 2 else "left", wrap=(ci == 2))
                ws.row_dimensions[row].height = 28   # taller for wrapped text
                row += 1

        if ts.get("pallet_blocks") or ts.get("np_box_rows"):
            _write_pallet_rows(ts.get("pallet_blocks", []), "TAIL")
            _write_np_rows(ts.get("np_box_rows", []), "TAIL")

        if as_.get("pallet_blocks") or as_.get("np_box_rows"):
            _write_pallet_rows(as_.get("pallet_blocks", []), "ATOP")
            _write_np_rows(as_.get("np_box_rows", []), "ATOP")

        # Totals row
        _merge_write(ws, row, 1, row, 7,
                     f"   Total additions for Container {idx}",
                     bold=True, bg=_ACCENT_BG)
        _set_cell(ws, row, 8, n_pal, bold=True, bg=_ACCENT_BG, h="center")
        _set_cell(ws, row, 9, f"+{n_np} boxes", bold=True, bg=_ACCENT_BG, h="center")
        ws.row_dimensions[row].height = 18
        row += 1

        ws.row_dimensions[row].height = 8   # spacer
        row += 1


# ═══════════════════════════════════════════════════════════════════════════
# Public API
# ═══════════════════════════════════════════════════════════════════════════

def export_excel_report(
    containers: List[Dict[str, Any]],
    recs: List[Dict[str, Any]],
    np_boxes: Optional[List[Dict[str, Any]]],
    unplaced: Optional[List[Dict[str, Any]]],
    out_dir: "Path",
    config: Dict[str, Any],
) -> Optional["Path"]:
    """
    Generate outputs/report.xlsx.

    Parameters
    ----------
    containers : solved container list from main pipeline
    recs       : recommendation list from recommend_fill_containers
    np_boxes   : NP box type list (may be None)
    unplaced   : unplaced NP boxes list (may be None)
    out_dir    : output directory Path
    config     : dict with container config constants

    Returns path to the written file, or None if openpyxl is unavailable.
    """
    if not _HAS_OPENPYXL:
        print("[export_excel] openpyxl not installed — skipping Excel report")
        return None

    wb = Workbook()
    wb.remove(wb.active)   # delete default blank sheet

    ws1 = wb.create_sheet("Overview")
    ws2 = wb.create_sheet("Container Details")
    ws3 = wb.create_sheet("Packing Layout")
    ws4 = wb.create_sheet("Recommendations")

    # Tab colours
    ws1.sheet_properties.tabColor = "1A3A5C"
    ws2.sheet_properties.tabColor = "2E6DA4"
    ws3.sheet_properties.tabColor = "27AE60"
    ws4.sheet_properties.tabColor = "E67E22"

    _write_overview(ws1, containers, recs, np_boxes or [], unplaced or [], config)
    _write_details(ws2, containers, config)
    _write_layout(ws3, containers, recs, config)
    _write_recommendations(ws4, recs)

    out_path = out_dir / "report.xlsx"
    wb.save(str(out_path))
    return out_path
