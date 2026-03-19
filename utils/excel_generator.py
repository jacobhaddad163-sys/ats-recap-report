"""
Excel output generator for ATS Recap Report.

Produces formatted .xlsx with:
  - Detail sheets (yellow category headers, grey sub-headers, summaries, images)
  - RECAP tab (brand summaries, size totals, grand total)
"""

import io
import logging
import re
from collections import OrderedDict
from datetime import date
from typing import Dict, List, Optional

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XlImage
from openpyxl.styles import (
    Alignment, Border, Font, PatternFill, Side,
)
from openpyxl.utils import get_column_letter

from .ats_parser import get_recap_data

logger = logging.getLogger(__name__)

_FORMULA_STARTERS = ('=', '+', '-', '@', '\t', '\r', '\n')

def _safe_cell_text(value: str) -> str:
    if not value:
        return value
    value = str(value).strip()
    value = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', value)
    if value and value[0] in _FORMULA_STARTERS:
        value = "'" + value
    return value

# ─── Style Constants ─────────────────────────────────────────────────────────
YELLOW_FILL = PatternFill("solid", fgColor="FFFF00")
GREY_FILL = PatternFill("solid", fgColor="D3D3D3")
WHITE_FILL = PatternFill("solid", fgColor="FFFFFF")
LIGHT_BLUE_FILL = PatternFill("solid", fgColor="C0E6F5")

BOLD_FONT = Font(name="Aptos Narrow", bold=True, size=11)
NORMAL_FONT = Font(name="Aptos Narrow", size=11)

THIN_BORDER = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin'),
)

CENTER_ALIGN = Alignment(horizontal='center', vertical='center')
LEFT_ALIGN = Alignment(horizontal='left', vertical='center')
RIGHT_ALIGN = Alignment(horizontal='right', vertical='center')
WRAP_CENTER_ALIGN = Alignment(horizontal='center', vertical='center', wrap_text=True)

NUM_FMT = '#,##0'
ACCT_NUM_FMT = '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
PRICE_FMT = '#,##0.00'
TEXT_FMT = '@'


# ─── Detail Sheet ────────────────────────────────────────────────────────────

def _set_detail_col_widths(ws):
    widths = {'A': 28, 'B': 6, 'C': 22, 'D': 28,
              'E': 6, 'F': 6, 'G': 6, 'H': 6, 'I': 6, 'J': 6,
              'K': 10, 'L': 12, 'M': 10, 'N': 14, 'O': 10}
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width


def _write_sheet_header(ws, report_date: date = None):
    if report_date is None:
        report_date = date.today()

    ws.merge_cells('A7:H7')
    cell = ws.cell(row=7, column=1, value="ATS RECAP")
    cell.font = Font(name="Aptos Narrow", bold=True, size=14)
    cell.alignment = LEFT_ALIGN

    ws.merge_cells('A8:H8')
    cell = ws.cell(row=8, column=1, value=report_date)
    cell.font = Font(name="Aptos Narrow", bold=True, size=11)
    cell.alignment = LEFT_ALIGN
    cell.number_format = 'M/D/YYYY'


def _write_category_summary(ws, row: int, label: str, oh: int, wip: int,
                             is_category_row: bool = False, category_name: str = ""):
    """Write a summary row (TODDLER or 4-7 line)."""
    # Column K: size range label
    cell = ws.cell(row=row, column=11, value=label)
    cell.font = BOLD_FONT
    cell.alignment = CENTER_ALIGN
    cell.border = THIN_BORDER
    # "4-7" needs text format to prevent Excel date interpretation
    if label == "4-7":
        cell.number_format = TEXT_FMT

    # Column L: OH — accounting format shows dash for zero
    cell = ws.cell(row=row, column=12, value=oh)
    cell.font = BOLD_FONT
    cell.alignment = CENTER_ALIGN
    cell.number_format = ACCT_NUM_FMT
    cell.border = THIN_BORDER

    # Column M: WIP — accounting format shows dash for zero
    cell = ws.cell(row=row, column=13, value=wip)
    cell.font = BOLD_FONT
    cell.alignment = CENTER_ALIGN
    cell.number_format = ACCT_NUM_FMT
    cell.border = THIN_BORDER

    # Column N: Total
    cell = ws.cell(row=row, column=14, value=oh + wip)
    cell.font = BOLD_FONT
    cell.alignment = CENTER_ALIGN
    cell.number_format = ACCT_NUM_FMT
    cell.border = THIN_BORDER

    # Category name in column A (yellow fill, bold)
    if is_category_row and category_name:
        cell = ws.cell(row=row, column=1, value=_safe_cell_text(category_name))
        cell.fill = YELLOW_FILL
        cell.font = BOLD_FONT
        cell.alignment = LEFT_ALIGN

    # Bottom border on C:J for the 4-7 row (category boundary)
    if label == "4-7":
        bottom_border = Border(bottom=Side(style='thin'))
        for col in range(3, 11):  # C through J
            c = ws.cell(row=row, column=col)
            if c.border and c.border != Border():
                # Preserve existing borders, add bottom
                c.border = Border(
                    left=c.border.left, right=c.border.right,
                    top=c.border.top, bottom=Side(style='thin'))
            else:
                c.border = bottom_border


def _write_block_header(ws, row: int):
    headers = {3: "STYLE", 4: "COLOR", 5: "SIZE SCALE",
               12: "ON HAND", 13: "WIP", 14: "AVAILABILITY", 15: "MSRP"}
    ws.merge_cells(f'E{row}:K{row}')
    for col, label in headers.items():
        cell = ws.cell(row=row, column=col, value=label)
        cell.fill = GREY_FILL
        cell.font = NORMAL_FONT
        cell.alignment = CENTER_ALIGN
    for col in range(1, 16):
        ws.cell(row=row, column=col).fill = GREY_FILL


def _write_data_rows(ws, start_row: int, rows_data: list) -> int:
    current_row = start_row
    for row_data in rows_data:
        ws.cell(row=current_row, column=3, value=row_data["style_num"]).font = NORMAL_FONT
        ws.cell(row=current_row, column=4, value=row_data["color"]).font = NORMAL_FONT

        for col_idx in range(5, 12):
            val = row_data["cells"].get(col_idx)
            if val is not None:
                cell = ws.cell(row=current_row, column=col_idx, value=val)
                cell.font = NORMAL_FONT
                cell.alignment = CENTER_ALIGN

        if row_data.get("is_label_row", True) and row_data["oh"] > 0:
            cell = ws.cell(row=current_row, column=12, value=row_data["oh"])
            cell.font = NORMAL_FONT
            cell.alignment = CENTER_ALIGN
            cell.number_format = NUM_FMT

        if row_data.get("is_label_row", True) and row_data["wip"] > 0:
            cell = ws.cell(row=current_row, column=13, value=row_data["wip"])
            cell.font = NORMAL_FONT
            cell.alignment = CENTER_ALIGN
            cell.number_format = NUM_FMT

        avail = row_data.get("availability", "")
        if avail:
            ws.cell(row=current_row, column=14, value=avail).font = NORMAL_FONT

        msrp = row_data.get("msrp", 0)
        if msrp > 0:
            cell = ws.cell(row=current_row, column=15, value=msrp)
            cell.font = NORMAL_FONT
            cell.alignment = CENTER_ALIGN
            cell.number_format = PRICE_FMT

        current_row += 1
    return current_row


def _write_total_row(ws, row: int, total_oh: int, total_wip: int):
    """Write TOTAL row with OH, WIP, and combined total in N."""
    cell = ws.cell(row=row, column=3, value="TOTAL :")
    cell.fill = GREY_FILL
    cell.font = BOLD_FONT

    for col in range(1, 16):
        ws.cell(row=row, column=col).fill = GREY_FILL

    # OH
    cell = ws.cell(row=row, column=12, value=total_oh)
    cell.fill = GREY_FILL
    cell.font = BOLD_FONT
    cell.alignment = CENTER_ALIGN
    cell.number_format = NUM_FMT

    # WIP
    if total_wip > 0:
        cell = ws.cell(row=row, column=13, value=total_wip)
        cell.fill = GREY_FILL
        cell.font = BOLD_FONT
        cell.alignment = CENTER_ALIGN
        cell.number_format = NUM_FMT


def _add_product_image(ws, row: int, img_bytes: bytes):
    try:
        img = XlImage(io.BytesIO(img_bytes))
        img.width = 120
        img.height = 120
        ws.add_image(img, f'A{row}')
    except Exception:
        pass


def write_detail_sheet(ws, categories: list, report_date: date = None):
    """
    Write a complete detail sheet.
    categories: list of category dicts from parser (with toddler_oh, boys47_oh, blocks, etc.)
    """
    _set_detail_col_widths(ws)
    _write_sheet_header(ws, report_date)
    current_row = 10

    for cat in categories:
        cat_name = cat["name"]
        tod_oh, tod_wip = cat["toddler_oh"], cat["toddler_wip"]
        b47_oh, b47_wip = cat["boys47_oh"], cat["boys47_wip"]
        blocks = cat["blocks"]

        has_toddler = (tod_oh > 0 or tod_wip > 0)
        has_boys47 = (b47_oh > 0 or b47_wip > 0)
        if not has_toddler and not has_boys47:
            continue

        # OH/WIP/TOTAL column sub-headers (yellow, in K/L/M per spec)
        for col, label in [(11, "OH"), (12, "WIP"), (13, "TOTAL")]:
            cell = ws.cell(row=current_row, column=col, value=label)
            cell.fill = YELLOW_FILL
            cell.font = BOLD_FONT
            cell.alignment = CENTER_ALIGN
            cell.border = THIN_BORDER
        # Col N also gets yellow fill and border (no label)
        ws.cell(row=current_row, column=14).fill = YELLOW_FILL
        ws.cell(row=current_row, column=14).border = THIN_BORDER
        current_row += 1

        # TODDLER summary row
        if has_toddler:
            _write_category_summary(ws, current_row, "TODDLER", tod_oh, tod_wip,
                                     is_category_row=not has_boys47, category_name=cat_name)
            current_row += 1

        # 4-7 summary row (category name goes here if both exist)
        if has_boys47:
            _write_category_summary(ws, current_row, "4-7", b47_oh, b47_wip,
                                     is_category_row=True, category_name=cat_name)
            current_row += 1

        # Write all style blocks for this category
        for block in blocks:
            _write_block_header(ws, current_row)
            current_row += 1
            current_row = _write_data_rows(ws, current_row, block["rows"])
            _write_total_row(ws, current_row, block["total_oh"], block["total_wip"])
            current_row += 1
            if block.get("product_image"):
                _add_product_image(ws, current_row, block["product_image"])
            current_row += 8

        current_row += 2


# ─── RECAP Sheet ─────────────────────────────────────────────────────────────

def _set_recap_col_widths(ws):
    widths = {'A': 20.5, 'B': 15.3, 'C': 26.0, 'D': 35.2, 'E': 9.9, 'F': 9.9, 'G': 9.9}
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width


def write_recap_sheet(ws, recap_sections: list, title: str = ""):
    _set_recap_col_widths(ws)

    # Collect merge ranges — apply them at the very end so all cell styling
    # happens on real Cell objects (not MergedCell proxies).
    pending_merges = []

    # Row 1: Title — style ALL cells in the range before merging
    for col in range(1, 8):
        c = ws.cell(row=1, column=col)
        c.fill = YELLOW_FILL
        c.font = BOLD_FONT
        c.alignment = CENTER_ALIGN
        c.border = THIN_BORDER
    ws.cell(row=1, column=1, value=_safe_cell_text(title.upper() if title else "ATS RECAP"))
    pending_merges.append('A1:G1')

    # Row 2: Headers
    headers = ["BRAND", "SIZE RANGE", "CATEGORY", "REF #", "OH", "WIP", "TOTAL ATS"]
    for col, label in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col, value=label)
        cell.fill = YELLOW_FILL
        cell.font = BOLD_FONT
        cell.alignment = CENTER_ALIGN
        cell.border = THIN_BORDER

    current_row = 3
    all_data_rows = []
    brand_total_rows = []

    for section in recap_sections:
        brand_label = section["brand_label"]
        rows = section["rows"]
        if not rows:
            continue

        section_start_row = current_row

        # Group rows by category for merging.
        # Use cat_id (unique per category entry) to keep same-named categories
        # separate — spec says duplicate category names are SEPARATE entries.
        cat_groups = []
        current_cat_id = None
        for row_data in rows:
            cat_id = row_data.get("cat_id")
            if cat_id != current_cat_id:
                cat_groups.append([row_data])
                current_cat_id = cat_id
            else:
                cat_groups[-1].append(row_data)

        for cat_group in cat_groups:
            cat_start_row = current_row

            for row_data in cat_group:
                # Apply white fill and borders to all cells in this data row
                for col in range(1, 8):
                    c = ws.cell(row=current_row, column=col)
                    c.fill = WHITE_FILL
                    c.border = THIN_BORDER

                # B: SIZE RANGE
                cell_b = ws.cell(row=current_row, column=2, value=row_data["size_range"])
                cell_b.font = NORMAL_FONT
                cell_b.alignment = CENTER_ALIGN

                # C: Category name (set on every row — merge later collapses to anchor)
                ws.cell(row=current_row, column=3).font = NORMAL_FONT
                ws.cell(row=current_row, column=3).alignment = WRAP_CENTER_ALIGN

                # D: REF # — text format (@) with wrap_text
                cell_d = ws.cell(row=current_row, column=4,
                                 value=_safe_cell_text(row_data["ref_nums"]))
                cell_d.font = NORMAL_FONT
                cell_d.alignment = WRAP_CENTER_ALIGN
                cell_d.number_format = TEXT_FMT

                # E: OH
                cell = ws.cell(row=current_row, column=5, value=row_data["oh"])
                cell.font = NORMAL_FONT
                cell.alignment = CENTER_ALIGN
                cell.number_format = NUM_FMT

                # F: WIP
                cell = ws.cell(row=current_row, column=6, value=row_data["wip"])
                cell.font = NORMAL_FONT
                cell.alignment = CENTER_ALIGN
                cell.number_format = NUM_FMT

                # G: TOTAL ATS (formula)
                cell = ws.cell(row=current_row, column=7, value=f'=E{current_row}+F{current_row}')
                cell.font = NORMAL_FONT
                cell.alignment = CENTER_ALIGN
                cell.number_format = NUM_FMT

                all_data_rows.append(current_row)
                current_row += 1

            # Category name value (on anchor cell)
            ws.cell(row=cat_start_row, column=3,
                    value=_safe_cell_text(cat_group[0]["category"]))

            # Queue category merge
            if len(cat_group) > 1:
                pending_merges.append(f'C{cat_start_row}:C{current_row - 1}')

        # Brand name value (on anchor cell)
        if current_row > section_start_row:
            cell = ws.cell(row=section_start_row, column=1, value=_safe_cell_text(brand_label))
            cell.font = BOLD_FONT
            cell.alignment = WRAP_CENTER_ALIGN
            # Queue brand merge
            if current_row - section_start_row > 1:
                pending_merges.append(f'A{section_start_row}:A{current_row - 1}')

        # Brand total row (light blue) — style ALL cells A:G before merging
        for col in range(1, 8):
            c = ws.cell(row=current_row, column=col)
            c.fill = LIGHT_BLUE_FILL
            c.border = THIN_BORDER

        ws.cell(row=current_row, column=1,
                value=_safe_cell_text(f"{brand_label} TOTAL:"))
        ws.cell(row=current_row, column=1).font = BOLD_FONT
        ws.cell(row=current_row, column=1).alignment = CENTER_ALIGN

        for col_idx, formula in [(5, f'=SUM(E{section_start_row}:E{current_row - 1})'),
                                  (6, f'=SUM(F{section_start_row}:F{current_row - 1})'),
                                  (7, f'=E{current_row}+F{current_row}')]:
            cell = ws.cell(row=current_row, column=col_idx, value=formula)
            cell.font = BOLD_FONT
            cell.alignment = CENTER_ALIGN
            cell.number_format = NUM_FMT

        pending_merges.append(f'A{current_row}:D{current_row}')
        brand_total_rows.append(current_row)
        current_row += 1

    # Size range totals (TODDLER TOTAL, 4-7 TOTAL, etc.)
    unique_srs = []
    for section in recap_sections:
        for row_data in section["rows"]:
            if row_data["size_range"] not in unique_srs:
                unique_srs.append(row_data["size_range"])

    sr_sort_order = ["TODDLER", "BOYS 4-7", "NEWBORN", "INFANT", "4-6X", "7-16", "8-20"]
    unique_srs.sort(key=lambda x: sr_sort_order.index(x) if x in sr_sort_order else 99)

    first_data = min(all_data_rows) if all_data_rows else 3
    last_data = max(all_data_rows) if all_data_rows else 3

    for sr_name in unique_srs:
        # Style ALL cells A:G before merging
        for col in range(1, 8):
            c = ws.cell(row=current_row, column=col)
            c.fill = LIGHT_BLUE_FILL
            c.border = THIN_BORDER

        ws.cell(row=current_row, column=1, value=_safe_cell_text(f"{sr_name} TOTAL"))
        ws.cell(row=current_row, column=1).font = BOLD_FONT
        ws.cell(row=current_row, column=1).alignment = CENTER_ALIGN

        safe_sr = re.sub(r'["\';=+@]', '', sr_name)

        for col_idx, formula in [
            (5, f'=SUMIF($B${first_data}:$B${last_data},"{safe_sr}",E${first_data}:E${last_data})'),
            (6, f'=SUMIF($B${first_data}:$B${last_data},"{safe_sr}",F${first_data}:F${last_data})'),
            (7, f'=E{current_row}+F{current_row}'),
        ]:
            cell = ws.cell(row=current_row, column=col_idx, value=formula)
            cell.font = BOLD_FONT
            cell.alignment = CENTER_ALIGN
            cell.number_format = NUM_FMT

        pending_merges.append(f'A{current_row}:D{current_row}')
        current_row += 1

    # Grand Total (yellow) — style ALL cells before merging
    for col in range(1, 8):
        c = ws.cell(row=current_row, column=col)
        c.fill = YELLOW_FILL
        c.border = THIN_BORDER

    ws.cell(row=current_row, column=1, value="GRAND TOTAL:")
    ws.cell(row=current_row, column=1).font = BOLD_FONT
    ws.cell(row=current_row, column=1).alignment = CENTER_ALIGN

    oh_formula = "+".join(f"E{r}" for r in brand_total_rows) if brand_total_rows else "0"
    wip_formula = "+".join(f"F{r}" for r in brand_total_rows) if brand_total_rows else "0"

    for col_idx, formula in [(5, f'={oh_formula}'), (6, f'={wip_formula}'),
                              (7, f'=E{current_row}+F{current_row}')]:
        cell = ws.cell(row=current_row, column=col_idx, value=formula)
        cell.font = BOLD_FONT
        cell.alignment = CENTER_ALIGN
        cell.number_format = NUM_FMT

    pending_merges.append(f'A{current_row}:D{current_row}')

    # NOW apply all merges — after all cell styling is complete
    for merge_range in pending_merges:
        ws.merge_cells(merge_range)


# ─── Main Generator ─────────────────────────────────────────────────────────

def generate_ats_report(categories_by_sheet: Dict[str, dict],
                        title: str = "", report_date: date = None) -> bytes:
    """
    Generate the complete ATS report Excel file.

    categories_by_sheet: {
        "NIKE LONG BOTTOMS 2-7": {
            "brand": "NIKE",
            "general_category": "LONG BOTTOMS",
            "categories": OrderedDict of {cat_name: {"blocks": [...], "size_ranges": {...}}},
        },
        ...
    }
    """
    wb = Workbook()
    ws_recap = wb.active
    ws_recap.title = "RECAP SHEET"

    recap_sections = get_recap_data(categories_by_sheet)
    write_recap_sheet(ws_recap, recap_sections, title=title)

    for sheet_name, sheet_info in categories_by_sheet.items():
        ws = wb.create_sheet(title=sheet_name[:31])
        write_detail_sheet(ws, sheet_info["categories"], report_date=report_date)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()
