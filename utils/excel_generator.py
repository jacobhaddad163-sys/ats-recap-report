"""
Excel output generator for ATS Recap Report.

Produces formatted .xlsx with:
  - Detail sheets matching exact Haddad ATS format
    (yellow category headers, grey sub-headers, product images,
     TODDLER/4-7 summary rows, TOTAL rows per ref# block)
  - RECAP tab (brand summaries, size totals, grand total)
"""

import io
import logging
import re
from datetime import date
from typing import Dict, List, Optional

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XlImage
from openpyxl.styles import (
    Alignment, Border, Font, PatternFill, Side,
)
from openpyxl.utils import get_column_letter

logger = logging.getLogger(__name__)

# Formula injection characters to sanitize in cell text values
_FORMULA_STARTERS = ('=', '+', '-', '@', '\t', '\r', '\n')


def _safe_cell_text(value: str) -> str:
    """Sanitize text for Excel cells to prevent formula injection."""
    if not value:
        return value
    value = str(value).strip()
    # Remove control characters
    value = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', value)
    # If starts with formula character, prefix with apostrophe
    if value and value[0] in _FORMULA_STARTERS:
        value = "'" + value
    return value

from .ats_parser import get_recap_data


# ─── Style Constants ─────────────────────────────────────────────────────────

YELLOW_FILL = PatternFill("solid", fgColor="FFFF00")
GREY_FILL = PatternFill("solid", fgColor="D3D3D3")
WHITE_FILL = PatternFill("solid", fgColor="FFFFFF")
LIGHT_BLUE_FILL = PatternFill("solid", fgColor="C0E6F5")

BOLD_FONT = Font(name="Aptos Narrow", bold=True, size=11)
NORMAL_FONT = Font(name="Aptos Narrow", size=11)
BOLD_FONT_SM = Font(name="Aptos Narrow", bold=True, size=10)
NORMAL_FONT_SM = Font(name="Aptos Narrow", size=10)

THIN_BORDER = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin'),
)

CENTER_ALIGN = Alignment(horizontal='center', vertical='center')
LEFT_ALIGN = Alignment(horizontal='left', vertical='center')
RIGHT_ALIGN = Alignment(horizontal='right', vertical='center')

# Number format for quantities (comma thousands separator)
NUM_FMT = '#,##0'
PRICE_FMT = '#,##0.00'


# ─── Detail Sheet Generator ─────────────────────────────────────────────────

def _set_detail_col_widths(ws):
    """Set column widths for detail sheets to match the Haddad format."""
    widths = {
        'A': 28,    # Category name / image area
        'B': 6,     # Spacer / image area
        'C': 22,    # STYLE
        'D': 28,    # COLOR
        'E': 6,     # Size 2T / 4
        'F': 6,     # Size 3T / 5
        'G': 6,     # Size 4T / 6
        'H': 6,     # Size 4 / 7
        'I': 6,     # Size 5
        'J': 6,     # Size 6
        'K': 10,    # Size 7 / TODDLER/4-7 label
        'L': 12,    # ON HAND
        'M': 10,    # WIP
        'N': 14,    # AVAILABILITY / TOTAL
        'O': 10,    # MSRP
    }
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width


def _write_sheet_header(ws, report_date: date = None):
    """Write the ATS RECAP header area (rows 1-9)."""
    if report_date is None:
        report_date = date.today()

    # Row 1-6: Empty (Haddad logo area)
    # Row 7: "ATS RECAP" (bold, merged A7:H7)
    ws.merge_cells('A7:H7')
    cell = ws.cell(row=7, column=1, value="ATS RECAP")
    cell.font = Font(name="Aptos Narrow", bold=True, size=14)
    cell.alignment = LEFT_ALIGN

    # Row 8: Date (merged A8:H8)
    ws.merge_cells('A8:H8')
    cell = ws.cell(row=8, column=1, value=report_date)
    cell.font = Font(name="Aptos Narrow", bold=True, size=11)
    cell.alignment = LEFT_ALIGN
    cell.number_format = 'M/D/YYYY'


def _write_category_summary_headers(ws, row: int):
    """Write the OH / WIP / TOTAL column headers (yellow fill, bold)."""
    for col, label in [(12, "OH"), (13, "WIP"), (14, "TOTAL")]:
        cell = ws.cell(row=row, column=col, value=label)
        cell.fill = YELLOW_FILL
        cell.font = BOLD_FONT
        cell.alignment = CENTER_ALIGN


def _write_category_summary_row(ws, row: int, label: str, oh: int, wip: int,
                                 is_category_row: bool = False, category_name: str = ""):
    """
    Write a summary row (TODDLER or 4-7 line).
    If is_category_row=True, also writes the category name in column A.
    """
    # Column K: size range label
    cell = ws.cell(row=row, column=11, value=label)
    cell.font = BOLD_FONT
    cell.alignment = RIGHT_ALIGN

    # Column L: OH
    cell = ws.cell(row=row, column=12, value=oh)
    cell.font = BOLD_FONT
    cell.alignment = CENTER_ALIGN
    cell.number_format = NUM_FMT

    # Column M: WIP
    wip_display = wip if wip > 0 else "-"
    cell = ws.cell(row=row, column=13, value=wip_display)
    cell.font = BOLD_FONT
    cell.alignment = CENTER_ALIGN
    if isinstance(wip_display, int):
        cell.number_format = NUM_FMT

    # Column N: Total
    total = oh + wip
    cell = ws.cell(row=row, column=14, value=total)
    cell.font = BOLD_FONT
    cell.alignment = CENTER_ALIGN
    cell.number_format = NUM_FMT

    # Category name in column A (yellow fill, bold)
    if is_category_row and category_name:
        cell = ws.cell(row=row, column=1, value=_safe_cell_text(category_name))
        cell.fill = YELLOW_FILL
        cell.font = BOLD_FONT
        cell.alignment = LEFT_ALIGN


def _write_block_header(ws, row: int):
    """Write the grey header row for a ref# block."""
    headers = {
        3: "STYLE",
        4: "COLOR",
        5: "SIZE SCALE",  # This gets merged across E:K
        12: "ON HAND",
        13: "WIP",
        14: "AVAILABILITY",
        15: "MSRP",
    }

    # Merge SIZE SCALE across E:K
    ws.merge_cells(f'E{row}:K{row}')

    for col, label in headers.items():
        cell = ws.cell(row=row, column=col, value=label)
        cell.fill = GREY_FILL
        cell.font = NORMAL_FONT
        cell.alignment = CENTER_ALIGN

    # Apply grey fill to all columns
    for col in range(1, 16):
        ws.cell(row=row, column=col).fill = GREY_FILL


def _write_data_rows(ws, start_row: int, rows_data: list) -> int:
    """
    Write style data rows to the sheet.
    Returns the next available row.
    """
    current_row = start_row

    for row_data in rows_data:
        # Style number
        ws.cell(row=current_row, column=3, value=row_data["style_num"]).font = NORMAL_FONT

        # Color
        ws.cell(row=current_row, column=4, value=row_data["color"]).font = NORMAL_FONT

        # Size columns (E through K) - preserve original cell values
        for col_idx in range(5, 12):
            val = row_data["cells"].get(col_idx)
            if val is not None:
                cell = ws.cell(row=current_row, column=col_idx, value=val)
                cell.font = NORMAL_FONT
                cell.alignment = CENTER_ALIGN

        # OH (only on label rows)
        if row_data.get("is_label_row", True) and row_data["oh"] > 0:
            cell = ws.cell(row=current_row, column=12, value=row_data["oh"])
            cell.font = NORMAL_FONT
            cell.alignment = CENTER_ALIGN
            cell.number_format = NUM_FMT

        # WIP
        if row_data.get("is_label_row", True) and row_data["wip"] > 0:
            cell = ws.cell(row=current_row, column=13, value=row_data["wip"])
            cell.font = NORMAL_FONT
            cell.alignment = CENTER_ALIGN
            cell.number_format = NUM_FMT

        # Availability
        avail = row_data.get("availability", "")
        if avail:
            cell = ws.cell(row=current_row, column=14, value=avail)
            cell.font = NORMAL_FONT
            cell.alignment = CENTER_ALIGN

        # MSRP
        msrp = row_data.get("msrp", 0)
        if msrp > 0:
            cell = ws.cell(row=current_row, column=15, value=msrp)
            cell.font = NORMAL_FONT
            cell.alignment = CENTER_ALIGN
            cell.number_format = PRICE_FMT

        current_row += 1

    return current_row


def _write_total_row(ws, row: int, total_oh: int, total_wip: int):
    """Write a grey TOTAL row for a ref# block."""
    cell = ws.cell(row=row, column=3, value="TOTAL :")
    cell.fill = GREY_FILL
    cell.font = BOLD_FONT

    # Apply grey fill to all columns
    for col in range(1, 16):
        ws.cell(row=row, column=col).fill = GREY_FILL

    # OH total
    cell = ws.cell(row=row, column=12, value=total_oh)
    cell.fill = GREY_FILL
    cell.font = BOLD_FONT
    cell.alignment = CENTER_ALIGN
    cell.number_format = NUM_FMT

    # WIP total
    if total_wip > 0:
        cell = ws.cell(row=row, column=13, value=total_wip)
        cell.fill = GREY_FILL
        cell.font = BOLD_FONT
        cell.alignment = CENTER_ALIGN
        cell.number_format = NUM_FMT


def _add_product_image(ws, row: int, img_bytes: bytes):
    """Add a product image to the sheet at the specified row."""
    try:
        img = XlImage(io.BytesIO(img_bytes))
        # Scale to fit (roughly 6 rows tall, columns A-B wide)
        img.width = 120
        img.height = 120
        ws.add_image(img, f'A{row}')
    except Exception:
        pass  # Images are nice-to-have, never break on them


def write_detail_sheet(ws, categories: list, report_date: date = None):
    """
    Write a complete detail sheet with all categories, blocks, and formatting.

    categories: list of category dicts from group_blocks_by_category()
    """
    _set_detail_col_widths(ws)
    _write_sheet_header(ws, report_date)

    current_row = 10  # Start after header area

    for cat in categories:
        cat_name = cat["name"]
        size_ranges = cat["size_ranges"]
        blocks = cat["blocks"]

        # Determine size range order and which exist
        sr_order = []
        for sr_name in ["TODDLER", "BOYS 4-7", "NEWBORN", "INFANT", "4-6X", "7-16", "8-20"]:
            if sr_name in size_ranges and (size_ranges[sr_name]["oh"] > 0 or size_ranges[sr_name]["wip"] > 0):
                sr_order.append(sr_name)
        for sr_name in size_ranges:
            if sr_name not in sr_order and (size_ranges[sr_name]["oh"] > 0 or size_ranges[sr_name]["wip"] > 0):
                sr_order.append(sr_name)

        if not sr_order:
            continue

        # Write OH/WIP/TOTAL headers
        _write_category_summary_headers(ws, current_row)
        current_row += 1

        # Write summary rows for each size range
        for idx, sr_name in enumerate(sr_order):
            sr_data = size_ranges[sr_name]
            is_cat_row = (idx == len(sr_order) - 1)  # Category name goes on last summary row
            # If only one size range, put category on that row
            # If two, put category on the second (4-7) row per the example format
            _write_category_summary_row(
                ws, current_row, sr_name,
                sr_data["oh"], sr_data["wip"],
                is_category_row=is_cat_row,
                category_name=cat_name,
            )
            current_row += 1

        # Write ref# blocks
        for block in blocks:
            # Grey header row
            _write_block_header(ws, current_row)
            current_row += 1

            # Data rows
            current_row = _write_data_rows(ws, current_row, block["rows"])

            # TOTAL row
            _write_total_row(ws, current_row, block["total_oh"], block["total_wip"])
            current_row += 1

            # Product image (in empty rows after total)
            if block.get("product_image"):
                _add_product_image(ws, current_row, block["product_image"])

            # Empty rows for image area
            current_row += 8  # Space for product image

        # Extra spacing between categories
        current_row += 2


# ─── RECAP Sheet Generator ──────────────────────────────────────────────────

def _set_recap_col_widths(ws):
    """Set column widths for the RECAP sheet."""
    widths = {
        'A': 27,    # BRAND
        'B': 20,    # SIZE RANGE
        'C': 34,    # CATEGORY
        'D': 46,    # REF #
        'E': 13,    # OH
        'F': 13,    # WIP
        'G': 13,    # TOTAL ATS
    }
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width


def write_recap_sheet(ws, recap_sections: list, title: str = ""):
    """
    Write the RECAP sheet with brand summaries, size totals, and grand total.

    recap_sections: output of get_recap_data()
    """
    _set_recap_col_widths(ws)

    # Row 1: Title (merged A1:G1, yellow fill, bold)
    ws.merge_cells('A1:G1')
    safe_title = _safe_cell_text(title.upper() if title else "ATS RECAP")
    cell = ws.cell(row=1, column=1, value=safe_title)
    cell.fill = YELLOW_FILL
    cell.font = BOLD_FONT
    cell.alignment = CENTER_ALIGN

    # Row 2: Column headers (yellow fill, bold)
    headers = ["BRAND", "SIZE RANGE", "CATEGORY", "REF #", "OH", "WIP", "TOTAL ATS"]
    for col, label in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col, value=label)
        cell.fill = YELLOW_FILL
        cell.font = BOLD_FONT
        cell.alignment = CENTER_ALIGN
        cell.border = THIN_BORDER

    current_row = 3

    # Track all data rows for size-range totals
    all_data_rows = []
    brand_total_rows = []  # Track brand total row numbers for grand total formula

    for section in recap_sections:
        brand_label = section["brand_label"]
        rows = section["rows"]

        if not rows:
            continue

        section_start_row = current_row

        # Group rows by category for merging
        cat_groups = []
        current_cat = None
        for row_data in rows:
            if row_data["category"] != current_cat:
                cat_groups.append([row_data])
                current_cat = row_data["category"]
            else:
                cat_groups[-1].append(row_data)

        # Write data rows
        for cat_group in cat_groups:
            cat_start_row = current_row

            for row_data in cat_group:
                # Column B: SIZE RANGE
                ws.cell(row=current_row, column=2, value=row_data["size_range"]).font = NORMAL_FONT
                ws.cell(row=current_row, column=2).alignment = CENTER_ALIGN

                # Column D: REF #
                ws.cell(row=current_row, column=4, value=_safe_cell_text(row_data["ref_nums"])).font = NORMAL_FONT

                # Column E: OH
                cell = ws.cell(row=current_row, column=5, value=row_data["oh"])
                cell.font = NORMAL_FONT
                cell.alignment = CENTER_ALIGN
                cell.number_format = NUM_FMT

                # Column F: WIP
                cell = ws.cell(row=current_row, column=6, value=row_data["wip"])
                cell.font = NORMAL_FONT
                cell.alignment = CENTER_ALIGN
                cell.number_format = NUM_FMT

                # Column G: TOTAL ATS (formula)
                cell = ws.cell(row=current_row, column=7, value=f'=E{current_row}+F{current_row}')
                cell.font = NORMAL_FONT
                cell.alignment = CENTER_ALIGN
                cell.number_format = NUM_FMT

                # Borders
                for col in range(1, 8):
                    ws.cell(row=current_row, column=col).border = THIN_BORDER

                all_data_rows.append(current_row)
                current_row += 1

            # Merge category name (column C) if multiple rows
            if len(cat_group) > 1:
                ws.merge_cells(f'C{cat_start_row}:C{current_row - 1}')
            ws.cell(row=cat_start_row, column=3, value=_safe_cell_text(cat_group[0]["category"])).font = NORMAL_FONT
            ws.cell(row=cat_start_row, column=3).alignment = Alignment(
                horizontal='center', vertical='center', wrap_text=True
            )

        # Merge brand name (column A) for entire section
        if current_row > section_start_row:
            if current_row - section_start_row > 1:
                ws.merge_cells(f'A{section_start_row}:A{current_row - 1}')
            cell = ws.cell(row=section_start_row, column=1, value=_safe_cell_text(brand_label))
            cell.font = BOLD_FONT
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        # Brand total row (light blue fill)
        ws.merge_cells(f'A{current_row}:D{current_row}')
        cell = ws.cell(row=current_row, column=1, value=_safe_cell_text(f"{brand_label} TOTAL:"))
        cell.fill = LIGHT_BLUE_FILL
        cell.font = BOLD_FONT
        cell.alignment = CENTER_ALIGN

        # OH total (SUM formula)
        cell = ws.cell(row=current_row, column=5,
                       value=f'=SUM(E{section_start_row}:E{current_row - 1})')
        cell.fill = LIGHT_BLUE_FILL
        cell.font = BOLD_FONT
        cell.alignment = CENTER_ALIGN
        cell.number_format = NUM_FMT

        # WIP total
        cell = ws.cell(row=current_row, column=6,
                       value=f'=SUM(F{section_start_row}:F{current_row - 1})')
        cell.fill = LIGHT_BLUE_FILL
        cell.font = BOLD_FONT
        cell.alignment = CENTER_ALIGN
        cell.number_format = NUM_FMT

        # TOTAL ATS
        cell = ws.cell(row=current_row, column=7,
                       value=f'=E{current_row}+F{current_row}')
        cell.fill = LIGHT_BLUE_FILL
        cell.font = BOLD_FONT
        cell.alignment = CENTER_ALIGN
        cell.number_format = NUM_FMT

        # Apply light blue fill and borders to all columns
        for col in range(1, 8):
            ws.cell(row=current_row, column=col).fill = LIGHT_BLUE_FILL
            ws.cell(row=current_row, column=col).border = THIN_BORDER

        brand_total_rows.append(current_row)
        current_row += 1

    # Size range total rows (TODDLER TOTAL, 4-7 TOTAL, etc.)
    # Collect unique size ranges across all sections
    unique_srs = []
    for section in recap_sections:
        for row_data in section["rows"]:
            if row_data["size_range"] not in unique_srs:
                unique_srs.append(row_data["size_range"])

    # Sort: TODDLER first, then BOYS 4-7, then others
    sr_sort_order = ["TODDLER", "BOYS 4-7", "NEWBORN", "INFANT", "4-6X", "7-16", "8-20"]
    unique_srs.sort(key=lambda x: sr_sort_order.index(x) if x in sr_sort_order else 99)

    sr_total_rows = []
    for sr_name in unique_srs:
        # Find all data rows with this size range
        sr_data_rows = []
        for section in recap_sections:
            for i, row_data in enumerate(section["rows"]):
                if row_data["size_range"] == sr_name:
                    # Find the actual Excel row for this data row
                    # We need to track this better - use the row index offset
                    pass

        # Use SUMIF formula
        ws.merge_cells(f'A{current_row}:D{current_row}')
        cell = ws.cell(row=current_row, column=1, value=_safe_cell_text(f"{sr_name} TOTAL"))
        cell.fill = LIGHT_BLUE_FILL
        cell.font = BOLD_FONT
        cell.alignment = CENTER_ALIGN

        # SUMIF formulas based on column B (SIZE RANGE)
        first_data = min(all_data_rows) if all_data_rows else 3
        last_data = max(all_data_rows) if all_data_rows else 3

        # Sanitize sr_name for use inside formula (remove quotes and formula chars)
        safe_sr = re.sub(r'["\';=+\-@]', '', sr_name)

        cell = ws.cell(row=current_row, column=5,
                       value=f'=SUMIF(B{first_data}:B{last_data},"{safe_sr}",E{first_data}:E{last_data})')
        cell.fill = LIGHT_BLUE_FILL
        cell.font = BOLD_FONT
        cell.alignment = CENTER_ALIGN
        cell.number_format = NUM_FMT

        cell = ws.cell(row=current_row, column=6,
                       value=f'=SUMIF(B{first_data}:B{last_data},"{safe_sr}",F{first_data}:F{last_data})')
        cell.fill = LIGHT_BLUE_FILL
        cell.font = BOLD_FONT
        cell.alignment = CENTER_ALIGN
        cell.number_format = NUM_FMT

        cell = ws.cell(row=current_row, column=7,
                       value=f'=E{current_row}+F{current_row}')
        cell.fill = LIGHT_BLUE_FILL
        cell.font = BOLD_FONT
        cell.alignment = CENTER_ALIGN
        cell.number_format = NUM_FMT

        for col in range(1, 8):
            ws.cell(row=current_row, column=col).fill = LIGHT_BLUE_FILL
            ws.cell(row=current_row, column=col).border = THIN_BORDER

        sr_total_rows.append(current_row)
        current_row += 1

    # Grand Total row (yellow fill)
    ws.merge_cells(f'A{current_row}:D{current_row}')
    cell = ws.cell(row=current_row, column=1, value="GRAND TOTAL:")
    cell.fill = YELLOW_FILL
    cell.font = BOLD_FONT
    cell.alignment = CENTER_ALIGN

    # Grand total = sum of brand totals
    if brand_total_rows:
        oh_formula = "+".join(f"E{r}" for r in brand_total_rows)
        wip_formula = "+".join(f"F{r}" for r in brand_total_rows)
    else:
        oh_formula = "0"
        wip_formula = "0"

    cell = ws.cell(row=current_row, column=5, value=f'={oh_formula}')
    cell.fill = YELLOW_FILL
    cell.font = BOLD_FONT
    cell.alignment = CENTER_ALIGN
    cell.number_format = NUM_FMT

    cell = ws.cell(row=current_row, column=6, value=f'={wip_formula}')
    cell.fill = YELLOW_FILL
    cell.font = BOLD_FONT
    cell.alignment = CENTER_ALIGN
    cell.number_format = NUM_FMT

    cell = ws.cell(row=current_row, column=7, value=f'=E{current_row}+F{current_row}')
    cell.fill = YELLOW_FILL
    cell.font = BOLD_FONT
    cell.alignment = CENTER_ALIGN
    cell.number_format = NUM_FMT

    for col in range(1, 8):
        ws.cell(row=current_row, column=col).fill = YELLOW_FILL
        ws.cell(row=current_row, column=col).border = THIN_BORDER


# ─── Main Generator ─────────────────────────────────────────────────────────

def generate_ats_report(
    categories_by_sheet: Dict[str, dict],
    title: str = "",
    report_date: date = None,
) -> bytes:
    """
    Generate the complete ATS report Excel file.

    categories_by_sheet: {
        "NIKE BOYS 2-7 LONG BOTTOMS": {
            "brand": "NIKE",
            "general_category": "LONG BOTTOMS",
            "categories": [category_dict, ...],
        },
        ...
    }

    Returns the Excel file as bytes.
    """
    wb = Workbook()

    # Create RECAP sheet first
    ws_recap = wb.active
    ws_recap.title = "RECAP SHEET"

    # Build recap data
    recap_sections = get_recap_data(categories_by_sheet)

    # Write RECAP sheet
    write_recap_sheet(ws_recap, recap_sections, title=title)

    # Create detail sheets
    for sheet_name, sheet_info in categories_by_sheet.items():
        ws = wb.create_sheet(title=sheet_name[:31])  # Excel 31-char limit
        write_detail_sheet(ws, sheet_info["categories"], report_date=report_date)

    # Output
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()
