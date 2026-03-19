"""
ATS (Available to Ship) report parser.

Parses raw ATS Excel files that have been pasted from the internal system.
Detects ref# blocks, extracts style/color/size data, images, and totals.

Structure of raw ATS data per ref# block:
  1. Grey header row: STYLE | COLOR | SIZE SCALE | ON HAND | WIP | AVAILABILITY | MSRP
  2. Data row pairs (label row + ratio row per style/color/pack)
  3. Grey TOTAL row
  4. Empty rows / product image area

Style number format: e.g. "76F610-C5E-P3"
  - First 6 chars = base style ("76F610")
  - Last 4 of those 6 = ref# ("F610")
  - First char = size range code (7 = toddler, 8 = 4-7, etc.)
"""

import io
import logging
import re
import xml.etree.ElementTree as _ET
import zipfile
from collections import defaultdict
from typing import Dict, List, Optional, Tuple

import openpyxl
from PIL import Image

logger = logging.getLogger(__name__)

# Maximum image size to process (10MB per image)
MAX_IMAGE_SIZE = 10 * 1024 * 1024


# ─── Size range mapping ──────────────────────────────────────────────────────

SIZE_PREFIX_MAP = {
    "0": "NEWBORN",
    "1": "INFANT",
    "2": "TODDLER",
    "3": "4-6X",
    "4": "7-16",
    "5": "NEWBORN",
    "6": "INFANT",
    "7": "TODDLER",
    "8": "4-7",
    "9": "8-20",
}

# For the recap sheet display
SIZE_RANGE_DISPLAY = {
    "0": "NEWBORN",
    "1": "INFANT",
    "2": "TODDLER",
    "3": "4-6X",
    "4": "7-16",
    "5": "NEWBORN",
    "6": "INFANT",
    "7": "TODDLER",
    "8": "BOYS 4-7",
    "9": "8-20",
}

BRAND_KEYWORDS = {
    "JORDAN": "JORDAN",
    "NIKE": "NIKE",
    "HURLEY": "HURLEY",
    "CONVERSE": "CONVERSE",
    "LEVIS": "LEVIS",
    "LEVI": "LEVIS",
    "UMBRO": "UMBRO",
    "UNDER ARMOUR": "UNDER ARMOUR",
    "CHAMPION": "CHAMPION",
    "REEBOK": "REEBOK",
    "3BRAND": "3BRAND",
}


def detect_brand(text: str) -> str:
    """Detect brand from text (sheet name, etc.)."""
    t = text.upper()
    for kw, brand in BRAND_KEYWORDS.items():
        if kw in t:
            return brand
    return ""


def ref_from_style(style_raw: str) -> Tuple[str, str]:
    """
    Extract (base_style, ref_num) from a style string.
    "76F610-C5E-P3" → ("76F610", "F610")
    "86F651-C5E"    → ("86F651", "F651")
    """
    style_raw = style_raw.strip()
    base = style_raw.split("-")[0]
    if len(base) >= 6:
        ref = base[-4:]
    else:
        ref = base[-4:] if len(base) >= 4 else base
    return base, ref


def size_range_from_style(style_num: str) -> str:
    """Get size range code from first digit of style number."""
    first_char = ""
    for ch in style_num:
        if ch.isdigit():
            first_char = ch
            break
    return SIZE_PREFIX_MAP.get(first_char, "UNKNOWN")


def size_range_display(style_num: str) -> str:
    """Get display name for size range from style number."""
    first_char = ""
    for ch in style_num:
        if ch.isdigit():
            first_char = ch
            break
    return SIZE_RANGE_DISPLAY.get(first_char, "UNKNOWN")


def _safe_num(v, default=0) -> int:
    """Convert cell value to int, returning default for None/errors."""
    if v is None:
        return default
    s = str(v).strip()
    if not s or s.startswith("#"):
        return default
    # Skip date-like values
    if re.match(r"\d{1,2}/\d{1,2}/\d{2,4}", s):
        return default
    try:
        return int(float(s))
    except (ValueError, TypeError):
        return default


def _safe_float(v, default=0.0) -> float:
    """Convert cell value to float."""
    if v is None:
        return default
    s = str(v).strip()
    if not s or s.startswith("#"):
        return default
    try:
        return float(s)
    except (ValueError, TypeError):
        return default


def _safe_str(v) -> str:
    """Convert cell value to string."""
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.startswith("#") else s


def _safe_zip_path(target: str, prefix: str = 'xl/') -> str:
    """
    Safely resolve a ZIP-internal path, preventing path traversal.
    Returns cleaned path or empty string if suspicious.
    """
    # Remove all path traversal attempts
    cleaned = target
    while '../' in cleaned:
        cleaned = cleaned.replace('../', '')
    while '..' in cleaned:
        cleaned = cleaned.replace('..', '')
    cleaned = cleaned.lstrip('/')

    # Ensure it starts with expected prefix or add it
    if not cleaned.startswith(prefix):
        cleaned = prefix + cleaned

    # Final check: no traversal patterns remain
    if '..' in cleaned or cleaned.startswith('/'):
        logger.warning(f"Blocked suspicious ZIP path: {target}")
        return ''

    return cleaned


# ─── Image extraction ────────────────────────────────────────────────────────

def _extract_images(file_bytes: bytes) -> Dict[str, Dict[int, bytes]]:
    """
    Extract images from xlsx file.
    Returns {sheet_name: {row_1indexed: img_bytes}}.
    """
    result: Dict[str, Dict[int, bytes]] = {}
    _NS_REL = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
    _NS_XDR = {
        'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
        'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
        'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    }
    _NS_WB = {
        'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    }
    _R_EMBED = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'

    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
            names = set(zf.namelist())

            wb_root = _ET.fromstring(zf.read('xl/workbook.xml'))
            sheet_rid: Dict[str, str] = {}
            for sh in wb_root.findall('.//x:sheet', _NS_WB):
                rid = sh.get(_R_EMBED, '')
                name = sh.get('name', '')
                if rid and name:
                    sheet_rid[rid] = name

            wb_rels_path = 'xl/_rels/workbook.xml.rels'
            if wb_rels_path not in names:
                return result
            rels_root = _ET.fromstring(zf.read(wb_rels_path))
            rid_to_file: Dict[str, str] = {}
            for rel in rels_root.findall('r:Relationship', _NS_REL):
                rid_to_file[rel.get('Id', '')] = rel.get('Target', '')

            for rid, sheet_name in sheet_rid.items():
                ws_file = rid_to_file.get(rid, '')
                if not ws_file:
                    continue

                ws_base = ws_file.replace('worksheets/', '')
                ws_rels = f'xl/worksheets/_rels/{ws_base}.rels'
                if ws_rels not in names:
                    continue

                ws_rels_root = _ET.fromstring(zf.read(ws_rels))
                drawing_file = None
                for rel in ws_rels_root.findall('r:Relationship', _NS_REL):
                    if 'drawing' in rel.get('Type', '').lower():
                        target = rel.get('Target', '')
                        drawing_file = _safe_zip_path(target, 'xl/')
                        break

                if not drawing_file or drawing_file not in names:
                    continue

                draw_rels_path = drawing_file.replace('/drawings/', '/drawings/_rels/') + '.rels'
                img_rid_to_path: Dict[str, str] = {}
                if draw_rels_path in names:
                    dr_root = _ET.fromstring(zf.read(draw_rels_path))
                    for rel in dr_root.findall('r:Relationship', _NS_REL):
                        if 'image' in rel.get('Type', '').lower():
                            img_rid_val = rel.get('Id', '')
                            img_target = rel.get('Target', '')
                            img_path = _safe_zip_path(img_target, 'xl/')
                            if img_path:  # Only add if path is safe
                                img_rid_to_path[img_rid_val] = img_path

                if not img_rid_to_path:
                    continue

                draw_root = _ET.fromstring(zf.read(drawing_file))
                sheet_imgs: Dict[int, List[Tuple[bytes, int, int]]] = defaultdict(list)

                for anchor_tag in ('xdr:oneCellAnchor', 'xdr:twoCellAnchor'):
                    for anchor in draw_root.findall(anchor_tag, _NS_XDR):
                        from_el = anchor.find('xdr:from', _NS_XDR)
                        if from_el is None:
                            continue
                        row_el = from_el.find('xdr:row', _NS_XDR)
                        col_el = from_el.find('xdr:col', _NS_XDR)
                        if row_el is None:
                            continue
                        row_1 = int(row_el.text or '0') + 1
                        col_0 = int(col_el.text or '0') if col_el is not None else 0

                        blip = anchor.find('.//a:blip', _NS_XDR)
                        if blip is None:
                            continue
                        img_rid_val = blip.get(
                            '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed', '')
                        if not img_rid_val:
                            continue
                        img_path = img_rid_to_path.get(img_rid_val, '')
                        if not img_path or img_path not in names:
                            continue

                        # Check image size before reading
                        img_info = zf.getinfo(img_path)
                        if img_info.file_size > MAX_IMAGE_SIZE:
                            continue

                        img_data = zf.read(img_path)
                        sheet_imgs[row_1].append((img_data, col_0, row_1))

                if sheet_imgs:
                    # Classify images: product images are large (>100px), swatches are small
                    product_imgs: Dict[int, bytes] = {}
                    for row_num, img_list in sheet_imgs.items():
                        for img_data, col, _ in img_list:
                            if _is_product_image(img_data):
                                product_imgs[row_num] = img_data
                                break
                    result[sheet_name] = product_imgs

    except Exception:
        pass

    return result


def _is_product_image(img_bytes: bytes) -> bool:
    """
    Determine if an image is a product image (large) vs a color swatch (small).
    Product images are typically 400x400+ pixels.
    Color swatches are typically 20x20 to 80x80 pixels.
    """
    try:
        img = Image.open(io.BytesIO(img_bytes))
        w, h = img.size
        # Product images are large; swatches are small
        return w > 100 and h > 100
    except Exception:
        return False


# ─── Sheet parsing ───────────────────────────────────────────────────────────

def _is_header_row(row_cells: dict) -> bool:
    """Check if a row is a grey header row (STYLE, COLOR, SIZE SCALE, etc.)."""
    values_upper = {c: str(v).upper().strip() for c, v in row_cells.items() if v is not None}
    has_style = any(v in ("STYLE", "STYLE #", "STYLE NUMBER", "STYLE NO") for v in values_upper.values())
    has_data_col = any(v in ("ON HAND", "OH", "OH ATS", "SIZE SCALE", "COLOR", "WIP", "MSRP")
                       for v in values_upper.values())
    return has_style and has_data_col


def _is_total_row(row_cells: dict) -> bool:
    """Check if a row is a TOTAL row."""
    for v in row_cells.values():
        if v is not None and "TOTAL" in str(v).upper().strip():
            return True
    return False


def _find_col_mapping(row_cells: dict) -> dict:
    """Map column indices to field names from a header row."""
    mapping = {}
    for col, val in row_cells.items():
        if val is None:
            continue
        v = str(val).upper().strip()
        if v in ("STYLE", "STYLE #", "STYLE NUMBER", "STYLE NO"):
            mapping["style"] = col
        elif v in ("COLOR", "COLOR NAME"):
            mapping["color"] = col
        elif v in ("SIZE SCALE", "SIZE"):
            mapping["size_scale_start"] = col
        elif v in ("ON HAND", "OH", "OH ATS"):
            mapping["oh"] = col
        elif v in ("WIP", "TOTAL WIP"):
            mapping["wip"] = col
        elif v in ("AVAILABILITY",):
            mapping["availability"] = col
        elif v in ("MSRP", "RRP", "RETAIL"):
            mapping["msrp"] = col
    return mapping


def parse_ats_file(file_bytes: bytes) -> dict:
    """
    Parse a raw ATS Excel file.

    Returns:
    {
        "sheets": [
            {
                "name": str,
                "brand": str,
                "blocks": [
                    {
                        "ref_num": str,
                        "base_style": str,
                        "rows": [  # raw row data, in pairs (labels + ratios)
                            {
                                "style_num": str,
                                "color": str,
                                "cells": {col_idx: value, ...},  # all cell values
                                "oh": int,
                                "wip": int,
                                "availability": str,
                                "msrp": float,
                                "size_range": str,
                                "size_range_display": str,
                                "is_label_row": bool,  # True for size labels, False for ratios
                            },
                            ...
                        ],
                        "total_oh": int,
                        "total_wip": int,
                        "product_image": bytes or None,
                    },
                    ...
                ],
                "all_ref_nums": [str, ...],  # unique ref nums found
            },
            ...
        ]
    }
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    all_images = _extract_images(file_bytes)

    result = {"sheets": []}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        brand = detect_brand(sheet_name)
        images_by_row = all_images.get(sheet_name, {})

        sheet_data = {
            "name": sheet_name,
            "brand": brand,
            "blocks": [],
            "all_ref_nums": [],
        }

        # Read all rows into memory
        all_rows = []
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column, values_only=False):
            row_dict = {}
            for cell in row:
                if cell.value is not None:
                    row_dict[cell.column] = cell.value
            all_rows.append(row_dict)

        # Scan for ref# blocks
        col_map = None
        current_block_rows = []
        current_block_start = None
        i = 0

        while i < len(all_rows):
            row_cells = all_rows[i]
            row_num = i + 1  # 1-indexed

            # Check if this is a header row
            if _is_header_row(row_cells):
                col_map = _find_col_mapping(row_cells)
                current_block_rows = []
                current_block_start = row_num
                i += 1
                continue

            # Check if this is a TOTAL row
            if _is_total_row(row_cells) and current_block_rows:
                total_oh = _safe_num(row_cells.get(col_map.get("oh", 12) if col_map else 12))
                total_wip = _safe_num(row_cells.get(col_map.get("wip", 13) if col_map else 13))

                # Group rows by ref#
                blocks_by_ref = defaultdict(list)
                for row_data in current_block_rows:
                    blocks_by_ref[row_data["ref_num"]].append(row_data)

                for ref_num, ref_rows in blocks_by_ref.items():
                    # Calculate OH/WIP for this specific ref# from data rows
                    block_oh = sum(r["oh"] for r in ref_rows if r.get("is_label_row", True))
                    block_wip = sum(r["wip"] for r in ref_rows if r.get("is_label_row", True))

                    # Find product image for this block (scan backward from header)
                    product_img = None
                    if current_block_start:
                        for look_row in range(current_block_start, max(0, current_block_start - 15), -1):
                            if look_row in images_by_row:
                                product_img = images_by_row[look_row]
                                break

                    base_style = ref_rows[0]["base_style"] if ref_rows else ""

                    sheet_data["blocks"].append({
                        "ref_num": ref_num,
                        "base_style": base_style,
                        "rows": ref_rows,
                        "total_oh": block_oh if len(blocks_by_ref) > 1 else total_oh,
                        "total_wip": block_wip if len(blocks_by_ref) > 1 else total_wip,
                        "product_image": product_img,
                    })

                current_block_rows = []
                current_block_start = None
                i += 1
                continue

            # Regular data row - extract style data
            if col_map and row_cells:
                style_col = col_map.get("style", 3)
                style_raw = _safe_str(row_cells.get(style_col))

                if style_raw and any(c.isdigit() for c in style_raw):
                    # Skip header-like rows
                    if style_raw.upper() not in ("STYLE", "STYLE #", "STYLE NUMBER"):
                        base_style, ref_num = ref_from_style(style_raw)
                        color = _safe_str(row_cells.get(col_map.get("color", 4)))
                        oh_col = col_map.get("oh", 12)
                        wip_col = col_map.get("wip", 13)
                        avail_col = col_map.get("availability", 14)
                        msrp_col = col_map.get("msrp", 15)

                        oh_val = _safe_num(row_cells.get(oh_col))
                        wip_val = _safe_num(row_cells.get(wip_col))

                        # Determine if this is a label row or ratio row
                        # Label rows have size labels like "2T", "3T", "4T" or "4", "5", "6", "7"
                        # AND typically have OH values
                        # Ratio rows have just numbers in size columns and no OH
                        is_label_row = oh_val > 0 or wip_val > 0

                        # Also check: if size columns contain text labels like "2T", it's a label row
                        for col_idx in range(5, 12):  # columns E through K
                            cell_val = row_cells.get(col_idx)
                            if cell_val is not None and isinstance(cell_val, str) and 'T' in str(cell_val).upper():
                                is_label_row = True
                                break

                        size_range = size_range_from_style(style_raw)
                        size_display = size_range_display(style_raw)

                        row_data = {
                            "style_num": style_raw,
                            "base_style": base_style,
                            "ref_num": ref_num,
                            "color": color,
                            "cells": dict(row_cells),
                            "oh": oh_val,
                            "wip": wip_val,
                            "availability": _safe_str(row_cells.get(avail_col)),
                            "msrp": _safe_float(row_cells.get(msrp_col)),
                            "size_range": size_range,
                            "size_range_display": size_display,
                            "is_label_row": is_label_row,
                            "row_num": row_num,
                        }
                        current_block_rows.append(row_data)

            i += 1

        # Collect unique ref nums
        seen_refs = []
        for block in sheet_data["blocks"]:
            if block["ref_num"] not in seen_refs:
                seen_refs.append(block["ref_num"])
        sheet_data["all_ref_nums"] = seen_refs

        result["sheets"].append(sheet_data)

    return result


def filter_blocks(blocks: list, min_units: int = 120, max_units: int = None) -> list:
    """
    Filter blocks: remove style/color packs with OH + WIP < min_units.
    Optionally also remove packs with OH + WIP > max_units.
    Re-totals after filtering.

    Returns filtered blocks (blocks with no remaining styles are removed entirely).
    """
    filtered_blocks = []

    for block in blocks:
        filtered_rows = []
        i = 0
        rows = block["rows"]

        while i < len(rows):
            row = rows[i]

            if row.get("is_label_row", True):
                # This is a label row - check the OH + WIP
                total_units = row["oh"] + row["wip"]

                # Find the corresponding ratio row (next row with same style+color)
                ratio_row = None
                if i + 1 < len(rows):
                    next_row = rows[i + 1]
                    if (next_row.get("style_num") == row.get("style_num") and
                            next_row.get("color") == row.get("color") and
                            not next_row.get("is_label_row", True)):
                        ratio_row = next_row

                keep = total_units >= min_units
                if max_units is not None and total_units > max_units:
                    keep = False

                if keep:
                    filtered_rows.append(row)
                    if ratio_row:
                        filtered_rows.append(ratio_row)
                        i += 2
                    else:
                        i += 1
                else:
                    # Skip this pack (and its ratio row if present)
                    if ratio_row:
                        i += 2
                    else:
                        i += 1
            else:
                # Ratio row without a preceding label row - skip
                i += 1

        if filtered_rows:
            # Re-calculate totals
            new_oh = sum(r["oh"] for r in filtered_rows if r.get("is_label_row", True))
            new_wip = sum(r["wip"] for r in filtered_rows if r.get("is_label_row", True))

            filtered_blocks.append({
                **block,
                "rows": filtered_rows,
                "total_oh": new_oh,
                "total_wip": new_wip,
            })

    return filtered_blocks


def group_blocks_by_category(blocks: list, category_map: dict) -> list:
    """
    Group blocks into categories based on a ref# → category mapping.

    category_map: {"BURPEE JOGGER": ["F610", "F651"], "THERMA PANT": ["J785", "N271"], ...}

    Returns list of category dicts:
    [
        {
            "name": "BURPEE JOGGER",
            "blocks": [block1, block2, ...],
            "size_ranges": {
                "TODDLER": {"oh": 19236, "wip": 0, "ref_nums": ["F610"]},
                "BOYS 4-7": {"oh": 37152, "wip": 0, "ref_nums": ["F610", "F651"]},
            }
        },
        ...
    ]
    """
    # Build reverse map: ref_num → category
    ref_to_cat = {}
    for cat_name, ref_nums in category_map.items():
        for ref in ref_nums:
            ref_to_cat[ref.strip().upper()] = cat_name

    # Group blocks by category
    cats = defaultdict(list)
    for block in blocks:
        ref_upper = block["ref_num"].strip().upper()
        cat_name = ref_to_cat.get(ref_upper, "UNCATEGORIZED")
        cats[cat_name].append(block)

    # Build category summaries
    result = []
    # Preserve category order from category_map
    ordered_cats = list(category_map.keys())
    # Add any uncategorized at the end
    for cat_name in cats:
        if cat_name not in ordered_cats:
            ordered_cats.append(cat_name)

    for cat_name in ordered_cats:
        if cat_name not in cats:
            continue
        cat_blocks = cats[cat_name]

        # Calculate per-size-range summaries
        size_ranges = defaultdict(lambda: {"oh": 0, "wip": 0, "ref_nums": set()})

        for block in cat_blocks:
            for row in block["rows"]:
                if row.get("is_label_row", True):
                    sr = row["size_range_display"]
                    size_ranges[sr]["oh"] += row["oh"]
                    size_ranges[sr]["wip"] += row["wip"]
                    size_ranges[sr]["ref_nums"].add(row["ref_num"])

        # Convert sets to sorted lists
        for sr_data in size_ranges.values():
            sr_data["ref_nums"] = sorted(sr_data["ref_nums"])

        result.append({
            "name": cat_name,
            "blocks": cat_blocks,
            "size_ranges": dict(size_ranges),
        })

    return result


def get_recap_data(categories_by_sheet: dict) -> list:
    """
    Build recap data for the RECAP tab.

    categories_by_sheet: {
        "NIKE BOYS 2-7 LONG BOTTOMS": {
            "brand": "NIKE",
            "general_category": "LONG BOTTOMS",
            "categories": [category_dict, ...]
        },
        ...
    }

    Returns list of recap sections:
    [
        {
            "brand_label": "NIKE LONG BOTTOMS",
            "brand": "NIKE",
            "rows": [
                {"size_range": "TODDLER", "category": "BURPEE JOGGER",
                 "ref_nums": "F610", "oh": 19236, "wip": 0, "total": 19236},
                {"size_range": "BOYS 4-7", "category": "BURPEE JOGGER",
                 "ref_nums": "F610, F651", "oh": 37152, "wip": 0, "total": 37152},
                ...
            ],
            "total_oh": int,
            "total_wip": int,
            "total_ats": int,
        },
        ...
    ]
    """
    recap_sections = []

    for sheet_name, sheet_info in categories_by_sheet.items():
        brand = sheet_info["brand"]
        gen_cat = sheet_info.get("general_category", "")

        # Build brand label like "NIKE LONG BOTTOMS" or "JORDAN TEES"
        brand_label = f"{brand} {gen_cat}".strip() if gen_cat else brand

        section_rows = []
        total_oh = 0
        total_wip = 0

        for cat in sheet_info["categories"]:
            cat_name = cat["name"]
            size_ranges = cat["size_ranges"]

            # Determine order: TODDLER first, then 4-7, then others
            sr_order = []
            for sr_name in ["TODDLER", "BOYS 4-7", "NEWBORN", "INFANT", "4-6X", "7-16", "8-20"]:
                if sr_name in size_ranges:
                    sr_order.append(sr_name)
            # Add any remaining
            for sr_name in size_ranges:
                if sr_name not in sr_order:
                    sr_order.append(sr_name)

            for sr_name in sr_order:
                sr_data = size_ranges[sr_name]
                if sr_data["oh"] == 0 and sr_data["wip"] == 0:
                    continue  # Omit empty size ranges

                oh = sr_data["oh"]
                wip = sr_data["wip"]
                ref_str = ", ".join(sr_data["ref_nums"])

                section_rows.append({
                    "size_range": sr_name,
                    "category": cat_name,
                    "ref_nums": ref_str,
                    "oh": oh,
                    "wip": wip,
                    "total": oh + wip,
                })

                total_oh += oh
                total_wip += wip

        recap_sections.append({
            "brand_label": brand_label,
            "brand": brand,
            "rows": section_rows,
            "total_oh": total_oh,
            "total_wip": total_wip,
            "total_ats": total_oh + total_wip,
        })

    return recap_sections
