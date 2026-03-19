"""
ATS (Available to Ship) report parser.

Parses raw ATS Excel files pasted from the internal system.

Category detection: Scans for TODDLER/4-7 pattern in column K.
  - Row with K="TODDLER" has summary OH/WIP/TOTAL in L/M/N
  - Row below has category name in A, K="4-7", and summary OH/WIP/TOTAL
  - All style blocks between two category headers belong to the same category

REF# extraction from style codes:
  - Strip first 2 characters (size prefix): 76F610-C5E-P3 → F610-C5E-P3
  - Take everything before the first dash: F610-C5E-P3 → F610

Size range from first digit of style code:
  - 7x = Toddler (76=Nike, 75=Jordan)
  - 8x = Boys 4-7 (86=Nike, 85=Jordan)
"""

import io
import logging
import re
import xml.etree.ElementTree as _ET
import zipfile
from collections import defaultdict, OrderedDict
from typing import Dict, List, Tuple

import openpyxl
try:
    from PIL import Image
except ImportError:
    Image = None

logger = logging.getLogger(__name__)
MAX_IMAGE_SIZE = 10 * 1024 * 1024

BRAND_KEYWORDS = {
    "JORDAN": "JORDAN", "NIKE": "NIKE", "HURLEY": "HURLEY",
    "CONVERSE": "CONVERSE", "LEVIS": "LEVIS",
    "UMBRO": "UMBRO", "UNDER ARMOUR": "UNDER ARMOUR",
    "CHAMPION": "CHAMPION", "REEBOK": "REEBOK",
}


def detect_brand(text: str) -> str:
    t = text.upper()
    for kw, brand in BRAND_KEYWORDS.items():
        if kw in t:
            return brand
    return ""


def ref_from_style(style_raw: str) -> str:
    """
    Extract REF# from style code.
    1. Strip first 2 chars (size prefix): 76F610-C5E-P3 → F610-C5E-P3
    2. Take everything before first dash: F610-C5E-P3 → F610
    """
    style_raw = style_raw.strip()
    if len(style_raw) < 3:
        return style_raw
    stripped = style_raw[2:]  # Remove first 2 chars
    ref = stripped.split("-")[0]  # Take before first dash
    return ref


def size_range_from_style(style_num: str) -> str:
    """
    First digit: 7=Toddler, 8=Boys 4-7.
    Works for both Nike (76/86) and Jordan (75/85).
    """
    if not style_num:
        return "UNKNOWN"
    first = style_num[0]
    if first == '7':
        return "TODDLER"
    elif first == '8':
        return "BOYS 4-7"
    else:
        return "UNKNOWN"


def _safe_num(v, default=0) -> int:
    if v is None:
        return default
    s = str(v).strip()
    if not s or s.startswith("#"):
        return default
    if re.match(r"\d{1,2}/\d{1,2}/\d{2,4}", s):
        return default
    try:
        return int(float(s))
    except (ValueError, TypeError):
        return default


def _safe_float(v, default=0.0) -> float:
    if v is None:
        return default
    try:
        return float(str(v).strip())
    except (ValueError, TypeError):
        return default


def _safe_str(v) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.startswith("#") else s


# ─── Image extraction (unchanged) ────────────────────────────────────────────

def _is_yellow(cell) -> bool:
    """Check if a cell has yellow fill (category header)."""
    try:
        if cell.fill and cell.fill.start_color:
            rgb = str(cell.fill.start_color.rgb or '').upper()
            return 'FFFF00' in rgb
    except Exception:
        pass
    return False


def _safe_zip_path(target: str, prefix: str = 'xl/') -> str:
    cleaned = target
    while '../' in cleaned:
        cleaned = cleaned.replace('../', '')
    cleaned = cleaned.lstrip('/')
    if not cleaned.startswith(prefix):
        cleaned = prefix + cleaned
    if '..' in cleaned:
        return ''
    return cleaned


def _extract_images(file_bytes: bytes) -> Dict[str, Dict[int, bytes]]:
    result = {}
    _NS_REL = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
    _NS_XDR = {
        'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
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
            sheet_rid = {}
            for sh in wb_root.findall('.//x:sheet', _NS_WB):
                rid, name = sh.get(_R_EMBED, ''), sh.get('name', '')
                if rid and name:
                    sheet_rid[rid] = name

            if 'xl/_rels/workbook.xml.rels' not in names:
                return result
            rels_root = _ET.fromstring(zf.read('xl/_rels/workbook.xml.rels'))
            rid_to_file = {r.get('Id', ''): r.get('Target', '') for r in rels_root.findall('r:Relationship', _NS_REL)}

            for rid, sheet_name in sheet_rid.items():
                ws_file = rid_to_file.get(rid, '')
                if not ws_file:
                    continue
                ws_rels = f'xl/worksheets/_rels/{ws_file.replace("worksheets/", "")}.rels'
                if ws_rels not in names:
                    continue
                ws_rels_root = _ET.fromstring(zf.read(ws_rels))
                drawing_file = None
                for rel in ws_rels_root.findall('r:Relationship', _NS_REL):
                    if 'drawing' in rel.get('Type', '').lower():
                        drawing_file = _safe_zip_path(rel.get('Target', ''), 'xl/')
                        break
                if not drawing_file or drawing_file not in names:
                    continue
                draw_rels = drawing_file.replace('/drawings/', '/drawings/_rels/') + '.rels'
                img_rid_map = {}
                if draw_rels in names:
                    for rel in _ET.fromstring(zf.read(draw_rels)).findall('r:Relationship', _NS_REL):
                        if 'image' in rel.get('Type', '').lower():
                            p = _safe_zip_path(rel.get('Target', ''), 'xl/')
                            if p:
                                img_rid_map[rel.get('Id', '')] = p
                if not img_rid_map:
                    continue
                draw_root = _ET.fromstring(zf.read(drawing_file))
                product_imgs = {}
                for tag in ('xdr:oneCellAnchor', 'xdr:twoCellAnchor'):
                    for anchor in draw_root.findall(tag, _NS_XDR):
                        fr = anchor.find('xdr:from', _NS_XDR)
                        if fr is None:
                            continue
                        re_ = fr.find('xdr:row', _NS_XDR)
                        if re_ is None:
                            continue
                        row_1 = int(re_.text or '0') + 1
                        blip = anchor.find('.//a:blip', _NS_XDR)
                        if blip is None:
                            continue
                        irid = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed', '')
                        ipath = img_rid_map.get(irid, '')
                        if not ipath or ipath not in names:
                            continue
                        info = zf.getinfo(ipath)
                        if info.file_size > MAX_IMAGE_SIZE:
                            continue
                        data = zf.read(ipath)
                        try:
                            if Image is None:
                                product_imgs[row_1] = data
                                continue
                            img = Image.open(io.BytesIO(data))
                            if img.size[0] > 100 and img.size[1] > 100:
                                product_imgs[row_1] = data
                        except Exception:
                            pass
                if product_imgs:
                    result[sheet_name] = product_imgs
    except Exception:
        pass
    return result


# ─── Main Parser ─────────────────────────────────────────────────────────────

def parse_ats_file(file_bytes: bytes) -> dict:
    """
    Parse raw ATS Excel file.

    Returns:
    {
        "sheets": [
            {
                "name": str,
                "brand": str,
                "categories": list of {
                    "name": str,
                    "toddler_oh": int, "toddler_wip": int, "toddler_total": int,
                    "boys47_oh": int, "boys47_wip": int, "boys47_total": int,
                    "toddler_refs": [str, ...],  # unique REF#s from toddler styles
                    "boys47_refs": [str, ...],    # unique REF#s from boys 4-7 styles
                    "blocks": [block_dict, ...],
                },
                "all_ref_nums": [str, ...],
            },
            ...
        ]
    }
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    # Second load for formatting (yellow cell detection)
    wb_fmt = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
    all_images = _extract_images(file_bytes)

    result = {"sheets": []}

    for sheet_name in wb.sheetnames:
        # Skip recap/summary sheets
        sn_upper = sheet_name.upper().strip()
        if 'RECAP' in sn_upper and 'SHEET' in sn_upper:
            continue
        if sn_upper in ('RECAP', 'SUMMARY'):
            continue

        ws = wb[sheet_name]
        ws_fmt = wb_fmt[sheet_name]
        brand = detect_brand(sheet_name)
        images_by_row = all_images.get(sheet_name, {})

        # ── Step 1: Detect categories from yellow cells in column A ──
        category_rows = []  # [(row_num, category_name)]
        for row_num in range(1, ws.max_row + 1):
            cell_val = ws.cell(row=row_num, column=1).value
            cell_fmt = ws_fmt.cell(row=row_num, column=1)
            if cell_val and _is_yellow(cell_fmt):
                cat_name = str(cell_val).strip()
                if cat_name.upper() not in ("ATS RECAP", "ATS", "RECAP", "OH", "WIP", "TOTAL"):
                    category_rows.append((row_num, cat_name))

        if not category_rows:
            result["sheets"].append({
                "name": sheet_name, "brand": brand,
                "categories": [], "all_ref_nums": [],
            })
            continue

        logger.info(f"Sheet '{sheet_name}': {len(category_rows)} categories: "
                    f"{[c[1] for c in category_rows]}")

        # ── Step 2: Parse style blocks per category ──
        categories = []
        all_ref_nums = []

        for cat_idx, (cat_row, cat_name) in enumerate(category_rows):
            # Data range: from category row to next category (or end of sheet)
            data_start = cat_row + 1
            data_end = category_rows[cat_idx + 1][0] - 1 if cat_idx + 1 < len(category_rows) else ws.max_row

            blocks = []
            toddler_refs, boys47_refs = [], []

            row = data_start
            while row <= data_end:
                c_val = _safe_str(ws.cell(row=row, column=3).value)

                # Sub-header row (STYLE, COLOR, SIZE SCALE)
                if c_val.upper() in ("STYLE", "STYLE #", "STYLE NUMBER"):
                    block_header_row = row
                    row += 1
                    block_rows = []

                    while row <= data_end:
                        cv = _safe_str(ws.cell(row=row, column=3).value)

                        if cv.upper().startswith("TOTAL"):
                            total_oh = _safe_num(ws.cell(row=row, column=12).value)
                            total_wip = _safe_num(ws.cell(row=row, column=13).value)

                            block_ref = block_rows[0]["ref_num"] if block_rows else ""

                            # Product image
                            product_img = None
                            for lr in range(block_header_row, max(0, block_header_row - 15), -1):
                                if lr in images_by_row:
                                    product_img = images_by_row[lr]
                                    break

                            blocks.append({
                                "ref_num": block_ref,
                                "rows": block_rows,
                                "total_oh": total_oh,
                                "total_wip": total_wip,
                                "product_image": product_img,
                            })

                            # Collect refs by size range
                            for br in block_rows:
                                sr = br.get("size_range", "")
                                ref = br.get("ref_num", "")
                                if ref:
                                    if sr == "TODDLER" and ref not in toddler_refs:
                                        toddler_refs.append(ref)
                                    elif sr == "BOYS 4-7" and ref not in boys47_refs:
                                        boys47_refs.append(ref)

                            row += 1
                            break

                        # Style data row
                        if cv and any(ch.isdigit() for ch in cv):
                            color = _safe_str(ws.cell(row=row, column=4).value)
                            oh = _safe_num(ws.cell(row=row, column=12).value)
                            wip = _safe_num(ws.cell(row=row, column=13).value)
                            avail = _safe_str(ws.cell(row=row, column=14).value)
                            msrp = _safe_float(ws.cell(row=row, column=15).value)

                            is_label_row = oh > 0 or wip > 0
                            for ci in range(5, 12):
                                cval = ws.cell(row=row, column=ci).value
                                if cval and isinstance(cval, str) and 'T' in cval.upper():
                                    is_label_row = True
                                    break

                            ref = ref_from_style(cv)
                            sr = size_range_from_style(cv)

                            cells = {}
                            for ci in range(5, 12):
                                cval = ws.cell(row=row, column=ci).value
                                if cval is not None:
                                    cells[ci] = cval

                            block_rows.append({
                                "style_num": cv, "ref_num": ref, "color": color,
                                "cells": cells, "oh": oh, "wip": wip,
                                "availability": avail, "msrp": msrp,
                                "size_range": sr, "is_label_row": is_label_row,
                            })
                        row += 1
                else:
                    row += 1

            # Compute OH/WIP from style data
            tod_oh = sum(r["oh"] for b in blocks for r in b["rows"]
                        if r.get("is_label_row") and r.get("size_range") == "TODDLER")
            tod_wip = sum(r["wip"] for b in blocks for r in b["rows"]
                         if r.get("is_label_row") and r.get("size_range") == "TODDLER")
            b47_oh = sum(r["oh"] for b in blocks for r in b["rows"]
                        if r.get("is_label_row") and r.get("size_range") == "BOYS 4-7")
            b47_wip = sum(r["wip"] for b in blocks for r in b["rows"]
                         if r.get("is_label_row") and r.get("size_range") == "BOYS 4-7")

            for r in toddler_refs + boys47_refs:
                if r not in all_ref_nums:
                    all_ref_nums.append(r)

            categories.append({
                "name": cat_name,
                "toddler_oh": tod_oh, "toddler_wip": tod_wip,
                "toddler_total": tod_oh + tod_wip,
                "boys47_oh": b47_oh, "boys47_wip": b47_wip,
                "boys47_total": b47_oh + b47_wip,
                "toddler_refs": toddler_refs,
                "boys47_refs": boys47_refs,
                "blocks": blocks,
            })

        result["sheets"].append({
            "name": sheet_name, "brand": brand,
            "categories": categories, "all_ref_nums": all_ref_nums,
        })

    return result


# ─── Filter ──────────────────────────────────────────────────────────────────

def filter_blocks(blocks: list, min_units: int = 120, max_units: int = None) -> list:
    """Filter style packs by OH + WIP threshold. Returns filtered blocks."""
    filtered = []
    for block in blocks:
        frows = []
        i = 0
        rows = block["rows"]
        while i < len(rows):
            r = rows[i]
            if r.get("is_label_row", True):
                total = r["oh"] + r["wip"]
                ratio = None
                if i + 1 < len(rows):
                    nxt = rows[i + 1]
                    if nxt.get("style_num") == r.get("style_num") and not nxt.get("is_label_row", True):
                        ratio = nxt
                keep = total >= min_units
                if max_units and total > max_units:
                    keep = False
                if keep:
                    frows.append(r)
                    if ratio:
                        frows.append(ratio)
                        i += 2
                    else:
                        i += 1
                else:
                    if ratio:
                        i += 2
                    else:
                        i += 1
            else:
                i += 1
        if frows:
            new_oh = sum(x["oh"] for x in frows if x.get("is_label_row"))
            new_wip = sum(x["wip"] for x in frows if x.get("is_label_row"))
            filtered.append({**block, "rows": frows, "total_oh": new_oh, "total_wip": new_wip})
    return filtered


def filter_categories(categories: list, min_units: int = 120, max_units: int = None) -> list:
    """Filter all categories' blocks. Recompute ref#s after filtering."""
    result = []
    for cat in categories:
        fblocks = filter_blocks(cat["blocks"], min_units, max_units)
        if fblocks:
            # Recompute refs after filtering
            tod_refs, b47_refs = [], []
            for block in fblocks:
                for r in block["rows"]:
                    if r.get("is_label_row"):
                        ref = r.get("ref_num", "")
                        sr = r.get("size_range", "")
                        if sr == "TODDLER" and ref and ref not in tod_refs:
                            tod_refs.append(ref)
                        elif sr == "BOYS 4-7" and ref and ref not in b47_refs:
                            b47_refs.append(ref)

            # Recompute OH/WIP from filtered data
            tod_oh = sum(r["oh"] for b in fblocks for r in b["rows"]
                        if r.get("is_label_row") and r.get("size_range") == "TODDLER")
            tod_wip = sum(r["wip"] for b in fblocks for r in b["rows"]
                         if r.get("is_label_row") and r.get("size_range") == "TODDLER")
            b47_oh = sum(r["oh"] for b in fblocks for r in b["rows"]
                        if r.get("is_label_row") and r.get("size_range") == "BOYS 4-7")
            b47_wip = sum(r["wip"] for b in fblocks for r in b["rows"]
                         if r.get("is_label_row") and r.get("size_range") == "BOYS 4-7")

            result.append({
                **cat,
                "blocks": fblocks,
                "toddler_refs": tod_refs,
                "boys47_refs": b47_refs,
                "toddler_oh": tod_oh, "toddler_wip": tod_wip,
                "toddler_total": tod_oh + tod_wip,
                "boys47_oh": b47_oh, "boys47_wip": b47_wip,
                "boys47_total": b47_oh + b47_wip,
            })
    return result


# ─── Recap Data Builder ──────────────────────────────────────────────────────

def get_recap_data(categories_by_sheet: dict) -> list:
    """Build recap data for the RECAP tab."""
    recap_sections = []
    for sheet_name, sheet_info in categories_by_sheet.items():
        brand = sheet_info["brand"]
        gen_cat = sheet_info.get("general_category", "")
        brand_label = f"{brand} {gen_cat}".strip() if gen_cat else brand

        section_rows = []
        total_oh, total_wip = 0, 0

        for cat in sheet_info["categories"]:
            cat_name = cat["name"]

            # Toddler row (only if non-zero)
            tod_oh, tod_wip = cat["toddler_oh"], cat["toddler_wip"]
            if tod_oh > 0 or tod_wip > 0:
                ref_str = ", ".join(cat["toddler_refs"])
                section_rows.append({
                    "size_range": "TODDLER", "category": cat_name,
                    "ref_nums": ref_str, "oh": tod_oh, "wip": tod_wip,
                })
                total_oh += tod_oh
                total_wip += tod_wip

            # Boys 4-7 row (only if non-zero)
            b47_oh, b47_wip = cat["boys47_oh"], cat["boys47_wip"]
            if b47_oh > 0 or b47_wip > 0:
                ref_str = ", ".join(cat["boys47_refs"])
                section_rows.append({
                    "size_range": "BOYS 4-7", "category": cat_name,
                    "ref_nums": ref_str, "oh": b47_oh, "wip": b47_wip,
                })
                total_oh += b47_oh
                total_wip += b47_wip

        recap_sections.append({
            "brand_label": brand_label, "brand": brand,
            "rows": section_rows,
            "total_oh": total_oh, "total_wip": total_wip,
            "total_ats": total_oh + total_wip,
        })
    return recap_sections
