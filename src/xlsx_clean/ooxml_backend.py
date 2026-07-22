"""Surgical OOXML edits that preserve non-worksheet package parts.

Unlike openpyxl load/save, this rewrites only worksheet XML (and optionally
drops calcChain) so connections, queries, styles, drawings, etc. stay intact.
"""

from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WORKSHEET_REL_TYPE = f"{OFFICE_REL_NS}/worksheet"

ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", OFFICE_REL_NS)

NS = {"m": MAIN_NS, "r": REL_NS, "pr": OFFICE_REL_NS}

_CELL_REF_RE = re.compile(r"^([A-Za-z]+)(\d+)$")


def _q(tag: str) -> str:
    return f"{{{MAIN_NS}}}{tag}"


def col_letters_to_index(letters: str) -> int:
    result = 0
    for ch in letters.upper():
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result


def index_to_col_letters(index: int) -> str:
    letters = []
    while index:
        index, rem = divmod(index - 1, 26)
        letters.append(chr(ord("A") + rem))
    return "".join(reversed(letters))


def parse_cell_ref(ref: str) -> tuple[str, int]:
    match = _CELL_REF_RE.match(ref.strip())
    if not match:
        raise ValueError(f"Invalid cell reference: {ref!r}")
    return match.group(1).upper(), int(match.group(2))


def expand_a1_range(a1: str) -> list[str]:
    """Expand 'F2', 'F2:F3', or 'E12:G12' into individual A1 refs."""
    a1 = a1.strip()
    if ":" not in a1:
        col, row = parse_cell_ref(a1)
        return [f"{col}{row}"]

    start, end = a1.split(":", 1)
    start_col, start_row = parse_cell_ref(start)
    end_col, end_row = parse_cell_ref(end)
    c0, c1 = col_letters_to_index(start_col), col_letters_to_index(end_col)
    r0, r1 = start_row, end_row
    if c0 > c1:
        c0, c1 = c1, c0
    if r0 > r1:
        r0, r1 = r1, r0

    refs: list[str] = []
    for row in range(r0, r1 + 1):
        for col in range(c0, c1 + 1):
            refs.append(f"{index_to_col_letters(col)}{row}")
    return refs


def parse_sheet_cell_ref(spec: str) -> tuple[str, str]:
    """Parse 'Batch!F2:F3' or \"'KW'!C3:C4\" into (sheet_name, a1_range)."""
    spec = spec.strip()
    if "!" not in spec:
        raise ValueError(f"Expected Sheet!A1 reference, got {spec!r}")
    sheet, a1 = spec.split("!", 1)
    sheet = sheet.strip().strip("'").strip('"')
    return sheet, a1.strip()


def _sheet_targets(zf: zipfile.ZipFile) -> dict[str, str]:
    """Map sheet name -> zip member path (e.g. xl/worksheets/sheet1.xml)."""
    workbook_xml = ET.fromstring(zf.read("xl/workbook.xml"))
    rels_xml = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))

    rel_id_to_target: dict[str, str] = {}
    for rel in rels_xml.findall("r:Relationship", NS):
        if rel.get("Type") != WORKSHEET_REL_TYPE:
            continue
        rid = rel.get("Id")
        target = rel.get("Target")
        if not rid or not target:
            continue
        # Targets are relative to xl/
        target = target.lstrip("/")
        if not target.startswith("xl/"):
            target = f"xl/{target}"
        rel_id_to_target[rid] = target

    name_to_path: dict[str, str] = {}
    sheets = workbook_xml.find("m:sheets", NS)
    if sheets is None:
        return name_to_path
    for sheet in sheets.findall("m:sheet", NS):
        name = sheet.get("name")
        rid = sheet.get(f"{{{OFFICE_REL_NS}}}id")
        if name and rid and rid in rel_id_to_target:
            name_to_path[name] = rel_id_to_target[rid]
    return name_to_path


def _find_row(sheet_root: ET.Element, row_num: int) -> ET.Element | None:
    sheet_data = sheet_root.find("m:sheetData", NS)
    if sheet_data is None:
        return None
    for row in sheet_data.findall("m:row", NS):
        if row.get("r") == str(row_num):
            return row
    return None


def _ensure_sheet_data(sheet_root: ET.Element) -> ET.Element:
    sheet_data = sheet_root.find("m:sheetData", NS)
    if sheet_data is not None:
        return sheet_data
    sheet_data = ET.SubElement(sheet_root, _q("sheetData"))
    return sheet_data


def _ensure_row(sheet_root: ET.Element, row_num: int) -> ET.Element:
    existing = _find_row(sheet_root, row_num)
    if existing is not None:
        return existing

    sheet_data = _ensure_sheet_data(sheet_root)
    new_row = ET.Element(_q("row"), {"r": str(row_num)})
    # Insert in row-number order
    inserted = False
    for idx, row in enumerate(list(sheet_data)):
        if row.tag != _q("row"):
            continue
        r_attr = row.get("r")
        if r_attr is not None and int(r_attr) > row_num:
            sheet_data.insert(idx, new_row)
            inserted = True
            break
    if not inserted:
        sheet_data.append(new_row)
    return new_row


def _find_cell(row: ET.Element, ref: str) -> ET.Element | None:
    for cell in row.findall("m:c", NS):
        if cell.get("r") == ref:
            return cell
    return None


def _ensure_cell(row: ET.Element, ref: str) -> ET.Element:
    existing = _find_cell(row, ref)
    if existing is not None:
        return existing

    col_letters, _ = parse_cell_ref(ref)
    col_idx = col_letters_to_index(col_letters)
    new_cell = ET.Element(_q("c"), {"r": ref})
    inserted = False
    for idx, cell in enumerate(list(row)):
        if cell.tag != _q("c"):
            continue
        cell_ref = cell.get("r")
        if not cell_ref:
            continue
        other_col, _ = parse_cell_ref(cell_ref)
        if col_letters_to_index(other_col) > col_idx:
            row.insert(idx, new_cell)
            inserted = True
            break
    if not inserted:
        row.append(new_cell)
    return new_cell


def _clear_cell_contents(cell: ET.Element) -> None:
    """Clear value/formula like Excel ClearContents; keep style attrs."""
    # Drop type/value-related attrs; keep r and s (style).
    for attr in list(cell.attrib):
        if attr not in ("r", "s", "cm", "vm", "ph"):
            del cell.attrib[attr]
    for child in list(cell):
        cell.remove(child)


def _set_cell_string(cell: ET.Element, value: str) -> None:
    """Set an inline string value, preserving style attribute when present."""
    style = cell.get("s")
    _clear_cell_contents(cell)
    if style is not None:
        cell.set("s", style)
    if value == "":
        # Empty string: leave cell present with style only (cleared).
        return
    cell.set("t", "inlineStr")
    is_el = ET.SubElement(cell, _q("is"))
    t_el = ET.SubElement(is_el, _q("t"))
    t_el.text = value
    if value.startswith(" ") or value.endswith(" ") or "\n" in value:
        t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def _sheet_xml_tostring(root: ET.Element) -> bytes:
    # Preserve XML declaration; Excel is picky about namespaces on worksheets.
    body = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    return body


def _apply_ops_to_sheet(
    sheet_xml: bytes,
    clear_refs: Iterable[str],
    set_values: dict[str, str],
) -> bytes:
    root = ET.fromstring(sheet_xml)

    for ref in clear_refs:
        col, row_num = parse_cell_ref(ref)
        row = _find_row(root, row_num)
        if row is None:
            continue
        cell = _find_cell(row, f"{col}{row_num}")
        if cell is not None:
            _clear_cell_contents(cell)

    for ref, value in set_values.items():
        col, row_num = parse_cell_ref(ref)
        row = _ensure_row(root, row_num)
        cell = _ensure_cell(row, f"{col}{row_num}")
        _set_cell_string(cell, value)

    return _sheet_xml_tostring(root)


def apply_cell_edits(
    source: Path | str,
    destination: Path | str,
    clear_specs: Iterable[str],
    set_specs: dict[str, str],
) -> None:
    """Copy source workbook to destination with surgical cell edits.

    clear_specs: iterable of 'Sheet!A1:B2' ranges to ClearContents.
    set_specs: mapping of 'Sheet!A1' -> string value to write.
    """
    source = Path(source)
    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)

    clears_by_sheet: dict[str, list[str]] = {}
    for spec in clear_specs:
        sheet, a1 = parse_sheet_cell_ref(spec)
        clears_by_sheet.setdefault(sheet, []).extend(expand_a1_range(a1))

    sets_by_sheet: dict[str, dict[str, str]] = {}
    for spec, value in set_specs.items():
        sheet, a1 = parse_sheet_cell_ref(spec)
        refs = expand_a1_range(a1)
        sheet_sets = sets_by_sheet.setdefault(sheet, {})
        for ref in refs:
            sheet_sets[ref] = value

    touched_sheets = set(clears_by_sheet) | set(sets_by_sheet)

    with zipfile.ZipFile(source, "r") as zin:
        name_to_path = _sheet_targets(zin)
        missing = sorted(touched_sheets - set(name_to_path))
        if missing:
            raise KeyError(f"Worksheet(s) not found in workbook: {missing}")

        rewritten: dict[str, bytes] = {}
        for sheet_name in touched_sheets:
            member = name_to_path[sheet_name]
            rewritten[member] = _apply_ops_to_sheet(
                zin.read(member),
                clears_by_sheet.get(sheet_name, []),
                sets_by_sheet.get(sheet_name, {}),
            )

        # Drop calcChain so Excel rebuilds after formula clears, and scrub
        # dangling Content_Types / workbook rels entries that pointed at it.
        skip = {"xl/calcChain.xml"}
        if "xl/calcChain.xml" in zin.namelist():
            if "[Content_Types].xml" in zin.namelist():
                rewritten["[Content_Types].xml"] = _strip_calc_chain_content_types(
                    zin.read("[Content_Types].xml")
                )
            if "xl/_rels/workbook.xml.rels" in zin.namelist():
                rewritten["xl/_rels/workbook.xml.rels"] = _strip_calc_chain_rels(
                    zin.read("xl/_rels/workbook.xml.rels")
                )

        with zipfile.ZipFile(destination, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                if info.filename in skip:
                    continue
                if info.filename in rewritten:
                    zout.writestr(info, rewritten[info.filename])
                else:
                    zout.writestr(info, zin.read(info.filename))


def _strip_calc_chain_content_types(data: bytes) -> bytes:
    """Remove calcChain Override without rewriting package namespaces."""
    text = data.decode("utf-8")
    text = re.sub(
        r'\s*<[^>]*PartName="[^"]*calcChain\.xml"[^>]*/>',
        "",
        text,
        flags=re.IGNORECASE,
    )
    return text.encode("utf-8")


def _strip_calc_chain_rels(data: bytes) -> bytes:
    """Remove calcChain Relationship without rewriting package namespaces."""
    text = data.decode("utf-8")
    text = re.sub(
        r'\s*<[^>]*Target="[^"]*calcChain\.xml"[^>]*/>',
        "",
        text,
        flags=re.IGNORECASE,
    )
    return text.encode("utf-8")


def clean_workbook_ooxml(
    source: Path | str,
    destination: Path | str,
    cells_to_clear: str,
    notes_cell: str,
    serial_cell: str,
    batch_serial: str,
    notes_value: str = "",
) -> None:
    """High-level OOXML backend matching the COM clean sequence."""
    clear_specs = [part.strip() for part in cells_to_clear.split(",") if part.strip()]
    set_specs: dict[str, str] = {}
    for part in notes_cell.split(","):
        part = part.strip()
        if part:
            set_specs[part] = notes_value
    for part in serial_cell.split(","):
        part = part.strip()
        if part:
            set_specs[part] = batch_serial
    apply_cell_edits(source, destination, clear_specs, set_specs)
