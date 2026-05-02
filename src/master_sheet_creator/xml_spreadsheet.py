"""Parse Excel 2003 XML (SpreadsheetML) workbooks, often saved as ``.xls``."""

from __future__ import annotations

import xml.etree.ElementTree as ET

import pandas as pd

NS = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
_CELL_INDEX = "{urn:schemas-microsoft-com:office:spreadsheet}Index"


def _cell_text(cell: ET.Element) -> str:
    data = cell.find("ss:Data", NS)
    if data is None:
        return ""
    return "" if data.text is None else str(data.text)


def _row_cell_strings(row: ET.Element) -> list[str]:
    """
    Cell values left-to-right.

    Uses optional ``ss:Index`` (1-based column) when present; otherwise advances
    sequentially from the previous column (Excel 2003 XML rules).
    """
    cells = row.findall("ss:Cell", NS)
    if not cells:
        return []

    current = 1
    sparse: dict[int, str] = {}
    max_pos = 0

    for cell in cells:
        if _CELL_INDEX in cell.attrib:
            current = int(cell.attrib[_CELL_INDEX])
        val = _cell_text(cell)
        sparse[current] = val
        max_pos = max(max_pos, current)
        current += 1

    return [sparse.get(i, "") for i in range(1, max_pos + 1)]


def load_xml_spreadsheet(raw_bytes: bytes) -> pd.DataFrame:
    """
    Load the first worksheet table from Excel 2003 XML into a DataFrame.

    Expects ``Workbook`` → ``Worksheet`` → ``Table`` → ``Row`` structure.
    """
    text = raw_bytes.decode("utf-8-sig")
    root = ET.fromstring(text)

    table = root.find(".//ss:Table", NS)
    if table is None:
        raise ValueError("No Table element found in XML spreadsheet")

    xml_rows = table.findall("ss:Row", NS)
    if not xml_rows:
        raise ValueError("No rows found in XML spreadsheet")

    row_lists = [_row_cell_strings(r) for r in xml_rows]
    width = max((len(r) for r in row_lists), default=0)

    grid = [
        r + [""] * (width - len(r)) if len(r) < width else r[:width] for r in row_lists
    ]

    header = grid[0]
    body = grid[1:]

    seen: dict[str, int] = {}
    uniq_header: list[str] = []
    for h in header:
        base = h if h != "" else "Column"
        if base in seen:
            seen[base] += 1
            uniq_header.append(f"{base}.{seen[base]}")
        else:
            seen[base] = 0
            uniq_header.append(base)

    return pd.DataFrame(body, columns=uniq_header)
