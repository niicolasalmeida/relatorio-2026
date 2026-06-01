from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

NS = {"m": MAIN_NS}
DATE_STYLE_IDS = {
    14, 15, 16, 17, 18, 19, 20, 21, 22,
    27, 30, 36, 45, 46, 47, 50, 57,
}
DATE_FORMAT_CHARS = set("ymdhs")


def _column_index(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha()).upper()
    value = 0
    for ch in letters:
        value = value * 26 + (ord(ch) - 64)
    return value


def _excel_date(value: float):
    base = dt.datetime(1899, 12, 30)
    whole = int(value)
    fraction = float(value) - whole
    result = base + dt.timedelta(days=whole, seconds=round(fraction * 86400))
    return result.date() if fraction == 0 else result


def _is_date_format(code: str) -> bool:
    lowered = code.lower()
    lowered = re.sub(r'"[^"]*"', "", lowered)
    lowered = re.sub(r"\[[^\]]*\]", "", lowered)
    return any(ch in DATE_FORMAT_CHARS for ch in lowered)


@dataclass
class Cell:
    value: object
    column: int | None = None


class Worksheet:
    def __init__(self, rows: list[list[object]]):
        self._rows = rows
        self.max_row = len(rows)
        self.max_column = max((len(row) for row in rows), default=0)

    def iter_rows(self, min_row=1, max_row=None, values_only=False):
        max_row = self.max_row if max_row is None else min(max_row, self.max_row)
        for row_idx in range(min_row - 1, max_row):
            row = self._rows[row_idx] if row_idx < len(self._rows) else []
            padded = row + [None] * max(0, self.max_column - len(row))
            if values_only:
                yield tuple(padded)
            else:
                yield tuple(Cell(v, idx + 1) for idx, v in enumerate(padded))

    def cell(self, row: int, column: int) -> Cell:
        try:
            value = self._rows[row - 1][column - 1]
        except IndexError:
            value = None
        return Cell(value, column)

    def __getitem__(self, row: int):
        if not isinstance(row, int):
            raise TypeError("Worksheet indices must be integers")
        if row < 1:
            raise IndexError("Worksheet row indices start at 1")
        values = self._rows[row - 1] if row - 1 < len(self._rows) else []
        padded = values + [None] * max(0, self.max_column - len(values))
        return tuple(Cell(v, idx + 1) for idx, v in enumerate(padded))


class Workbook:
    def __init__(self, sheets: dict[str, Worksheet]):
        self._sheets = sheets
        self.sheetnames = list(sheets.keys())

    def __getitem__(self, name: str) -> Worksheet:
        return self._sheets[name]


def _read_shared_strings(zf: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    out = []
    for si in root.findall("m:si", NS):
        text = "".join(t.text or "" for t in si.iterfind(".//m:t", NS))
        out.append(text)
    return out


def _read_date_styles(zf: ZipFile) -> set[int]:
    if "xl/styles.xml" not in zf.namelist():
        return set()
    root = ET.fromstring(zf.read("xl/styles.xml"))
    numfmts = {}
    numfmts_node = root.find("m:numFmts", NS)
    if numfmts_node is not None:
        for fmt in numfmts_node.findall("m:numFmt", NS):
            fmt_id = int(fmt.attrib["numFmtId"])
            code = fmt.attrib.get("formatCode", "")
            numfmts[fmt_id] = _is_date_format(code)

    styles = set()
    cellxfs = root.find("m:cellXfs", NS)
    if cellxfs is None:
        return styles
    for idx, xf in enumerate(cellxfs.findall("m:xf", NS)):
        numfmt_id = int(xf.attrib.get("numFmtId", "0"))
        if numfmt_id in DATE_STYLE_IDS or numfmts.get(numfmt_id, False):
            styles.add(idx)
    return styles


def _sheet_paths(zf: ZipFile):
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels.findall(f"{{{PKG_REL_NS}}}Relationship")
    }
    sheets = {}
    for sheet in wb.find("m:sheets", NS):
        rid = sheet.attrib[f"{{{REL_NS}}}id"]
        target = rel_map[rid]
        if not target.startswith("xl/"):
            target = f"xl/{target}"
        sheets[sheet.attrib["name"]] = target
    return sheets


def _parse_sheet(zf: ZipFile, path: str, shared_strings: list[str], date_styles: set[int]) -> Worksheet:
    root = ET.fromstring(zf.read(path))
    data = []
    max_col = 0
    for row_node in root.findall(".//m:sheetData/m:row", NS):
        row_values = []
        current_col = 1
        for cell in row_node.findall("m:c", NS):
            ref = cell.attrib.get("r", "")
            col_idx = _column_index(ref) if ref else current_col
            while current_col < col_idx:
                row_values.append(None)
                current_col += 1

            ctype = cell.attrib.get("t")
            style_idx = int(cell.attrib.get("s", "0"))
            value = None

            if ctype == "inlineStr":
                is_node = cell.find("m:is", NS)
                if is_node is not None:
                    value = "".join(t.text or "" for t in is_node.iterfind(".//m:t", NS))
            else:
                v = cell.findtext("m:v", default=None, namespaces=NS)
                if v is not None:
                    if ctype == "s":
                        value = shared_strings[int(v)]
                    elif ctype == "b":
                        value = v == "1"
                    elif ctype == "str":
                        value = v
                    else:
                        num = float(v)
                        if style_idx in date_styles:
                            value = _excel_date(num)
                        elif num.is_integer():
                            value = int(num)
                        else:
                            value = num

            row_values.append(value)
            current_col += 1

        max_col = max(max_col, len(row_values))
        data.append(row_values)

    for row in data:
        row.extend([None] * (max_col - len(row)))
    return Worksheet(data)


def load_workbook(path, read_only=True, data_only=True):
    del read_only, data_only
    xlsx_path = Path(path)
    with ZipFile(xlsx_path) as zf:
        shared_strings = _read_shared_strings(zf)
        date_styles = _read_date_styles(zf)
        sheets = {
            name: _parse_sheet(zf, sheet_path, shared_strings, date_styles)
            for name, sheet_path in _sheet_paths(zf).items()
        }
    return Workbook(sheets)
