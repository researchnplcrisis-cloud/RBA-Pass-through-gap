import datetime as dt
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def col_to_idx(cell_ref):
    match = re.match(r"([A-Z]+)", cell_ref)
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
        parts = []
        for text in item.findall(".//a:t", NS):
            parts.append(text.text or "")
        strings.append("".join(parts))
    return strings


def first_sheet_path(zf):
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {}
    for rel in rels:
        rel_map[rel.attrib["Id"]] = rel.attrib["Target"]
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


def read_rows(path):
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
                max_col = max(values)
                rows.append([values.get(idx, "") for idx in range(max_col + 1)])
            else:
                rows.append([])
        return rows


def main():
    for arg in sys.argv[1:]:
        path = Path(arg)
        rows = read_rows(path)
        print(f"\n=== {path.name}: rows 0-12 ===")
        for idx, row in enumerate(rows[:13]):
            print(idx, row)


if __name__ == "__main__":
    main()
