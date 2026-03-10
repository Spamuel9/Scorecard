#!/usr/bin/env python3
"""Apply row-specific input messages to indicator cells in Endpoint Security workbook.

This keeps the workbook as .xlsx and edits worksheet XML directly so the change is
reviewable as text in pull requests (script-based change instead of binary diff).
"""

from __future__ import annotations

import argparse
import uuid
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_XR = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
NS = {"a": NS_MAIN}

ET.register_namespace("", NS_MAIN)
ET.register_namespace("xr", NS_XR)


def _shared_strings(files: dict[str, bytes]) -> list[str]:
    root = ET.fromstring(files["xl/sharedStrings.xml"])
    out: list[str] = []
    for si in root.findall("a:si", NS):
        out.append("".join(t.text or "" for t in si.findall(".//a:t", NS)))
    return out


def _rows(root: ET.Element) -> dict[int, dict[str, tuple[str, str]]]:
    rows: dict[int, dict[str, tuple[str, str]]] = {}
    for row in root.find("a:sheetData", NS).findall("a:row", NS):
        row_idx = int(row.attrib["r"])
        row_vals: dict[str, tuple[str, str]] = {}
        for cell in row.findall("a:c", NS):
            ref = cell.attrib["r"]
            col = "".join(ch for ch in ref if ch.isalpha())
            cell_type = cell.attrib.get("t", "")
            value = cell.find("a:v", NS)
            row_vals[col] = (cell_type, (value.text if value is not None else "") or "")
        rows[row_idx] = row_vals
    return rows


def _resolve_text(row_map: dict[int, dict[str, tuple[str, str]]], shared: list[str], row: int, col: str) -> str:
    cell_type, value = row_map.get(row, {}).get(col, ("", ""))
    if cell_type == "s" and value:
        return shared[int(value)]
    return value


def apply_messages(workbook_path: Path) -> int:
    with zipfile.ZipFile(workbook_path, "r") as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    sheet_path = "xl/worksheets/sheet2.xml"
    root = ET.fromstring(files[sheet_path])
    rows = _rows(root)
    shared = _shared_strings(files)

    data_validations = root.find("a:dataValidations", NS)
    if data_validations is None:
        data_validations = ET.SubElement(root, f"{{{NS_MAIN}}}dataValidations")

    # Keep non-indicator validations (e.g., existing drop-downs on D/E columns).
    retained: list[ET.Element] = []
    for dv in data_validations.findall("a:dataValidation", NS):
        sqref = dv.attrib.get("sqref", "")
        if not sqref.startswith("A"):
            retained.append(dv)

    for child in list(data_validations):
        data_validations.remove(child)

    for dv in retained:
        data_validations.append(dv)

    # Add A2:A57 messages based on Measure (B) and Reference (C).
    for row in range(2, 58):
        measure = _resolve_text(rows, shared, row, "B")
        reference = _resolve_text(rows, shared, row, "C")
        attrs = {
            "allowBlank": "1",
            "showInputMessage": "1",
            "showErrorMessage": "1",
            "promptTitle": measure,
            "prompt": reference,
            "sqref": f"A{row}",
            f"{{{NS_XR}}}uid": "{" + str(uuid.uuid4()).upper() + "}",
        }
        ET.SubElement(data_validations, f"{{{NS_MAIN}}}dataValidation", attrs)

    data_validations.set("count", str(len(data_validations.findall("a:dataValidation", NS))))

    files[sheet_path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    with zipfile.ZipFile(workbook_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, payload in files.items():
            zout.writestr(name, payload)

    return len(data_validations.findall("a:dataValidation", NS))


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("workbook", nargs="?", default="Endpoint Security Test1.xlsx")
    args = parser.parse_args()

    count = apply_messages(Path(args.workbook))
    print(f"Updated data validations: {count}")


if __name__ == "__main__":
    main()
