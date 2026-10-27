"""Compare cached cell VALUES of two Excel packages, offline and dependency-free.

Why this exists
---------------
A file hash answers "is this the same file?", not "did only the intended cells
change?". Saving a workbook through Excel re-serialises styles, shared strings
and the VBA binary, so the hash and the ZIP part list change even when a single
cell was edited. That makes a hash useless as a change-scope check and produces
either false alarms or, worse, a blanket "everything changed, cannot tell".

This tool compares the cached values of every worksheet, so a targeted edit can
be proven to have touched exactly the intended sheet and rows. Values only:
formulas are read as their cached results, formatting is ignored.

No Excel, no openpyxl, no third-party package: OOXML is read from the ZIP with
the standard library.

Usage
-----
  compare_workbook_values.py A.xlsm B.xlsm
  compare_workbook_values.py A.xlsm                      # inventory
  compare_workbook_values.py A.xlsm B.xlsm --sheet "Sheet1" --out r.json
  compare_workbook_values.py A.xlsm B.xlsm --max-rows 20000

Exit codes: 0 identical, 1 differences found, 2 usage/format error.
"""
from __future__ import annotations

import argparse
from decimal import Decimal, InvalidOperation
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from common import validate_new_output, write_new_json

MAX_PACKAGE = 512 * 1024 * 1024
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
WORKSHEET_TYPE = ("http://schemas.openxmlformats.org/officeDocument/"
                  "2006/relationships/worksheet")


class PackageError(ValueError):
    pass


def _col_index(ref: str) -> int:
    """'A1' -> 1, 'AB12' -> 28 (1-based)."""
    m = re.match(r"([A-Z]+)", ref or "")
    if not m:
        raise PackageError("Cell reference without column letters: %r" % ref)
    n = 0
    for ch in m.group(1):
        n = n * 26 + (ord(ch) - 64)
    return n


def _cell_value(cell, shared: list[str]) -> tuple[str, str] | None:
    """A typed cached value. Never coerce exact Excel XML decimals through float."""
    kind = cell.get("t")
    if kind == "inlineStr":
        node = cell.find(NS + "is")
        if node is None:
            return None
        return ("text", "".join(t.text or "" for t in node.iter(NS + "t")))
    v = cell.find(NS + "v")
    if v is None or v.text is None:
        return None
    raw = v.text
    if kind == "s":
        try:
            idx = int(raw)
            if idx < 0:
                raise ValueError("negative shared-string index")
            return ("text", shared[idx])
        except (ValueError, IndexError) as exc:
            raise PackageError("Shared-string index out of range: %r" % raw) from exc
    if kind == "inlineStr" or kind == "str":
        return ("text", raw)
    if kind == "b":
        if raw not in ("0", "1"):
            raise PackageError("Invalid boolean cell value: %r" % raw)
        return ("bool", "TRUE" if raw == "1" else "FALSE")
    if kind == "e":
        return ("error", raw)
    if kind not in (None, "n"):
        raise PackageError("Unsupported cell type: %r" % kind)
    try:
        number = Decimal(raw)
        if not number.is_finite():
            raise InvalidOperation("non-finite numeric value")
        return ("number", "0" if number == 0 else str(number.normalize()))
    except InvalidOperation as exc:
        raise PackageError("Invalid numeric cached value: %r" % raw) from exc


def read_values(package: Path, sheet_filter: str | None = None,
                max_rows: int = 200000) -> dict[str, dict[int, tuple]]:
    """{sheet: {actual Excel row number: cached values by column}}.

    Unlike a sequence of XML <row> elements, row IDs preserve the addresses
    of sparse sheets; moving A1 to A5 must be reported as a difference.
    """
    if max_rows < 1:
        raise PackageError("--max-rows must be positive")
    if package.is_symlink():
        raise PackageError("Symlink inputs are not followed")
    if package.stat().st_size > MAX_PACKAGE:
        raise PackageError("Package exceeds inspection limit")
    if not zipfile.is_zipfile(package):
        raise PackageError("Not an OOXML package (legacy .xls, encrypted or "
                           "corrupt files need a different approved reader)")
    with zipfile.ZipFile(package) as z:
        names = set(z.namelist())
        for needed in ("xl/workbook.xml", "[Content_Types].xml"):
            if needed not in names:
                raise PackageError("Missing %s: not a spreadsheet package" % needed)

        def get_part(name: str) -> bytes:
            if z.getinfo(name).file_size > MAX_PACKAGE:
                raise PackageError("Expanded OOXML part exceeds inspection limit: " + name)
            return z.read(name)

        shared: list[str] = []
        if "xl/sharedStrings.xml" in names:
            for si in ET.fromstring(get_part("xl/sharedStrings.xml")).findall(NS + "si"):
                shared.append("".join(t.text or "" for t in si.iter(NS + "t")))

        rels: dict[str, str] = {}
        if "xl/_rels/workbook.xml.rels" in names:
            for rel in ET.fromstring(get_part("xl/_rels/workbook.xml.rels")):
                if rel.get("Type") == WORKSHEET_TYPE and rel.get("Target"):
                    target = rel.get("Target").lstrip("/")
                    if not target.startswith("xl/"):
                        target = "xl/" + target
                    rels[rel.get("Id")] = target

        sheets: list[tuple[str, str]] = []
        for sh in ET.fromstring(get_part("xl/workbook.xml")).iter(NS + "sheet"):
            rid, name = sh.get(REL_NS + "id"), sh.get("name")
            if not name or not rid or rid not in rels:
                raise PackageError("Worksheet %r has no resolvable part" % name)
            sheets.append((name, rels[rid]))
        if not sheets:
            raise PackageError("Workbook contains no worksheets")
        if sheet_filter is not None and sheet_filter not in {n for n, _ in sheets}:
            raise PackageError("Requested sheet not found: %r" % sheet_filter)

        out: dict[str, dict[int, tuple]] = {}
        for name, part in sheets:
            if sheet_filter is not None and name != sheet_filter:
                continue
            if part not in names:
                raise PackageError("Worksheet part missing: %s" % part)
            rows: dict[int, tuple] = {}
            prev_row = 0
            for count, row in enumerate(ET.fromstring(get_part(part)).iter(NS + "row"), 1):
                if count > max_rows:
                    raise PackageError(
                        "Sheet %r exceeds --max-rows (%d); raise the limit "
                        "explicitly rather than comparing a truncated sheet"
                        % (name, max_rows))
                attr = row.get("r")
                if attr is None:
                    row_id = prev_row + 1
                else:
                    if not attr.isdecimal():
                        raise PackageError("Invalid row number: %r" % attr)
                    row_id = int(attr)
                if row_id <= prev_row or row_id > 1048576:
                    raise PackageError("Out-of-order or invalid Excel row number: %r" % row_id)
                prev_row = row_id
                cells: dict[int, tuple[str, str]] = {}
                for cell in row.findall(NS + "c"):
                    ref = cell.get("r", "")
                    m = re.fullmatch(r"([A-Z]+)([1-9][0-9]*)", ref)
                    if m is None or int(m.group(2)) != row_id:
                        raise PackageError("Cell address does not match row: %r" % ref)
                    col = _col_index(ref)
                    if col > 16384:
                        raise PackageError("Column outside Excel limit: %r" % ref)
                    if col in cells:
                        raise PackageError("Duplicate cell address: %r" % ref)
                    val = _cell_value(cell, shared)
                    if val is not None:
                        cells[col] = val
                # Styled blank rows are not cached values; ignore them.
                if cells:
                    last = max(cells)
                    rows[row_id] = tuple(cells.get(i) for i in range(1, last + 1))
            out[name] = rows
        return out


def compare(a: Path, b: Path, sheet_filter: str | None, max_rows: int) -> dict:
    va, vb = read_values(a, sheet_filter, max_rows), read_values(b, sheet_filter, max_rows)
    total_a = sum(sum(v is not None for v in r) for rows in va.values() for r in rows.values())
    total_b = sum(sum(v is not None for v in r) for rows in vb.values() for r in rows.values())
    if not va or not vb or total_a == 0 or total_b == 0:
        raise PackageError("Read %d/%d cached values - refusing to report equality "
                           "on an empty read" % (total_a, total_b))
    only_a = sorted(set(va) - set(vb))
    only_b = sorted(set(vb) - set(va))
    changed: list[dict] = []
    for name in sorted(set(va) & set(vb)):
        ra, rb = va[name], vb[name]
        all_row_ids = sorted(set(ra) | set(rb))
        diff = [i for i in all_row_ids if ra.get(i) != rb.get(i)]
        if diff:
            changed.append({"sheet": name, "rows_a": len(ra), "rows_b": len(rb),
                            "differing_row_count": len(diff),
                            "first_differing_rows": diff[:15]})
    return {
        "status": ("IDENTICAL" if not changed and not only_a and not only_b
                   else "DIFFERS"),
        "a": str(a), "b": str(b),
        "sheets_a": len(va), "sheets_b": len(vb),
        "cells_a": total_a, "cells_b": total_b,
        "only_in_a": only_a, "only_in_b": only_b,
        "changed_sheets": changed,
        "scope": ("Cached cell values and their actual Excel row/column addresses. "
                  "Formula expressions, formatting, charts, VBA, print settings "
                  "and defined names are NOT compared."),
    }


def main() -> int:
    p = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("a", type=Path)
    p.add_argument("b", type=Path, nargs="?")
    p.add_argument("--sheet", help="restrict to one worksheet name")
    p.add_argument("--max-rows", type=int, default=200000)
    p.add_argument("--out", type=Path, help="write JSON report to a NEW path")
    a = p.parse_args()
    try:
        for f in (a.a, a.b):
            if f and not f.is_file():
                raise PackageError("File not found: %s" % f)
        validate_new_output(a.out)
        if a.b:
            report = compare(a.a, a.b, a.sheet, a.max_rows)
        else:
            vals = read_values(a.a, a.sheet, a.max_rows)
            if not vals:
                raise PackageError("No sheets read")
            report = {"status": "INVENTORY", "sheets": len(vals),
                      "inventory": {k: len(v) for k, v in vals.items()}}
        report["tool"] = "compare_workbook_values.py"
        text = json.dumps(report, indent=2, ensure_ascii=False)
        if a.out:
            write_new_json(a.out, report)
        print(text)
        return 0 if report["status"] in ("IDENTICAL", "INVENTORY") else 1
    except (OSError, ValueError, ET.ParseError, zipfile.BadZipFile) as exc:
        print("ERROR: %s" % exc, file=sys.stderr)
        return 2


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
