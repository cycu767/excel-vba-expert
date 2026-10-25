"""Read-only VBA/VBScript heuristic review, not a parser, compiler or security scanner."""
from __future__ import annotations
import argparse
import hashlib
import json
import re
import sys
from pathlib import Path
from common import decode_source, write_new_json

RULES = [
 ("ACTIVE-CONTEXT", r"\b(?:ActiveWorkbook|ActiveSheet|ActiveCell|Selection)\b", "review", "Validate intentional UI context and capture the correct owner."),
 ("UNQUALIFIED-RANGE", r"(?<![.\w])(?:Range|Cells|Rows|Columns|Worksheets|Sheets)\s*\(", "review", "Potential implicit owner; inspect module context and qualify deliberately."),
 ("ERROR-SUPPRESSION", r"\bOn\s+Error\s+Resume\s+Next\b", "review", "Verify narrow scope, immediate Err capture and restoration of handler."),
 ("HARD-END", r"(?:^|:)\s*End\s*(?::|$)", "high", "End may bypass cleanup; distinguish from End Sub/If/Function."),
 ("API-DECLARATION", r"\bDeclare\s+(?!PtrSafe\b)(?:Function|Sub)\b", "review", "Review conditional-compilation branch and native widths; no automatic conversion."),
 ("STATE-MUTATION", r"\bApplication\.(?:EnableEvents|ScreenUpdating|DisplayAlerts|Calculation)\s*=", "review", "Trace capture and all cleanup paths; restoration cannot be proved by text matching."),
 ("DESTRUCTIVE-IO", r"\b(?:Kill|RmDir)\s+|\.Delete\b|\.SaveAs\b", "high", "Review exact target, permission, backup and overwrite policy."),
 ("EXTERNAL-EXECUTION", r"\bShell\b|\bCreateObject\s*\(", "review", "Verify component/process intent and security; occurrence alone is not malicious."),
 ("REENTRY", r"\bDoEvents\b", "review", "Check re-entry guard and transaction boundary."),
 ("EXCEL-QUIT", r"\.Quit\b", "high", "Confirm ownership; never terminate a borrowed Excel instance."),
]

def strip_strings_comments(line: str) -> str:
    out = []
    i = 0
    quoted = False
    while i < len(line):
        ch = line[i]
        if ch == '"':
            if quoted and i + 1 < len(line) and line[i + 1] == '"':
                out.extend("  "); i += 2; continue
            quoted = not quoted
            out.append(" ")
        elif ch == "'" and not quoted:
            break
        else:
            out.append(" " if quoted else ch)
        i += 1
    text = "".join(out)
    return re.split(r"(?:^|:)\s*Rem(?:\s|$)", text, maxsplit=1, flags=re.I)[0]

def inspect_text(text: str, filename: str = "source.bas") -> list[dict]:
    clean = [strip_strings_comments(line) for line in text.splitlines()]
    findings = []
    if not any(re.match(r"^\s*Option\s+Explicit\s*$", line, re.I) for line in clean):
        findings.append({"rule": "OPTION-EXPLICIT", "severity": "review", "line": 1, "message": "No Option Explicit found in exported text."})
    for number, line in enumerate(clean, 1):
        for rule, pattern, severity, message in RULES:
            if re.search(pattern, line, re.I):
                findings.append({"rule": rule, "severity": severity, "line": number, "message": message})
    return [dict(item, file=filename, confidence="heuristic_candidate_not_confirmed") for item in findings]

def inspect_file(path: Path, encoding: str = "auto") -> dict:
    if path.is_symlink():
        raise ValueError("Symlink inputs are not followed")
    if path.suffix.lower() not in {".bas", ".cls", ".frm", ".vbs"}:
        raise ValueError("Expected .bas, .cls, .frm or .vbs exported source")
    if path.stat().st_size > 8 * 1024 * 1024:
        raise ValueError("Source exceeds 8 MiB safety limit")
    raw = path.read_bytes()
    text, used = decode_source(raw, encoding)
    report = {"path": str(path), "sha256": hashlib.sha256(raw).hexdigest(), "encoding": used,
              "lines": len(text.splitlines()), "findings": inspect_text(text, path.name)}
    if path.suffix.lower() == ".frm":
        resources = re.findall(r'"([^"\r\n]+\.frx)"', text, re.I)
        report["form_resources"] = [{"name": name, "exists": (path.parent / name).is_file() if Path(name).name == name else False,
                                     "note": "Presence only; binary layout not decoded"} for name in sorted(set(resources))]
    return report

def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("paths", nargs="+", type=Path)
    p.add_argument("--encoding", default="auto")
    p.add_argument("--out", type=Path)
    a = p.parse_args()
    try:
        result = {"evidence_level": "STATIC_HEURISTIC", "warning": "No compilation, call-graph proof, execution or automatic fixes.",
                  "files": [inspect_file(path, a.encoding) for path in a.paths]}
        if a.out:
            write_new_json(a.out, result)
        else:
            print(json.dumps(result, indent=2, ensure_ascii=False))
        return 0
    except (OSError, ValueError, UnicodeError, LookupError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr); return 2

if __name__ == "__main__":
    raise SystemExit(main())
