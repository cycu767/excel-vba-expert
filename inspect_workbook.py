"""Inspect OOXML/ZIP part names without Excel, macro execution or content extraction."""
from __future__ import annotations
import argparse
import json
import sys
import zipfile
from pathlib import Path, PurePosixPath
from common import sha256_file, write_new_json

MAX_FILE = 512 * 1024 * 1024
MAX_ENTRIES = 100000
MAX_DECLARED_UNCOMPRESSED = 2 * 1024 * 1024 * 1024

def inspect_package(path: Path) -> dict:
    if path.is_symlink():
        raise ValueError("Symlink inputs are not followed")
    if path.stat().st_size > MAX_FILE:
        raise ValueError("Package exceeds 512 MiB inspection limit")
    digest = sha256_file(path)
    if not zipfile.is_zipfile(path):
        return {"status": "UNSUPPORTED_OR_ENCRYPTED", "sha256": digest,
                "warning": "Not a ZIP package: legacy OLE, encryption, corruption or other format. Not a safety verdict."}
    with zipfile.ZipFile(path) as z:
        infos = z.infolist()
        if len(infos) > MAX_ENTRIES or sum(i.file_size for i in infos) > MAX_DECLARED_UNCOMPRESSED:
            raise ValueError("Archive exceeds metadata safety limits")
        names = [i.filename for i in infos]
        if any("\\" in n or n.startswith("/") or ".." in PurePosixPath(n).parts or ":" in n for n in names):
            raise ValueError("Unsafe archive member path")
        if len({n.casefold() for n in names}) != len(names):
            raise ValueError("Duplicate/case-colliding archive member names")
        if any(i.flag_bits & 1 for i in infos):
            raise ValueError("Encrypted ZIP members are unsupported")
        nlower = [n.casefold() for n in names]
        signatures = {
            "vba_project": lambda n: n.endswith("vbaproject.bin"),
            "xlm_macrosheets": lambda n: n.startswith("xl/macrosheets/"),
            "activex": lambda n: n.startswith("xl/activex/"),
            "embedded_objects": lambda n: n.startswith("xl/embeddings/"),
            "external_links": lambda n: n.startswith("xl/externallinks/"),
            "connections": lambda n: n == "xl/connections.xml",
            "custom_ui": lambda n: n.startswith("customui/"),
            "signature_related": lambda n: "signature" in n or n.startswith("_xmlsignatures/"),
        }
        parts = {key: [names[i] for i,n in enumerate(nlower) if predicate(n)] for key,predicate in signatures.items()}
        workbook = any(n in {"xl/workbook.xml", "xl/workbook.bin"} for n in nlower)
        return {"status": "METADATA_ONLY", "sha256": digest, "zip_entries": len(names),
                "looks_like_excel_package": workbook and "[content_types].xml" in nlower,
                "part_groups": parts, "source_modified": False,
                "warning": "No VBA decoding, XML semantic inspection, execution, signature verification or malware verdict. No detected macro part does not establish safety."}

def main() -> int:
    p=argparse.ArgumentParser(description=__doc__);p.add_argument("path",type=Path);p.add_argument("--out",type=Path);a=p.parse_args()
    try:
        result=inspect_package(a.path)
        if a.out: write_new_json(a.out,result)
        else: print(json.dumps(result,indent=2))
        return 0 if result["status"]=="METADATA_ONLY" else 1
    except (OSError,ValueError,zipfile.BadZipFile) as exc:
        print(f"ERROR: {exc}",file=sys.stderr);return 2

if __name__=="__main__":
    raise SystemExit(main())
