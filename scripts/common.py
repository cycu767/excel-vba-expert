"""Local-only helpers. Python 3.10+, standard library; no network or Excel."""
from __future__ import annotations
import hashlib
import json
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]

def read_json(path: Path) -> Any:
    with path.open(encoding="utf-8-sig") as stream:
        return json.load(stream)

def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()

def write_new_json(path: Path, data: Any) -> None:
    """Never replace an existing report. Exclusive creation, UTF-8, deterministic keys."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8", newline="\n") as stream:
        json.dump(data, stream, indent=2, ensure_ascii=False)
        stream.write("\n")

def decode_source(raw: bytes, encoding: str = "auto") -> tuple[str, str]:
    if encoding != "auto":
        return raw.decode(encoding, errors="strict"), encoding
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw.decode("utf-16", errors="strict"), "utf-16-bom"
    try:
        return raw.decode("utf-8-sig", errors="strict"), "utf-8-sig"
    except UnicodeDecodeError as exc:
        raise ValueError("Unknown legacy encoding; specify --encoding cp1250 or the actual code page. No automatic lossy conversion.") from exc


def validate_new_output(path: Path | None) -> None:
    """Preflight before side effects; exclusive creation still protects actual writes."""
    if path is not None and (path.exists() or path.is_symlink()):
        raise FileExistsError(f"Output already exists; use a new path: {path}")
