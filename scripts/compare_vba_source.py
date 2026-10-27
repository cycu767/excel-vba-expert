"""Compare VBA component SOURCE between macro-enabled packages, offline.

Why this exists
---------------
Identical `vbaProject.bin` bytes prove identity of that part, but a binary
difference does not prove a source change: Excel may re-serialise its OLE data.
Reading `Workbook.VBProject` needs "Trust access to the VBA project object
model", which must never be enabled automatically. This tool extracts the
decompressed module source straight from the package with no Excel, no
Trust-Center change and no third-party package.

What it proves
--------------
Source-level equality of VBA components. It does NOT prove designer resources,
project properties, references, form layout or compiled behaviour, and it is
not a malware verdict.

Usage
-----
  compare_vba_source.py A.xlsm B.xlsm            # compare two packages
  compare_vba_source.py A.xlsm --list            # component inventory + sizes
  compare_vba_source.py A.xlsm --dump-dir DIR    # write one .txt per component
  compare_vba_source.py A.xlsm B.xlsm --out r.json

Exit codes: 0 identical, 1 differences found, 2 usage/format error.
"""
from __future__ import annotations

import argparse
import hashlib
import io
import json
import re
import struct
import sys
import zipfile
from pathlib import Path

from common import validate_new_output, write_new_json

# --------------------------------------------------------------------------
# OLE2 / compound file reader (MS-CFB), minimal read-only subset
# --------------------------------------------------------------------------
MAX_PACKAGE = 512 * 1024 * 1024
_ENDOFCHAIN, _FREESECT, _FATSECT, _DIFSECT = (
    0xFFFFFFFE, 0xFFFFFFFF, 0xFFFFFFFD, 0xFFFFFFFC)


class CompoundFile:
    """Read-only MS-CFB reader sufficient for vbaProject.bin."""

    def __init__(self, data: bytes):
        if len(data) < 512:
            raise ValueError("OLE header too short")
        if data[:8] != bytes.fromhex("D0CF11E0A1B11AE1"):
            raise ValueError("Not an OLE2 compound file")
        self.d = data
        ver_minor, ver_major, order = struct.unpack_from("<HHH", data, 24)
        if order != 0xFFFE:
            raise ValueError("Big-endian compound files are unsupported")
        if ver_major not in (3, 4):
            raise ValueError("Unsupported compound file version %d" % ver_major)
        self.sshift = struct.unpack_from("<H", data, 30)[0]
        self.mshift = struct.unpack_from("<H", data, 32)[0]
        if self.sshift not in (9, 12) or self.mshift not in (6,):
            raise ValueError("Unsupported sector sizes (shift %d/%d)"
                             % (self.sshift, self.mshift))
        self.ssize = 1 << self.sshift
        self.msize = 1 << self.mshift
        n_fat = struct.unpack_from("<I", data, 44)[0]
        self.dirstart = struct.unpack_from("<I", data, 48)[0]
        self.cutoff = struct.unpack_from("<I", data, 56)[0]
        self.minifatstart = struct.unpack_from("<I", data, 60)[0]
        self.n_minifat = struct.unpack_from("<I", data, 64)[0]
        self.difatstart = struct.unpack_from("<I", data, 68)[0]
        self.n_difat = struct.unpack_from("<I", data, 72)[0]
        if len(data) > MAX_PACKAGE:
            raise ValueError("Package exceeds inspection limit")
        self._fat = self._read_fat(n_fat, data)
        self._dir = self._read_dir()
        mf_raw = b"".join(self._sector(s) for s in self._chain_sectors(
            self.minifatstart, self.n_minifat)) if self.n_minifat else b""
        self._minifat = (list(struct.unpack_from("<%dI" % (len(mf_raw) // 4),
                                                 mf_raw, 0)) if mf_raw
                         else [_ENDOFCHAIN] * 8)
        self._root_start = self._dir[0][5]  # root storage start sector
        self._mini = self._chain_bytes(self._root_start)

    # -- low level -------------------------------------------------------
    def _sector(self, sid: int) -> bytes:
        off = (sid + 1) * self.ssize
        chunk = self.d[off:off + self.ssize]
        if len(chunk) != self.ssize:
            raise ValueError("Truncated sector %d" % sid)
        return chunk

    def _read_fat(self, n_fat: int, data: bytes) -> list[int]:
        fat_sids = [struct.unpack_from("<I", data, 76 + 4 * i)[0]
                    for i in range(109)]
        fat_sids = [s for s in fat_sids if s not in (_FREESECT, _FATSECT,
                                                     _ENDOFCHAIN, _DIFSECT)]
        sid, guard = self.difatstart, 0
        while sid not in (_ENDOFCHAIN, _FREESECT) and guard <= self.n_difat + 16:
            sec = self._sector(sid)
            entries = len(sec) // 4 - 1
            for i in range(entries):
                v = struct.unpack_from("<I", sec, 4 * i)[0]
                if v not in (_FREESECT, _FATSECT, _ENDOFCHAIN, _DIFSECT):
                    fat_sids.append(v)
            sid = struct.unpack_from("<I", sec, 4 * entries)[0]
            guard += 1
        fat: list[int] = []
        for s in fat_sids[:max(n_fat, 1)]:
            sec = self._sector(s)
            fat.extend(struct.unpack_from("<%dI" % (len(sec) // 4), sec, 0))
        if not fat:
            raise ValueError("Empty FAT")
        return fat

    def _chain_sectors(self, start: int, limit: int,
                       table: list[int] | None = None) -> list[int]:
        """Follow an allocation chain in the FAT (default) or the mini-FAT."""
        tab = self._fat if table is None else table
        out, sid, seen = [], start, set()
        while sid not in (_ENDOFCHAIN, _FREESECT):
            if sid >= len(tab):
                raise ValueError("Allocation chain out of range")
            if sid in seen:
                raise ValueError("Cycle in allocation chain")
            seen.add(sid)
            out.append(sid)
            sid = tab[sid]
            if len(out) > len(tab):
                raise ValueError("Allocation chain exceeds FAT")
        if limit and len(out) > limit:
            raise ValueError("Allocation chain longer than declared limit")
        return out

    def _chain_bytes(self, start: int) -> bytes:
        return b"".join(self._sector(s) for s in self._chain_sectors(start, 0))

    def _read_dir(self) -> list[tuple]:
        raw = self._chain_bytes(self.dirstart)
        entries = []
        for off in range(0, len(raw) - 127, 128):
            e = raw[off:off + 128]
            nlen = struct.unpack_from("<H", e, 64)[0]
            name = ""
            if 2 < nlen <= 64:
                try:
                    name = e[:nlen - 2].decode("utf-16-le")
                except UnicodeDecodeError:
                    name = ""
            entries.append((name, e[66], struct.unpack_from("<i", e, 68)[0],
                            struct.unpack_from("<i", e, 72)[0],
                            struct.unpack_from("<i", e, 76)[0],
                            struct.unpack_from("<I", e, 116)[0],
                            struct.unpack_from("<Q", e, 120)[0]))
        if not entries:
            raise ValueError("Empty directory")
        return entries

    # -- public ----------------------------------------------------------
    def list_streams(self, storage: str | None = "VBA") -> dict[str, bytes]:
        """Return {path: bytes} for streams below a storage named `storage`.

        ``storage=None`` returns every stream in the file. Directory children
        form a red-black tree: the left and right fields are both siblings in
        the ordered set, so the whole subtree must be walked, not just one
        sibling chain. Paths include the root storage name, e.g.
        ``Project/VBA/Module1``.
        """
        want = storage.casefold() if storage else None
        root = self._dir[0]
        root_name = root[0] or "RootEntry"
        out: dict[str, bytes] = {}
        seen: set[int] = set()

        def walk(node: int, path: str) -> None:
            if node < 0:
                return
            if node in seen:
                raise ValueError("Cycle in directory tree")
            if len(seen) > 200000:
                raise ValueError("Directory traversal guard tripped")
            seen.add(node)
            if node >= len(self._dir):
                raise ValueError("Directory index out of range")
            name, etype, left, right, kid, start, size = self._dir[node]
            full = (path + "/" + name) if name else path
            if etype == 2 and size:
                if want is None or want in [p.casefold() for p in
                                            full.split("/")[:-1]]:
                    out[full] = self._stream_bytes(start, size)
            elif etype == 1:
                walk(kid, full)
            walk(left, path)
            walk(right, path)

        walk(root[4], root_name)
        return out

    def _stream_bytes(self, start: int, size: int) -> bytes:
        if size < self.cutoff:
            out = b"".join(self._mini[s * self.msize:(s + 1) * self.msize]
                           for s in self._chain_sectors(start, 0, self._minifat))
        else:
            out = self._chain_bytes(start)
        if len(out) < size:
            raise ValueError("OLE stream shorter than its declared size")
        return out[:size]


# --------------------------------------------------------------------------
# MS-OVBA 2.4.1 decompression
# --------------------------------------------------------------------------
def _inflate_chunk(chunk: bytes) -> bytes:
    """Decompress one CompressedChunk (MS-OVBA 2.4.1.3.2).

    BitCount is derived from DecompressedCurrent *within this chunk*
    (MS-OVBA 2.4.1.3.16), so the offset grows as the chunk inflates and
    restarts at 4 bits for every chunk.
    """
    out = bytearray()
    i = 0
    while i < len(chunk):
        flag = chunk[i]
        i += 1
        for bit in range(8):
            if i >= len(chunk):
                return bytes(out)
            if flag & (1 << bit):
                if i + 1 >= len(chunk):
                    raise ValueError("Truncated copy token")
                if not out:
                    raise ValueError("Copy token with empty decompressed buffer")
                bit_size = max(4, (len(out) - 1).bit_length())
                length_mask = 0xFFFF >> bit_size
                pos_mask = 0xFFFF ^ length_mask
                tok = chunk[i] | (chunk[i + 1] << 8)
                i += 2
                length = (tok & length_mask) + 3
                pos = ((tok & pos_mask) >> (16 - bit_size)) + 1
                start = len(out) - pos
                if start < 0:
                    raise ValueError("Copy token offset before chunk start")
                if len(out) + length > 4096:
                    raise ValueError("Decompressed chunk exceeds 4096 bytes")
                for k in range(length):
                    out.append(out[start + k])
            else:
                if len(out) == 4096:
                    raise ValueError("Decompressed chunk exceeds 4096 bytes")
                out.append(chunk[i])
                i += 1
    return bytes(out)


def vba_decompress(data: bytes) -> bytes:
    """Decompress an MS-OVBA 2.4.1 CompressedContainer.

    A stream that does not start with the 0x01 signature byte is returned
    unchanged (uncompressed PerformanceCache/text).
    """
    if not data:
        return b""
    if data[0] != 0x01:
        return data
    out = bytearray()
    i = 1
    while i + 2 <= len(data):
        header = data[i] | (data[i + 1] << 8)
        i += 2
        size = (header & 0x0FFF) + 3
        if (header >> 12) & 0x07 != 0b011:
            raise ValueError("Invalid CompressedChunkSignature")
        compressed = bool((header >> 15) & 0x01)
        end = i + size - 2
        if end > len(data):
            raise ValueError("Truncated CompressedChunk payload")
        chunk = data[i:end]
        i = end
        if compressed:
            out += _inflate_chunk(chunk)
        else:
            if len(chunk) > 4096:
                raise ValueError("Uncompressed chunk exceeds 4096 bytes")
            out += chunk
    if i != len(data):
        raise ValueError("Truncated trailing CompressedChunk header")
    return bytes(out)


# --------------------------------------------------------------------------
# Component extraction
# --------------------------------------------------------------------------
_NAME_RE = re.compile(r'^\s*(?:Attribute\s+VB_Name\s*=\s*)?"([^"]+)"',
                      re.IGNORECASE | re.MULTILINE)


_VBNAME_RE = re.compile(rb'(?i)Attribute\s+VB_Name')


def looks_like_vba(text: bytes) -> bool:
    """True when the text carries the component-name attribute.

    Modules start with ``Attribute VB_Name``, forms and classes start with a
    ``VERSION``/``BEGIN`` header that contains it, so this marker covers every
    component kind instead of matching only one line shape.
    """
    return bool(_VBNAME_RE.search(text[:8192]))


def find_source_start(raw: bytes) -> int:
    """Offset where the module's CompressedContainer begins.

    A module stream is PerformanceCache followed by the compressed source at
    TextOffset (MS-OVBA 2.3.4.2.2.1 C), so decompressing from byte 0 yields
    garbage whenever the cache is non-empty. Candidates are located by the
    container signature plus a valid chunk signature, and are accepted only
    when the decoded text really starts with VBA source.
    """
    if not raw:
        return 0
    def probe(off: int) -> bool:
        """True when a container at `off` inflates to plausible VBA source.

        Only the first chunk is inflated, so scanning stays linear.
        """
        blob = raw[off:]
        if not blob or blob[0] != 0x01:
            return False
        try:
            if len(blob) > 3:
                header = blob[1] | (blob[2] << 8)
                if (header >> 12) & 0x07 != 0b011:
                    return False
                size = (header & 0x0FFF) + 3
                blob = blob[:min(len(blob), size + 1)]
            return looks_like_vba(vba_decompress(blob)[:8192])
        except ValueError:
            return False

    if probe(0):
        return 0
    for i in range(1, len(raw) - 3):
        if raw[i] == 0x01 and probe(i):
            return i
    raise ValueError("No VBA source container found in module stream")


def dir_code_page(dir_stream: bytes | None) -> str | None:
    """Charset declared by the DIR stream CODEPAGE record (tag 0x0003).

    VBA stores module source in the *project* code page. Guessing it is not
    safe: byte 0x9C is ``œ`` in cp1252 but ``ś`` in cp1250, so a Central
    European project decoded as cp1252 silently corrupts every accented word
    while still looking like valid text.
    """
    if not dir_stream:
        return None
    i = 0
    while i + 6 <= len(dir_stream):
        tag, size = struct.unpack_from("<HI", dir_stream, i)
        body = i + 6
        if size > 0x10000 or body + size > len(dir_stream):
            return None
        if tag == 0x0003:  # CODEPAGE
            if size < 2:
                return None
            cp = struct.unpack_from("<H", dir_stream, body)[0]
            if cp == 0:
                return "utf-8"  # 0 means Unicode storage
            try:
                "".encode("cp%d" % cp)
                return "cp%d" % cp
            except LookupError:
                return None
        if tag == 0x0002 and size >= 4:  # VERSION: stop after project records
            pass
        i = body + size
    return None


def _decode(raw: bytes, enc: str | None) -> str:
    """Decode strictly using the declared project code page, never guess.

    Replacement decoding collapses distinct byte sequences to the same U+FFFD;
    this can generate a false IDENTICAL result in an integrity comparison.
    """
    if enc is None:
        raise ValueError("VBA project code page unavailable: comparison cannot "
                         "safely decode component source")
    try:
        return raw.decode(enc, "strict")
    except UnicodeDecodeError as exc:
        raise ValueError("Invalid byte sequence in declared VBA code page "
                         + enc) from exc


def extract_vba_source(package: Path) -> dict[str, str]:
    """{component_name: source_text} for every VBA module in the package."""
    if package.is_symlink():
        raise ValueError("Symlink inputs are not followed")
    if not zipfile.is_zipfile(package):
        raise ValueError("Not an OOXML package (legacy .xls or encrypted "
                         "files need a different approved reader)")
    size = package.stat().st_size
    if size > MAX_PACKAGE:
        raise ValueError("Package exceeds inspection limit")
    with zipfile.ZipFile(package) as z:
        names = [n for n in z.namelist() if n.lower().endswith("vbaproject.bin")]
        if not names:
            raise ValueError("No vbaProject.bin part: macro-free package")
        if len(names) > 1:
            raise ValueError("Multiple vbaProject.bin parts: ambiguous package")
        if z.getinfo(names[0]).file_size > MAX_PACKAGE:
            raise ValueError("Expanded VBA project exceeds inspection limit")
        blob = z.read(names[0])
    cf = CompoundFile(blob)
    streams = cf.list_streams("VBA")
    if not streams:
        raise ValueError("No VBA streams found in vbaProject.bin")
    by_leaf = {full.rsplit("/", 1)[-1]: raw for full, raw in streams.items()}
    enc: str | None = None
    if "dir" not in by_leaf:
        raise ValueError("Missing VBA project DIR stream and declared code page")
    enc = dir_code_page(vba_decompress(by_leaf["dir"]))
    if enc is None:
        raise ValueError("Missing or unsupported VBA project code page in DIR")
    out: dict[str, str] = {}
    for leaf, raw in by_leaf.items():
        if leaf.lower() in ("_vba_project", "dir", "project", "projectwm",
                            "sections") or leaf.lower().startswith("__srp_"):
            # _VBA_PROJECT = performance cache, dir = project records,
            # __SRP_* = auto-recovery records: none of them is module source.
            continue
        text = _decode(vba_decompress(raw[find_source_start(raw):]), enc)
        m = _NAME_RE.search(text[:4096])
        name = m.group(1) if m else leaf
        if name in out:
            raise ValueError("Duplicate component name: " + name)
        out[name] = text
    if not out:
        raise ValueError("VBA streams present but no component source extracted")
    return out


def compare(a: Path, b: Path) -> dict:
    sa, sb = extract_vba_source(a), extract_vba_source(b)
    # Non-vacuity: an empty read is a tool failure, never "identical".
    lines_a = sum(len(v.splitlines()) for v in sa.values())
    lines_b = sum(len(v.splitlines()) for v in sb.values())
    if not sa or not sb or lines_a == 0 or lines_b == 0:
        raise ValueError("Extracted %d/%d source lines - refusing to report "
                         "equality on an empty read" % (lines_a, lines_b))
    keys_a, keys_b = set(sa), set(sb)
    changed = sorted(k for k in keys_a & keys_b if sa[k] != sb[k])
    detail = []
    for k in changed:
        la, lb = sa[k].splitlines(), sb[k].splitlines()
        diff = [i + 1 for i, (x, y) in enumerate(zip(la, lb)) if x != y]
        detail.append({"component": k, "lines_a": len(la), "lines_b": len(lb),
                       "first_differing_lines": diff[:10],
                       "differing_line_count": len(diff) + abs(len(la) - len(lb))})
    return {
        "status": "IDENTICAL" if not changed and keys_a == keys_b else "DIFFERS",
        "a": str(a), "b": str(b),
        "components_a": len(keys_a), "components_b": len(keys_b),
        "source_lines_a": lines_a, "source_lines_b": lines_b,
        "only_in_a": sorted(keys_a - keys_b), "only_in_b": sorted(keys_b - keys_a),
        "changed_components": changed, "changed_detail": detail,
        "scope": ("Decompressed VBA component source only. Not designer "
                  "resources, project properties, references, form layout, "
                  "compiled behaviour or a malware verdict."),
    }


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("a", type=Path)
    p.add_argument("b", type=Path, nargs="?")
    p.add_argument("--list", action="store_true", help="component inventory")
    p.add_argument("--dump-dir", type=Path, help="write one file per component")
    p.add_argument("--out", type=Path, help="write JSON report to a NEW path")
    a = p.parse_args()
    try:
        for f in (a.a, a.b):
            if f and not f.is_file():
                raise ValueError("File not found: %s" % f)
        validate_new_output(a.out)
        if a.list or a.dump_dir:
            src = extract_vba_source(a.a)
            if a.dump_dir:
                if a.dump_dir.exists() and any(a.dump_dir.iterdir()):
                    raise ValueError("--dump-dir must be absent or empty; "
                                     "refusing to mix output with other files")
                a.dump_dir.mkdir(parents=True, exist_ok=True)
                for name, text in sorted(src.items()):
                    safe = re.sub(r'[^\w.\-]', "_", name)[:80] or "unnamed"
                    (a.dump_dir / (safe + ".txt")).write_text(text, encoding="utf-8")
            report = {"status": "INVENTORY", "components": len(src),
                      "source_lines": sum(len(v.splitlines()) for v in src.values()),
                      "inventory": {k: len(v.splitlines()) for k, v in sorted(src.items())}}
        elif a.b:
            report = compare(a.a, a.b)
        else:
            raise ValueError("Provide a second package, or --list/--dump-dir")
        report["tool"] = "compare_vba_source.py"
        text = json.dumps(report, indent=2, ensure_ascii=False)
        if a.out:
            write_new_json(a.out, report)
        print(text)
        return 0 if report["status"] in ("IDENTICAL", "INVENTORY") else 1
    except (OSError, ValueError, zipfile.BadZipFile, struct.error) as exc:
        print("ERROR: %s" % exc, file=sys.stderr)
        return 2


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    raise SystemExit(main())
