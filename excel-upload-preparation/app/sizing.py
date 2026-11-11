"""Container-level  facts about  an uploaded workbook, read from the zip (OPC)
package alone -- no openpyxl, no Excel, no cell is parsed:

  * the size on disk and the *decompressed* size (what the scanner really
    has to chew  through -- an .xlsb is small on disk because its parts  are
    binary AND deflated; the unpacked size  is the honest measure),
  * every sheet, in workbook  order, with the decompressed size of its part,
    so a user asked  to skip sheets can see which ones  are heavy.

Sheet names come from ``xl/workbook.xml`` for .xlsx/.xlsm and from the
BIFF12 ``xl/workbook.bin`` (BrtBundleSh records, MS-XLSB 2.4.304) for
.xlsb; each is joined to its part through the workbook rels. This is the
information behind the upload size gate (config ``upload_size_threshold_mb``,
env ``MIND_READY_SIZE_THRESHOLD_MB``): above the threshold the app asks
whether some sheets should be ignored before it scans anything.
"""
from __future__ import annotations

import os
import struct
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Any

DEFAULT_THRESHOLD_MB = 25.0
MB = 1024 * 1024
_STATES = {0: "visible", 1: "hidden", 2: "veryHidden"}
_BRT_BUNDLE_SH = 0x009C
_MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_PKG_REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def size_threshold_bytes(config: dict[str, Any] | None = None) -> int:
    """Decompressed-size  threshold above  which the app asks about skipping
    sheets. Environment override first, then config/default.yaml, then the
    built-in default."""
    raw: Any = os.environ.get("MIND_READY_SIZE_THRESHOLD_MB")
    if raw in (None, ""):
        raw = (config or {}).get("upload_size_threshold_mb", DEFAULT_THRESHOLD_MB)
    try:
        mb = float(raw)
    except (TypeError, ValueError):
        mb = DEFAULT_THRESHOLD_MB
    if mb <= 0:
        mb = DEFAULT_THRESHOLD_MB
    return int(mb * MB)


# --- workbook part: sheet list ------------------------------------------------------------
def _rels(zf: zipfile.ZipFile, rels_name: str) -> dict[str, str]:
    """rId  -> zip  member name (targets are relative to xl/ unless absolute)."""
    out: dict[str, str] = {}
    if rels_name not in zf.namelist():
        return out
    try:
        root = ET.fromstring(zf.read(rels_name))
    except ET.ParseError:
        return out
    for rel in root.iter(f"{_PKG_REL_NS}Relationship"):
        rid, target = rel.get("Id"), rel.get("Target") or ""
        if not rid or not target:
            continue
        target = target.lstrip("/") if target.startswith("/") else f"xl/{target}"
        out[rid] = target.replace("\\", "/")
    return out


def _sheets_from_xml(zf: zipfile.ZipFile) -> list[dict[str, Any]]:
    root = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = _rels(zf, "xl/_rels/workbook.xml.rels")
    sheets = []
    for i, el in enumerate(root.iter(f"{_MAIN_NS}sheet")):
        rid = el.get(f"{_REL_NS}id") or ""
        sheets.append({"name": el.get("name") or f"Sheet{i + 1}", "state": el.get("state") or "visible", "index": i, "part": rels.get(rid)})
    return sheets


def _read_varint(data: bytes, pos: int, max_bytes: int) -> tuple[int, int]:
    """BIFF12  record type  (up to 2 bytes) / record size (up to 4 bytes):
    7 data bits per byte, high bit = another byte follows."""
    value, shift = 0, 0
    for n in range(max_bytes):
        if pos >= len(data):
            raise ValueError("truncated record header")
        b = data[pos]
        pos += 1
        value |= (b & 0x7F) << shift
        shift += 7
        if not b & 0x80:
            break
    return value, pos


def _wide_string(payload: bytes, pos: int) -> tuple[str | None, int]:
    """XLWideString  / XLNullableWideString:  4-byte char count (0xFFFFFFFF =
    null), then UTF-16LE characters."""
    (cch,) = struct.unpack_from("<I", payload, pos)
    pos += 4
    if cch == 0xFFFFFFFF:
        return None, pos
    text = payload[pos : pos + 2 * cch].decode("utf-16-le", errors="replace")
    return text, pos + 2 * cch


def _sheets_from_bin(zf: zipfile.ZipFile) -> list[dict[str, Any]]:
    """BrtBundleSh  records of  xl/workbook.bin: hsState (4), iTabID (4),
    strRelID (nullable wide string), strName (wide string)."""
    data = zf.read("xl/workbook.bin")
    rels = _rels(zf, "xl/_rels/workbook.bin.rels")
    sheets: list[dict[str, Any]] = []
    pos = 0
    while pos < len(data):
        rtype, pos = _read_varint(data, pos, 2)
        size, pos = _read_varint(data, pos, 4)
        payload = data[pos : pos + size]
        pos += size
        if rtype != _BRT_BUNDLE_SH:
            continue
        try:
            (hs_state, _tab_id) = struct.unpack_from("<II", payload, 0)
            rid, p = _wide_string(payload, 8)
            name, _ = _wide_string(payload, p)
        except (struct.error, IndexError):
            continue
        sheets.append({"name": name or f"Sheet{len(sheets) + 1}", "state": _STATES.get(hs_state, "visible"), "index": len(sheets), "part": rels.get(rid or "")})
    return sheets


def inspect_container(path: str | Path) -> dict[str, Any]:
    """Sizes  and the  sheet list of a workbook package, from the zip alone."""
    path = Path(path)
    fmt = path.suffix.lower().lstrip(".")
    file_bytes = path.stat().st_size if path.is_file() else 0
    info: dict[str, Any] = {
        "format": fmt,
        "file_bytes": file_bytes,
        "decompressed_bytes": file_bytes,
        "sheet_bytes": 0,
        "sheets": [],
        "largest_parts": [],
        "sheet_list_source": None,
        "warnings": [],
    }
    try:
        zf = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError) as exc:
        info["warnings"].append(f"not a zip package ({exc}); the size on disk is used")
        return info
    with zf:
        infos = zf.infolist()
        sizes = {i.filename: i.file_size for i in infos}
        compressed = {i.filename: i.compress_size for i in infos}
        info["decompressed_bytes"] = sum(sizes.values())
        info["largest_parts"] = [{"part": n, "bytes": b} for n, b in sorted(sizes.items(), key=lambda kv: -kv[1])[:6]]
        names = set(sizes)
        sheets: list[dict[str, Any]] = []
        try:
            if "xl/workbook.xml" in names:
                sheets, info["sheet_list_source"] = _sheets_from_xml(zf), "xl/workbook.xml"
            elif "xl/workbook.bin" in names:
                sheets, info["sheet_list_source"] = _sheets_from_bin(zf), "xl/workbook.bin"
            else:
                info["warnings"].append("no workbook part found; sheets are unknown until the file is opened")
        except Exception as exc:  # a malformed part must not block the upload
            info["warnings"].append(f"could not read the sheet list ({exc})")
            sheets = []
        for sh in sheets:
            part = sh.get("part")
            sh["bytes"] = sizes.get(part, 0) if part else 0
            sh["compressed_bytes"] = compressed.get(part, 0) if part else 0
        info["sheets"] = sheets
        # every worksheet part, whether or not the workbook part named it
        info["sheet_bytes"] = sum(b for n, b in sizes.items() if n.startswith("xl/worksheets/sheet"))
        info["model_bytes"] = sum(b for n, b in sizes.items() if n.startswith("xl/model/"))
    total_sheet = sum(sh["bytes"] for sh in info["sheets"]) or 1
    for sh in info["sheets"]:
        sh["share"] = round(sh["bytes"] / total_sheet, 4)
    return info


def size_gate(path: str | Path, config: dict[str, Any] | None = None) -> dict[str, Any]:
    """The  upload gate:  container facts plus the threshold verdict. The
    measure is always the decompressed size -- for an .xlsb that is the only
    honest one, and it is the same yardstick for every format."""
    info = inspect_container(path)
    threshold = size_threshold_bytes(config)
    above = info["decompressed_bytes"] > threshold
    info.update(
        {
            "measure": "decompressed",
            "threshold_bytes": threshold,
            "threshold_mb": round(threshold / MB, 1),
            "file_mb": round(info["file_bytes"] / MB, 1),
            "decompressed_mb": round(info["decompressed_bytes"] / MB, 1),
            "above_threshold": above,
            "message": (
                f"This workbook unpacks to {info['decompressed_bytes'] / MB:.1f} MB (limit {threshold / MB:.0f} MB), so the scan can take a while. "
                "You can skip sheets the app does not need to look at."
                if above
                else f"{info['decompressed_bytes'] / MB:.1f} MB unpacked -- under the {threshold / MB:.0f} MB limit."
            ),
        }
    )
    return info
