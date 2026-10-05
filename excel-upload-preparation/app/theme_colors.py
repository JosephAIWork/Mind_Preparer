"""FMT-002 where the colours live: the workbook's style table (1.8.0).

A cell does not carry a colour. It points at a font and a fill in
`xl/styles.xml`, and a few hundred fonts and fills serve hundreds of
thousands of cells. Until now the action "Replace theme colours with
explicit RGB" went through the cells in Excel (read the colour, assign it
back), which had two faults on real models:

  * it never finished its job: when the colour assigned is the one the cell
    already shows, Excel keeps the theme reference for the default text
    colour, so the same cells were planned again in every Prep round
    (11,807 -> 11,719 -> 11,720 cells on one sheet of Palermo);
  * it was slow where colours are mixed: about 10 COM calls a cell -- an hour
    for one 123,686-cell sheet of CNHI, with 60 more sheets behind it.

Here every theme colour of the style table's fonts and fills is replaced by
the RGB colour Excel shows for it (theme colour + tint, in Excel's own
arithmetic: see `tinted`). Same look, nothing depends on the theme any more,
every sheet at once, in well under a minute. Only `xl/styles.xml` changes inside
the package; `prep.apply_operations` does this on its fresh copy *before*
Excel opens it, so the file the user gets is still one Excel wrote.

Not touched: borders, conditional formats (`dxfs`) and rich-text runs --
FMT-002 looks at a cell's font colour and fill only.
"""
from __future__ import annotations

import os
import re
import zipfile
from pathlib import Path
from typing import Any

# SpreadsheetML numbers the theme's colours with the first two pairs swapped:
# theme="0" is the light background (lt1), theme="1" the dark text (dk1).
THEME_ORDER = ("lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink")
SYSTEM_COLORS = {"windowtext": "000000", "window": "FFFFFF"}
STYLES_PART = "xl/styles.xml"
DEFAULT_THEME_PART = "xl/theme/theme1.xml"

_SECTION = re.compile(r"<(fonts|fills)\b[^>]*>.*?</\1\s*>", re.S)
_COLOR_TAG = re.compile(r"<(color|fgColor|bgColor)\b([^>]*?)/>")
_ATTR = re.compile(r'([\w:]+)\s*=\s*"([^"]*)"')


def theme_palette(theme_xml: str) -> list[str] | None:
    """The 12 colours of the theme as RRGGBB, in SpreadsheetML's order.
    None when the scheme cannot be read (the caller then leaves the workbook alone)."""
    out = []
    for name in THEME_ORDER:
        m = re.search(rf"<(?:\w+:)?{name}\b[^>]*>(.*?)</(?:\w+:)?{name}\s*>", theme_xml, re.S)
        if not m:
            return None
        body = m.group(1)
        rgb = re.search(r'srgbClr\b[^>]*\bval="([0-9A-Fa-f]{6})"', body)
        if rgb:
            out.append(rgb.group(1).upper())
            continue
        system = re.search(r"<(?:\w+:)?sysClr\b([^>]*)>", body)
        if not system:
            return None
        attrs = dict(_ATTR.findall(system.group(1)))
        last = attrs.get("lastClr") or SYSTEM_COLORS.get(str(attrs.get("val", "")).lower())
        if not last or not re.fullmatch(r"[0-9A-Fa-f]{6}", last):
            return None
        out.append(last.upper())
    return out


HLSMAX = 240  # Windows' HLS scale, which is what Excel computes a tint on
RGBMAX = 255


def _to_hls(r: int, g: int, b: int) -> tuple[int, int, int]:
    high, low = max(r, g, b), min(r, g, b)
    lum = ((high + low) * HLSMAX + RGBMAX) // (2 * RGBMAX)
    if high == low:
        return HLSMAX * 2 // 3, lum, 0
    spread = high - low
    if lum <= HLSMAX // 2:
        sat = (spread * HLSMAX + (high + low) // 2) // (high + low)
    else:
        sat = (spread * HLSMAX + (2 * RGBMAX - high - low) // 2) // (2 * RGBMAX - high - low)
    rd, gd, bd = (((high - c) * (HLSMAX // 6) + spread // 2) // spread for c in (r, g, b))
    if r == high:
        hue = bd - gd
    elif g == high:
        hue = HLSMAX // 3 + rd - bd
    else:
        hue = 2 * HLSMAX // 3 + gd - rd
    if hue < 0:
        hue += HLSMAX
    if hue > HLSMAX:
        hue -= HLSMAX
    return hue, lum, sat


def _channel(n1: int, n2: int, hue: int) -> int:
    if hue < 0:
        hue += HLSMAX
    if hue > HLSMAX:
        hue -= HLSMAX
    if hue < HLSMAX // 6:
        return n1 + ((n2 - n1) * hue + HLSMAX // 12) // (HLSMAX // 6)
    if hue < HLSMAX // 2:
        return n2
    if hue < HLSMAX * 2 // 3:
        return n1 + ((n2 - n1) * (HLSMAX * 2 // 3 - hue) + HLSMAX // 12) // (HLSMAX // 6)
    return n1


def _to_rgb(hue: int, lum: int, sat: int) -> tuple[int, int, int]:
    if sat == 0:
        grey = (lum * RGBMAX + HLSMAX // 2) // HLSMAX
        return grey, grey, grey
    if lum <= HLSMAX // 2:
        n2 = (lum * (HLSMAX + sat) + HLSMAX // 2) // HLSMAX
    else:
        n2 = lum + sat - (lum * sat + HLSMAX // 2) // HLSMAX
    n1 = 2 * lum - n2
    r, g, b = (min(RGBMAX, (_channel(n1, n2, h) * RGBMAX + HLSMAX // 2) // HLSMAX) for h in (hue + HLSMAX // 3, hue, hue - HLSMAX // 3))
    return r, g, b


def tinted(rgb: str, tint: float) -> str:
    """A colour with Excel's tint applied (-1 = black ... 0 = unchanged ... 1 = white).

    The luminance moves, hue and saturation stay (ECMA-376 18.8.19) -- with
    Excel's own whole-number arithmetic: HLS on Windows' 0..240 scale, the
    darkened luminance cut to a whole number, the lightened one built from two
    parts cut separately. The textbook formula in floating point is 1 or 2
    off on more than half the shades; this one gave Excel's `Interior.Color`
    exactly for all 264 colour / tint pairs tried (12 theme colours, 22 tints;
    a sample is in tests/unit/test_theme_colors.py)."""
    if not tint:
        return rgb.upper()
    hue, lum, sat = _to_hls(*(int(rgb[i:i + 2], 16) for i in (0, 2, 4)))
    if tint < 0:
        lum = int(lum * (1.0 + tint))
    else:
        lum = int(lum * (1.0 - tint)) + (HLSMAX - int(HLSMAX * (1.0 - tint)))
    return "".join(f"{c:02X}" for c in _to_rgb(hue, min(HLSMAX, max(0, lum)), sat))


def explicit_styles(styles_xml: str, palette: list[str]) -> tuple[str, dict[str, int]]:
    """`xl/styles.xml` with every theme colour of its fonts and fills written
    as RGB. -> (the new text, {"fonts": n, "fills": n} colours converted).
    Everything else in the text is returned byte for byte."""
    counts = {"fonts": 0, "fills": 0}

    def section(found: re.Match) -> str:
        kind = found.group(1)

        def colour(tag: re.Match) -> str:
            attrs = dict(_ATTR.findall(tag.group(2)))
            if "theme" not in attrs:
                return tag.group(0)
            try:
                base = palette[int(attrs["theme"])]
                tint = float(attrs.get("tint") or 0.0)
            except (ValueError, IndexError):
                return tag.group(0)  # a theme index this palette does not have: left as it is
            counts[kind] += 1
            return f'<{tag.group(1)} rgb="FF{tinted(base, tint)}"/>'

        return _COLOR_TAG.sub(colour, found.group(0))

    return _SECTION.sub(section, styles_xml), counts


def _theme_part(names: set[str], read: Any) -> str | None:
    rels = "xl/_rels/workbook.xml.rels"
    if rels in names:
        for rel in re.findall(r"<Relationship\b[^>]*>", read(rels).decode("utf-8", errors="replace")):
            attrs = dict(_ATTR.findall(rel))
            if str(attrs.get("Type", "")).endswith("/theme") and attrs.get("Target"):
                target = attrs["Target"].lstrip("/")
                target = target if target.startswith("xl/") else "xl/" + target
                if target in names:
                    return target
    return DEFAULT_THEME_PART if DEFAULT_THEME_PART in names else None


def explicit_theme_colors_in_package(path: Path) -> dict[str, int] | None:
    """Rewrite `path` (an .xlsx / .xlsm package) with the theme colours of its
    fonts and fills as RGB. -> {"fonts": n, "fills": n}, or None when the
    package has no style table / no readable theme (nothing is written then).
    Only `xl/styles.xml` differs afterwards; the file is replaced in one move."""
    path = Path(path)
    if not zipfile.is_zipfile(path):
        return None
    with zipfile.ZipFile(path) as zin:
        names = set(zin.namelist())
        theme = _theme_part(names, zin.read)
        if STYLES_PART not in names or theme is None:
            return None
        palette = theme_palette(zin.read(theme).decode("utf-8", errors="replace"))
        if palette is None:
            return None
        new_styles, counts = explicit_styles(zin.read(STYLES_PART).decode("utf-8"), palette)
        if not (counts["fonts"] or counts["fills"]):
            return counts
        tmp = path.with_name(path.name + ".styles.tmp")
        try:
            # The copy is opened and saved again by Excel right after: speed matters here, not size.
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED, compresslevel=1) as zout:
                for info in zin.infolist():
                    data = new_styles.encode("utf-8") if info.filename == STYLES_PART else zin.read(info.filename)
                    fresh = zipfile.ZipInfo(info.filename, date_time=info.date_time)
                    fresh.compress_type = zipfile.ZIP_DEFLATED
                    fresh.external_attr = info.external_attr
                    zout.writestr(fresh, data, compresslevel=1)
        except Exception:
            tmp.unlink(missing_ok=True)
            raise
    os.replace(tmp, path)
    return counts
