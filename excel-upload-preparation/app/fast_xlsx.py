"""Faster openpyxl reading for big workbooks (1.11.4). Imported by app.inventory; installs itself once.

1. The XML parser. openpyxl reads every sheet with defusedxml's iterparse when defusedxml is installed,
   and that parser runs its handlers in Python: on Palermo (27 MB, 5.7 million used cells) half of the
   time a load took was spent inside it. The standard library's iterparse gives the same events from
   its C parser (expat), which never resolves external entities and, from expat 2.4, stops entity
   bombs itself -- what defusedxml guards against cannot happen there either.

2. The cached values. Several rules need the value Excel stored for each formula cell, and they got it
   by loading the whole workbook a second time with data_only=True -- another six minutes on Palermo.
   The value is in the same `<c>` element the formula is read from, so the first load keeps it
   (`ws.mind_ready_cached[(row, col)]`, exactly what a data_only load would give that cell) and
   `values_view` serves the data_only workbook from it. Cells without a formula are the same in both
   loads.

`MINDREADY_FAST_XLSX=0` turns both off (openpyxl exactly as shipped).
"""
from __future__ import annotations

import os
from typing import Any, Iterable

import openpyxl.reader.excel as _excel
import openpyxl.worksheet._reader as _reader
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.utils.datetime import from_excel, from_ISO8601

ENABLED = os.environ.get("MINDREADY_FAST_XLSX", "1").strip().lower() not in ("0", "off", "no", "false")
_INSTALLED = False


def _safe_iterparse():
    """The standard library's iterparse on its C parser (expat). Measured on Horizon: 5.4 s against
    7.4 s for defusedxml and 10.6 s for lxml (whose Python element proxies cost more than they save).
    ElementTree never fetches external entities or DTDs; expat >= 2.4 stops entity-expansion bombs
    (billion laughs) by itself. An older expat keeps defusedxml."""
    import pyexpat
    from xml.etree.ElementTree import iterparse

    version = tuple(int(x) for x in pyexpat.version_info[:2])
    if version < (2, 4):
        return None, None
    return iterparse, f"expat {pyexpat.EXPAT_VERSION.split('_')[-1]}"


class _CachedValueParser(_reader.WorkSheetParser):
    """openpyxl's sheet parser, which also keeps the cached value of every formula cell."""

    def __init__(self, *args: Any, **kwargs: Any) -> None:
        super().__init__(*args, **kwargs)
        self.cached: dict[tuple[int, int], Any] = {}

    def parse_cell(self, element):  # noqa: D401 -- openpyxl's signature
        out = super().parse_cell(element)
        if out["data_type"] == "f":
            self.cached[(out["row"], out["column"])] = self._cached_value(element, out["style_id"])
        return out

    def _cached_value(self, element, style_id):
        """What openpyxl's parse_cell returns as the value with data_only=True (same branches)."""
        data_type = element.get("t", "n")
        if data_type == "inlineStr":
            return None
        value = element.findtext(_reader.VALUE_TAG, None) or None
        if value is None:
            return None
        if data_type == "n":
            value = _reader._cast_number(value)
            if style_id in self.date_formats:
                try:
                    value = from_excel(value, self.epoch, timedelta=style_id in self.timedelta_formats)
                except (OverflowError, ValueError):
                    value = "#VALUE!"
        elif data_type == "s":
            value = self.shared_strings[int(value)]
        elif data_type == "b":
            value = bool(int(value))
        elif data_type == "d":
            value = from_ISO8601(value)
        return value


class _CachedValueReader(_reader.WorksheetReader):
    def __init__(self, ws, xml_source, shared_strings, data_only, rich_text):
        super().__init__(ws, xml_source, shared_strings, data_only, rich_text)
        if not data_only:
            self.parser = _CachedValueParser(xml_source, shared_strings, data_only, ws.parent.epoch, ws.parent._date_formats, ws.parent._timedelta_formats, rich_text)

    def bind_all(self):
        super().bind_all()
        if isinstance(self.parser, _CachedValueParser):
            self.ws.mind_ready_cached = self.parser.cached


def install() -> str | None:
    """Idempotent. -> the parser in use ('lxml' / 'expat'), None when turned off."""
    global _INSTALLED
    if not ENABLED:
        return None
    iterparse, name = _safe_iterparse()
    if not _INSTALLED:
        if iterparse is not None:
            _reader.iterparse = iterparse
        _excel.WorksheetReader = _CachedValueReader
        _INSTALLED = True
    return name or "defusedxml"


# --- the data_only view -------------------------------------------------------------------------
class _Value:
    __slots__ = ("value",)

    def __init__(self, value: Any) -> None:
        self.value = value


_EMPTY = _Value(None)


class _CellsView:
    """`ws._cells` of the data_only view: the formula cells answer with their cached value."""

    def __init__(self, cells: dict, cached: dict) -> None:
        self._cells = cells
        self._cached = cached

    def get(self, key, default=None):
        if key in self._cached:
            return _Value(self._cached[key])
        return self._cells.get(key, default)

    def __contains__(self, key) -> bool:
        return key in self._cells

    def __iter__(self):
        return iter(self._cells)

    def __len__(self) -> int:
        return len(self._cells)

    def keys(self):
        return self._cells.keys()

    def items(self):
        for key in self._cells:
            yield key, self.get(key)

    def __getitem__(self, key):
        c = self.get(key)
        if c is None:
            raise KeyError(key)
        return c


class _SheetView:
    def __init__(self, ws) -> None:
        self._ws = ws
        self.title = ws.title
        self._cells = _CellsView(getattr(ws, "_cells", {}), getattr(ws, "mind_ready_cached", {}))  # a chart sheet has no cells

    def cell(self, row: int, column: int):
        return self._cells.get((row, column)) or _EMPTY

    def __getitem__(self, coordinate: str):
        row, col = coordinate_to_tuple(coordinate.replace("$", ""))
        return self.cell(row, col)

    @property
    def max_row(self) -> int:
        return getattr(self._ws, "max_row", 1)

    @property
    def max_column(self) -> int:
        return getattr(self._ws, "max_column", 1)

    @property
    def sheet_state(self) -> str:
        return self._ws.sheet_state


class ValuesView:
    """A read-only stand-in for `openpyxl.load_workbook(path, data_only=True)` built from a workbook
    already loaded with formulas: `sheetnames`, `wb[sheet]`, `ws.cell(r, c).value`, `ws['A1'].value`,
    `ws._cells.get((r, c))`, `max_row` / `max_column`."""

    def __init__(self, wb) -> None:
        self._wb = wb
        self.sheetnames = list(wb.sheetnames)
        self._views: dict[str, _SheetView] = {}

    def __getitem__(self, name: str) -> _SheetView:
        view = self._views.get(name)
        if view is None:
            view = self._views[name] = _SheetView(self._wb[name])
        return view

    def __contains__(self, name: str) -> bool:
        return name in self.sheetnames

    @property
    def worksheets(self) -> list[_SheetView]:
        return [self[n] for n in self.sheetnames]

    def close(self) -> None:  # the formula workbook stays open in its own cache
        self._views.clear()


def values_view(wb, ignore: Iterable[str] = ()) -> ValuesView | None:
    """The data_only view of `wb` (a formula load), or None when the load did not keep the cached
    values (turned off, or loaded before this module was installed)."""
    sheets = [ws for ws in wb.worksheets if hasattr(ws, "_cells")]
    if not sheets or any(not hasattr(ws, "mind_ready_cached") and not getattr(ws, "mind_ready_ignored", False) for ws in sheets):
        return None
    return ValuesView(wb)
