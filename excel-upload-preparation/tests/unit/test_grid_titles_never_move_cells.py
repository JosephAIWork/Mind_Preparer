"""1.14.0: a grid title never moves a cell (Joseph, 10 November 2026).

The automatic namer inserted cells over one grid's columns to make room for its title; on the IFRS 17
tool's Result_CF that pushed the key column (A1:A8, read as a grid of its own) one row down while the
block beside it (B2:CI8) stayed, so every key sat beside the wrong row. Now: a key column with a
caption on top and the block that reads it are ONE table (grids._merge_key_columns); the title goes
into the caption cell ('#' + its text) or the empty cell above the top-left cell, never anywhere that
needs a row insert; a grid starting on row 1 and a taken cell above are reported under "could not
name" with what to do by hand. A formula caption (Result_CF's A1 on the IFRS 17 tool) is replaced by
the title since 1.14.1 (Joseph, 11 November 2026): the cell loses its formula and gets '#' + the text
the formula last showed (the workbook's cached value), the action's note says so and calls it an open
issue; an empty, error or non-text cached value (a workbook never calculated, a year) leaves the table
reported instead, and so does an INDIRECT or a 3-D reference on the sheet (the reader cannot be pinned
to a cell). The apply path still carries out an old plan's insert_row."""
from __future__ import annotations

import html
import re
import shutil
import zipfile
from pathlib import Path

import openpyxl
import pytest

from app import prep
from app.excel_com import com_available
from app.grids import all_grids, detect_grids
from app.inventory import build_analysis

needs_excel = pytest.mark.skipif(not com_available(), reason="requires pywin32 + an installed Excel")


def _wb(path: Path, sheets: dict[str, dict[str, object]]) -> Path:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, cells in sheets.items():
        ws = wb.create_sheet(name)
        for ref, v in cells.items():
            ws[ref] = v
    wb.save(path)
    wb.close()
    return path


def _analysis(tmp_path: Path, sheets: dict[str, dict[str, object]]) -> dict:
    return build_analysis(_wb(tmp_path / "m.xlsx", sheets), tmp_path / "w", "t")


def _titles(analysis, **kw):
    return prep.plan_create_grid_titles(analysis, {"findings": []}, **kw)


# Result_CF in miniature: a caption in A1, the header row 2, the keys of each row in column A (the
# sheet names the row's lookups use) and the block beside them that reads $A3.
RESULT_CF = {
    "A1": "Running scenario", "A2": "csv file", "B2": "ID", "C2": "Scenario",
    "A3": "BBA_01", "A4": "BBA_02", "A5": "VFA_01",
    "B3": '=$A3&"_1"', "B4": '=$A4&"_1"', "B5": '=$A5&"_1"',
    "C3": "=LEN($A3)", "C4": "=LEN($A4)", "C5": "=LEN($A5)",
}


def test_a_key_column_with_a_caption_on_top_and_its_block_are_one_table(tmp_path):
    analysis = _analysis(tmp_path, {"R": RESULT_CF})
    grids = all_grids(analysis)
    assert [g["ref"] for g in grids] == ["A2:C5"]
    table = grids[0]
    assert table["caption_cell"] == "A1" and table["caption_value"] == "Running scenario"
    assert table["merged_from"] == ["A1:A5", "B2:C5"]
    assert table["header_values"] == ["csv file", "ID", "Scenario"] and table["header_is_all_text"]
    assert table["name"] is None and table["anchor"] == "A2"


def test_the_table_is_named_once_into_its_caption_cell_and_nothing_moves(tmp_path):
    analysis = _analysis(tmp_path, {"R": RESULT_CF})
    ops, skipped = _titles(analysis)
    assert [o["op"] for o in ops] == ["set_value"]
    op = ops[0]
    assert op["sheet"] == "R" and op["cell"] == "A1" and op["before"] == "Running scenario" and op["after"] == "#Running scenario"
    assert "the caption becomes the title" in op["note"] and "key column A1:A5" in op["note"] and "block B2:C5" in op["note"]
    assert not skipped


def test_keys_that_are_formulas_merge_whether_or_not_the_block_still_reads_them(tmp_path):
    """Result_CF's keys are formulas (=Ctrl_01!D35) and the block reads them ($A3) -- until FRM-007
    rewrites INDIRECT($A3&"!A:B") as BBA_01!A:B, after which nothing reads column A any more. The
    table is the same table in both states (once the caption holds a title Mind reads one grid), so
    the merge looks at the shape alone: the keys may be values or formulas, read or not."""
    keys_as_formulas = {**RESULT_CF, "A3": "=Ctrl!A1", "A4": "=Ctrl!A2", "A5": "=Ctrl!A3"}
    analysis = _analysis(tmp_path, {"R": keys_as_formulas, "Ctrl": {"A1": "BBA_01", "A2": "BBA_02", "A3": "VFA_01"}})
    assert [g["ref"] for g in all_grids(analysis) if g["sheet"] == "R"] == ["A2:C5"]
    after_frm_007 = {**keys_as_formulas, "B3": '=IFERROR(VLOOKUP("k",Ctrl!A:B,2,FALSE),"-")', "B4": "=Ctrl!A2", "B5": "=Ctrl!A3", "C3": "=1", "C4": "=2", "C5": "=3"}
    analysis = _analysis(tmp_path, {"R": after_frm_007, "Ctrl": {"A1": "BBA_01", "A2": "BBA_02", "A3": "VFA_01"}})
    assert [g["ref"] for g in all_grids(analysis) if g["sheet"] == "R"] == ["A2:C5"]
    ops, skipped = _titles(analysis)
    assert [(o["sheet"], o["cell"], o["after"]) for o in ops if o["sheet"] == "R"] == [("R", "A1", "#Running scenario")]


def test_a_number_a_date_or_a_switch_on_top_of_a_key_column_is_not_a_caption(tmp_path):
    """A year, a valuation date or TRUE above a key column is a value something may read (a macro, a
    dropdown, a person): never merged as the caption, so no title is ever written over it."""
    import datetime

    for top in (2024, datetime.datetime(2024, 1, 1), True, 0.5):
        cells = {"A1": top, "A2": "k1", "A3": "k2", "A4": "k3", "B2": 1, "C2": "x", "B3": 2, "C3": "y", "B4": 3, "C4": "z"}
        analysis = _analysis(tmp_path, {"S": cells})
        grids = [g for g in all_grids(analysis) if g["sheet"] == "S"]
        assert [g["ref"] for g in grids] == ["A1:A4", "B2:C4"] and not any(g.get("caption_cell") for g in grids), top
        ops, skipped = _titles(analysis)
        assert not any(o["cell"] == "A1" for o in ops), (top, ops)
        assert any(s.startswith("S!A1:A4: starts on row 1") for s in skipped), (top, skipped)
    # a text caption merges, as before
    analysis = _analysis(tmp_path, {"S": {"A1": "Keys", "A2": "k1", "A3": "k2", "A4": "k3", "B2": 1, "C2": "x", "B3": 2, "C3": "y", "B4": 3, "C4": "z"}})
    assert [g["ref"] for g in all_grids(analysis) if g["sheet"] == "S"] == ["A2:C4"]


def test_a_label_beside_the_title_is_moved_only_when_the_title_is_made_from_it(tmp_path):
    """The Grid Namer writes the user's own name (and the automatic pass may use the assistant's):
    the label's text would not reach the title, so clearing it would just lose it. Then the label
    stays and the grid is reported; the automatic pass that names the grid FROM the label moves it."""
    cells = {"B2": "Mortality", "C3": "Age", "D3": "Qx", "C4": 40, "D4": 0.001}
    analysis = _analysis(tmp_path, {"L": cells})
    out = prep.plan_named_areas(analysis, [{"sheet": "L", "ref": "C3:D4", "name": "Qx table", "flags": []}])
    assert not out["ops"], out["ops"]
    assert any(s.startswith("L!C3:D4: a title in C2 would touch B2 ('Mortality') beside it, which is a label of its own") for s in out["skipped"]), out["skipped"]
    # the assistant's name instead of the label's weak text: the label stays too
    weak = {"B2": "x", "C3": "Age", "D3": "Qx", "C4": 40, "D4": 0.001}
    analysis = _analysis(tmp_path, {"L": weak})
    ops, skipped = _titles(analysis, names={"L!C3:D4": "Mortality by age"})
    assert not any(o["op"] == "clear_cell" for o in ops), ops
    assert not any(o["sheet"] == "L" and o["cell"] == "C2" for o in ops) and any(s.startswith("L!C3:D4: a title in C2 would touch B2") for s in skipped), (ops, skipped)
    # the automatic pass naming the grid from that very label moves it, as before
    ops, skipped = _titles(_analysis(tmp_path, {"L": cells}))
    assert [(o["op"], o["cell"], o["after"]) for o in ops] == [("set_value", "C2", "#Mortality"), ("clear_cell", "B2", None)]


def test_two_lists_that_only_touch_are_not_joined():
    # different lengths: the left list runs on below the block
    occupied = {(1, 1): "Names", (2, 1): "ann", (3, 1): "bob", (4, 1): "cy", (5, 1): "dee", (2, 2): "Ages", (3, 2): 30, (4, 2): 31}
    grids, _ = detect_grids("S", occupied)
    assert [g["ref"] for g in grids] == ["A1:A5", "B2:B4"] and not any(g.get("caption_cell") for g in grids)
    # the block starts two rows below the column's top (two cells on top, not one caption)
    occupied = {(1, 1): "Cap", (2, 1): "Sub", (3, 1): "k1", (4, 1): "k2", (3, 2): 1, (4, 2): 2}
    grids, _ = detect_grids("S", occupied)
    assert [g["ref"] for g in grids] == ["A1:A4", "B3:B4"]
    # the column is already titled: it is a grid of its own by its author's choice
    occupied = {(1, 1): "#Keys", (2, 1): "Cap", (3, 1): "k1", (4, 1): "k2", (3, 2): 1, (4, 2): 2}
    grids, _ = detect_grids("S", occupied)
    assert [(g["ref"], g["name"]) for g in grids] == [("A2:A4", "Keys"), ("B3:B4", None)]
    # the column is two cells wide
    occupied = {(1, 1): "Cap", (1, 2): "x", (2, 1): "k1", (2, 2): "y", (3, 1): "k2", (3, 2): "z", (2, 3): 1, (3, 3): 2}
    grids, _ = detect_grids("S", occupied)
    assert [g["ref"] for g in grids] == ["A1:B3", "C2:C3"]


def test_an_empty_cell_above_gets_the_made_up_title(tmp_path):
    analysis = _analysis(tmp_path, {"D": {"B3": "Age", "C3": "Rate", "B4": 30, "C4": 0.01, "B5": 31, "C5": 0.02}})
    ops, skipped = _titles(analysis)
    assert [(o["op"], o["cell"], o["before"], o["after"]) for o in ops] == [("set_value", "B2", None, "#Age Rate")]
    assert not skipped


CAPTION_FORMULA = '="Running scenario ["&E1&"]"'
FORMULA_CAPTION = {**RESULT_CF, "A1": CAPTION_FORMULA, "E1": "Base"}


def _plant_cached_value(path: Path, sheet: str, cell: str, text: str, kind: str = "str") -> Path:
    """What Excel saves for a formula cell and openpyxl does not: the value it last computed
    (`<c t="str"><f>..</f><v>text</v></c>`; kind "e" for an error such as #N/A)."""
    wb = openpyxl.load_workbook(path)
    member = f"xl/worksheets/sheet{wb.sheetnames.index(sheet) + 1}.xml"
    wb.close()
    tmp = path.with_suffix(".tmp")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == member:
                xml = data.decode("utf-8")
                patched = re.sub(rf'<c r="{cell}"><f>(.*?)</f><v></v></c>', lambda m: f'<c r="{cell}" t="{kind}"><f>{m.group(1)}</f><v>{html.escape(text)}</v></c>', xml, count=1)
                assert patched != xml, f"{sheet}!{cell} is not a formula cell without a cached value"
                data = patched.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, path)
    return path


def _analysis_with_cached(tmp_path: Path, sheets: dict[str, dict[str, object]], sheet: str, cell: str, text: str, kind: str = "str") -> dict:
    path = _plant_cached_value(_wb(tmp_path / "m.xlsx", sheets), sheet, cell, text, kind)
    return build_analysis(path, tmp_path / "w", "t")


def test_a_formula_caption_is_replaced_by_the_title_made_from_the_text_it_showed(tmp_path):
    """Joseph, 11 November 2026: Result_CF's caption formula goes; A1 gets '#' + the text Excel last
    computed for it. The note says so and calls it an open issue; nothing is inserted; the table is
    no longer under "could not name"."""
    analysis = _analysis_with_cached(tmp_path, {"R": FORMULA_CAPTION}, "R", "A1", "Running scenario [Base]")
    assert [g["ref"] for g in all_grids(analysis) if g["sheet"] == "R"] == ["A2:C5"]
    assert prep.values_cell(analysis, "R", 1, 1) == "Running scenario [Base]"  # the planted cached value is read
    ops, skipped = _titles(analysis)
    mine = [o for o in ops if o["sheet"] == "R"]
    assert [(o["op"], o["cell"], o["before"], o["after"]) for o in mine] == [("set_value", "A1", CAPTION_FORMULA, "#Running scenario [Base]")]
    note = mine[0]["note"]
    assert note.startswith("name from the caption above at A1; this cell held a formula (" + CAPTION_FORMULA + "); it is replaced by the title"), note
    assert "key column A1:A5" in note and "block B2:C5" in note
    assert "Open issue: if the caption must stay a formula, leave an empty row above the table instead" in note, note
    assert not any(o["op"] == "insert_row" for o in ops)
    assert not [s for s in skipped if s.startswith("R!A2:C5")], skipped
    # ... but a caption cell something reads through is still refused, formula or not
    analysis = _analysis_with_cached(tmp_path, {"R": FORMULA_CAPTION, "T": {"A1": "=R!A1"}}, "R", "A1", "Running scenario [Base]")
    ops, skipped = _titles(analysis)
    assert not [o for o in ops if o["sheet"] == "R"], ops
    assert any(s.startswith("R!A2:C5: the caption above it (A1) is read by T!A1") for s in skipped), skipped


def test_a_formula_caption_whose_last_value_is_empty_an_error_or_not_text_is_reported(tmp_path):
    """12 November 2026 review: a literal year above a key column is protected (not a caption), so a
    formula showing a year, an error or nothing must not lose its formula to a made-up name either."""
    reason = "R!A2:C5: the caption above it (A1) is a formula whose last computed value is not text (or the workbook holds no computed value for it) -- recalculate the workbook and name it again, or leave an empty row above the table by hand"
    # no cached value at all (a workbook never calculated)
    analysis = _analysis(tmp_path, {"R": FORMULA_CAPTION})
    assert prep.values_cell(analysis, "R", 1, 1) in (None, "")
    ops, skipped = _titles(analysis)
    assert not [o for o in ops if o["sheet"] == "R"], ops
    assert reason in skipped, skipped
    # the last value was an error
    analysis = _analysis_with_cached(tmp_path, {"R": FORMULA_CAPTION}, "R", "A1", "#N/A", kind="e")
    assert prep.values_cell(analysis, "R", 1, 1) == "#N/A"
    ops, skipped = _titles(analysis)
    assert not [o for o in ops if o["sheet"] == "R"] and reason in skipped, skipped
    # the last value was a number (the run year)
    analysis = _analysis_with_cached(tmp_path, {"R": FORMULA_CAPTION}, "R", "A1", "2026", kind="n")
    assert prep.values_cell(analysis, "R", 1, 1) == 2026
    ops, skipped = _titles(analysis)
    assert not [o for o in ops if o["sheet"] == "R"] and reason in skipped, skipped
    assert not any(o["op"] == "insert_row" for o in ops)
    # the Grid Namer refuses the same way
    out = prep.plan_named_areas(analysis, [{"sheet": "R", "ref": "B3", "name": "Result Cashflows", "flags": []}])
    assert out["ops"] == [] and reason in out["skipped"], out["skipped"]


def test_a_caption_on_a_sheet_an_indirect_or_a_3d_reference_reads_is_refused(tmp_path):
    """12 November 2026 review: `referenced_by` sees direct references only; a reader through an address
    built as text, or a 3-D reference across the sheet, may read the caption cell too, so the title
    would change its value. Named as the reader, as _move_to_free_column does."""
    analysis = _analysis_with_cached(tmp_path, {"R": FORMULA_CAPTION, "T": {"A1": '=INDIRECT("R!A"&1)'}}, "R", "A1", "Running scenario [Base]")
    assert "R" in prep.reference_index(analysis)["text_addressed"]
    ops, skipped = _titles(analysis)
    assert not [o for o in ops if o["sheet"] == "R"], ops
    assert any(s.startswith("R!A2:C5: the caption above it (A1) may be read by T!A1 through an address built as text (INDIRECT)") for s in skipped), skipped
    # a text caption on such a sheet is refused the same way
    analysis = _analysis(tmp_path, {"R": RESULT_CF, "T": {"A1": '=INDIRECT("R!A"&1)'}})
    ops, skipped = _titles(analysis)
    assert not [o for o in ops if o["sheet"] == "R"] and any("may be read by T!A1 through an address built as text" in s for s in skipped), skipped
    # a 3-D reference across the sheet
    analysis = _analysis(tmp_path, {"Q": {"A1": 1}, "R": RESULT_CF, "S": {"A1": 2}, "T": {"A1": "=SUM(Q:S!A1)"}})
    index = prep.reference_index(analysis)
    assert "R" in index["positional"], index["positional"]
    ops, skipped = _titles(analysis)
    assert not [o for o in ops if o["sheet"] == "R"] and any("may be read by the 3-D reference in T!A1" in s for s in skipped), skipped


def test_nothing_is_ever_inserted_for_a_title(tmp_path):
    analysis = _analysis(tmp_path, {
        "Top": {"A1": "X", "B1": "Y", "A2": 1, "B2": 2},  # starts on row 1
        "CF": {"B2": 1, "C2": "Demographic Assumptions", "C3": "Gender", "D3": "M", "C4": "Smoker", "D4": "NS"},  # a heading row (a grid) right above
        "Stack": {"B1": "H", "C1": "K", "B2": 1, "C2": 2, "B3": 3, "C3": 4, "B5": "H2", "C5": "K2", "B6": 5, "C6": 6},  # a grid right above the title cell, same first column
        "Beside": {"A1": "H", "B1": "K", "C1": "L", "A2": 1, "B2": 2, "C2": 3, "A3": 3, "B3": 4, "C3": 5, "B5": "H2", "C5": "K2", "B6": 5, "C6": 6},  # the grid above starts in another column
    })
    ops, skipped = _titles(analysis)
    assert not any(o["op"] == "insert_row" for o in ops)
    assert not any(o.get("after_inserts") or o.get("insert_at") or o.get("columns") for o in ops)
    by_sheet = {o["sheet"]: o for o in ops}
    assert "Top" not in by_sheet and any(s.startswith("Top!A1:B2: starts on row 1") and "Insert an empty row above it by hand" in s for s in skipped), skipped
    assert "CF" not in by_sheet or by_sheet["CF"]["cell"] != "C2"
    assert any(s.startswith("CF!C3:D4: the cell above (C2) holds 'Demographic Assumptions' (part of the grid untitled B2, B2:C2)") and "nothing is moved" in s for s in skipped), skipped
    assert any(s.startswith("Stack!B5:C6: a title in B4 would touch B3") and "one grid" in s for s in skipped), skipped
    assert by_sheet["Beside"]["cell"] == "B4"  # A1:C3 is read down column A: a title in B4 stands on its own


def test_a_label_the_plan_moves_into_the_title_may_sit_right_above_it(tmp_path):
    analysis = _analysis(tmp_path, {"D": {"B2": "Premium table", "B4": "Key", "C4": "Value", "B5": "a", "C5": 1}})
    ops, skipped = _titles(analysis)
    assert [(o["op"], o["cell"], o["after"]) for o in ops] == [("set_value", "B3", "#Premium table"), ("clear_cell", "B2", None)]
    # ... but not when something reads the label (it stays, and would swallow the title below it)
    analysis = _analysis(tmp_path, {"D": {"B2": "Premium table", "B4": "Key", "C4": "Value", "B5": "a", "C5": 1}, "T": {"A1": "=D!B2"}})
    ops, skipped = _titles(analysis)
    assert not ops and any(s.startswith("D!B4:C5: a title in B3 would touch B2") for s in skipped), skipped


def test_a_cell_beside_the_title_is_moved_into_it_or_reported(tmp_path):
    """Mind reads left to right: a cell right beside the title cell is read first and runs on into
    the title. The old namer wrote such titles (the anonymized IFRS 17 tool: 33 grids with a
    '#Title' trapped inside after its own re-scan). A lone label nothing reads is moved into the
    title; a label something reads, or a grid starting on that row, is reported."""
    analysis = _analysis(tmp_path, {
        "L": {"B2": "Mortality", "C3": "Age", "D3": "Qx", "C4": 40, "D4": 0.001},  # lone label beside the title cell C2
        "R": {"B2": "Mortality", "C3": "Age", "D3": "Qx", "C4": 40, "D4": 0.001, "F1": "=B2"},  # ... that a formula reads
        "G": {"A2": "k", "B2": 1, "D3": "Age", "E3": "Qx", "D4": 40, "E4": 0.001, "C2": "x"},  # a grid A2:C2 ends right beside D2
    })
    ops, skipped = _titles(analysis)
    by_sheet: dict[str, list] = {}
    for o in ops:
        by_sheet.setdefault(o["sheet"], []).append((o["op"], o["cell"], o["after"]))
    assert by_sheet["L"] == [("set_value", "C2", "#Mortality"), ("clear_cell", "B2", None)]
    assert "R" not in by_sheet and any(s.startswith("R!C3:D4: a title in C2 would touch B2 ('Mortality') beside it, which is read by R!F1") for s in skipped), skipped
    assert not any(o[1] == "D2" for o in by_sheet.get("G", [])) and any(s.startswith("G!D3:E4: a title in D2 would touch C2 ('x') beside it, which is part of the grid") for s in skipped), skipped


def test_the_grid_namer_follows_the_same_rule(tmp_path):
    analysis = _analysis(tmp_path, {"R": RESULT_CF, "Top": {"A1": "X", "B1": "Y", "A2": 1, "B2": 2}})
    out = prep.plan_named_areas(analysis, [
        {"sheet": "R", "ref": "B3:C4", "name": "Result Cashflows", "flags": ["Input"]},
        {"sheet": "Top", "ref": "A1:B2", "name": "Top Grid", "flags": []},
    ])
    assert [(o["op"], o["sheet"], o["cell"], o["before"], o["after"]) for o in out["ops"]] == [("set_value", "R", "A1", "Running scenario", "#Result Cashflows /Input")]
    assert out["named"] == ["R!A2:C5"] and out["inserted_rows"] == {} and out["banded_rows"] == {}
    assert any(s.startswith("Top!A1:B2: starts on row 1") for s in out["skipped"]), out["skipped"]
    # selecting the caption cell alone names the table too
    out = prep.plan_named_areas(analysis, [{"sheet": "R", "ref": "A1", "name": "Result Cashflows", "flags": []}])
    assert [o["cell"] for o in out["ops"]] == ["A1"]
    # a formula caption is replaced by the user's name too, with the same note (1.14.1)
    analysis = _analysis_with_cached(tmp_path, {"R": FORMULA_CAPTION}, "R", "A1", "Running scenario [Base]")
    out = prep.plan_named_areas(analysis, [{"sheet": "R", "ref": "B3", "name": "Result Cashflows", "flags": ["Input"]}])
    assert [(o["op"], o["cell"], o["before"], o["after"]) for o in out["ops"]] == [("set_value", "A1", CAPTION_FORMULA, "#Result Cashflows /Input")]
    assert out["ops"][0]["note"].startswith("user-chosen name and flags; this cell held a formula (" + CAPTION_FORMULA + "); it is replaced by the title") and "Open issue" in out["ops"][0]["note"], out["ops"][0]["note"]
    assert out["named"] == ["R!A2:C5"] and out["inserted_rows"] == {} and not [s for s in out["skipped"] if s.startswith("R!")], out["skipped"]


def test_an_old_plan_with_a_row_insert_is_still_ordered_and_resolved():
    old = [
        {"op": "insert_row", "action_id": "create_grid_titles", "rule_id": "STR-004", "sheet": "S", "row": 1, "columns": "A:A"},
        {"op": "set_value", "action_id": "create_grid_titles", "rule_id": "STR-004", "sheet": "S", "cell": "A1", "before": None, "after": "#T", "after_inserts": True, "insert_at": 1},
        {"op": "set_value", "action_id": "create_grid_titles", "rule_id": "STR-004", "sheet": "S", "cell": "C1", "before": None, "after": "#U"},
    ]
    ordered = prep._ordered(old)
    assert [o["op"] for o in ordered] == ["set_value", "insert_row", "set_value"]
    assert ordered[0]["cell"] == "C1" and ordered[2]["cell"] == "A1"  # the banded insert at A does not move C1


@needs_excel
def test_excel_still_applies_an_old_plan_with_a_row_insert(tmp_path):
    path = _wb(tmp_path / "old.xlsx", {"S": {"A1": "k", "A2": "a", "A3": "b", "B2": 1, "B3": 2, "D1": "=A2&B2"}})
    old = [
        {"op": "insert_row", "action_id": "create_grid_titles", "rule_id": "STR-004", "sheet": "S", "row": 1},
        {"op": "set_value", "action_id": "create_grid_titles", "rule_id": "STR-004", "sheet": "S", "cell": "A1", "before": None, "after": "#T", "after_inserts": True, "insert_at": 1},
    ]
    out = prep.apply_operations(path, tmp_path / "prep", old)
    assert out["status"] == "APPLIED", out
    ws = openpyxl.load_workbook(out["output_path"])["S"]
    assert ws["A1"].value == "#T" and ws["A2"].value == "k" and ws["A3"].value == "a" and ws["B3"].value == 1
    assert ws["D2"].value == "=A3&B3"  # Excel followed the insert
