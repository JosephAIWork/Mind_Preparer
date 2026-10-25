"""Unit tests for the pure helpers behind the 1.3.0 validators: Mind-style
grid detection (app/grids.py) and the formula tokenizer (app/formula_utils.py).
No workbook files needed."""
from app import formula_utils
from app.formula_utils import called_functions, cell_refs_in_formula, find_calls, formula_memo, mask_strings, name_tokens, parse_ref, storage_prefixes, unquote
from app.grids import detect_grids, looks_like_title, parse_flags, parse_title


def test_parse_title_and_flags():
    name, flags = parse_title("#My grid /Input /Resize.grp /Group.(Block name).0.1")
    assert name == "My grid"
    assert [f["name"] for f in flags] == ["input", "resize", "group"]
    assert flags[1]["args"] == ["grp"]
    assert flags[2]["args"] == ["Block name", "0", "1"]
    assert parse_flags("/BackupSource.(b 1)")[0]["args"] == ["b 1"]


def test_detect_grids_follows_the_documented_algorithm():
    occ = {
        (1, 1): "#Title /Input",  # title above the grid
        (2, 1): "H1", (2, 2): "H2",
        (3, 1): 1, (3, 2): 2,
        (4, 1): 3, (4, 2): "=A4+B3",
        (6, 1): "alone text",  # ignored by Mind
        (6, 4): "#Merged",
        (7, 4): "h",
        (8, 4): 1,
        (9, 4): "#Trapped",  # no empty row -> becomes part of the grid above
        (10, 4): "h",
        (11, 4): 2,
    }
    grids, standalone = detect_grids("S", occ)
    assert [s["text"] for s in standalone] == ["alone text"]
    first = grids[0]
    assert first["name"] == "Title" and first["flag_names"] == ["input"]
    assert first["ref"] == "A2:B4" and first["header_is_all_text"] and first["n_rows"] == 3
    merged = grids[1]
    assert merged["name"] == "Merged" and merged["ref"] == "D7:D11"
    assert merged["inner_title_cells"] == ["D9"]


def test_untitled_grid_gets_coordinates_name():
    grids, _ = detect_grids("S", {(5, 3): 1, (5, 4): 2, (6, 3): 3, (6, 4): 4})
    assert grids[0]["display_name"] == "untitled C5" and grids[0]["name"] is None


def test_called_functions_ignores_strings_and_strips_storage_prefixes():
    assert called_functions('="SUM(" & _xlfn.XLOOKUP(1,A:A,B:B) & \'My (sheet)\'!A1') == ["XLOOKUP"]
    assert called_functions("=_xludf.MM_LOOP(\"x\", A1:A3)+MM_SETSIZE(2,1)") == ["MM_LOOP", "MM_SETSIZE"]
    assert storage_prefixes("=_xludf.MM_LOOP(1)+_xlfn.IFS(1,2)") == {"_xludf.", "_xlfn."}


def test_find_calls_handles_nesting_and_unquote():
    calls = find_calls('=MM_RESULT(B2,"Loop",MM_DIMINDEX("Loop"))', "MM_RESULT")
    assert calls == [["B2", '"Loop"', 'MM_DIMINDEX("Loop")']]
    assert unquote('"Loop"') == "Loop" and unquote("B2") is None


def test_parse_ref_and_cell_refs():
    r = parse_ref("'My Sheet'!$B$2:C10")
    assert (r["sheet"], r["c1"], r["r1"], r["c2"], r["r2"], r["cells"]) == ("My Sheet", 2, 2, 3, 10, 18)
    assert parse_ref("A:A")["whole_column"] and parse_ref("3:5")["whole_row"]
    refs = cell_refs_in_formula('=SUM(Data!A1:A10,"B1:B2")+C3')
    assert [(x["sheet"], x["ref"]) for x in refs] == [("Data", "A1:A10"), (None, "C3")]


def test_excel_error_values_are_not_read_as_grid_titles():
    """'#N/A' and friends are results, not '#Titles'. Reading them as titles
    made STR-001 report a trapped title inside a data table and offer to insert
    a row straight through it (seen on a real 411x84 policy table)."""
    for err in ("#N/A", "#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!", "#SPILL!"):
        assert not looks_like_title(err), err
        assert not looks_like_title(f"  {err.lower()}  ")
    assert looks_like_title("#Premium table")
    assert looks_like_title("#Rates /Input")
    assert not looks_like_title(123) and not looks_like_title(None) and not looks_like_title("no hash")


def test_a_data_table_holding_error_values_stays_one_grid():
    occ = {(1, 1): "#Policy data", (2, 1): "Id", (2, 2): "Value"}
    for r in range(3, 8):
        occ[(r, 1)] = r
        occ[(r, 2)] = "#N/A" if r == 5 else r * 2
    grids, _ = detect_grids("S", occ)
    assert len(grids) == 1
    grid = grids[0]
    assert grid["name"] == "Policy data" and grid["ref"] == "A2:B7"
    # the '#N/A' in the middle is data, so nothing is "trapped" and no row
    # insert will ever be proposed to free it
    assert grid["inner_title_cells"] == []


def test_range_length_keeps_quoted_sheet_names():
    """1.10.2: `'Policy Data'!A2:A6` is 5 cells, not unresolvable (LOOP-003 flagged every
    Mind-loop version of UC-1 as needing a decision)."""
    from app.formula_utils import range_length

    assert range_length("'Policy Data'!A2:A6") == 5
    assert range_length("'Policy Data'!$AP$1:$AP$6") == 6
    assert range_length("Sheet1!A1:A3") == 3 and range_length("A1") == 1
    assert range_length("MyName") is None and range_length("") is None and range_length("A1:A5+0") is None


def test_the_formula_memo_gives_the_same_answers_and_hands_out_fresh_copies():
    """1.11.10: inside formula_memo() each formula is parsed once; the answers are the ones
    without it, a caller changing what it got changes nothing, and the memo ends with the scan."""
    formulas = ["=SUM(Sheet1!A1:B2,C3)+'My Sheet'!D4+Rate", '=IF(A1="x!B2",VLOOKUP(A1,$A:$B,2,0),1:3)', "=A1", "plain"]

    def answers():
        return [(cell_refs_in_formula(f), called_functions(f), name_tokens(f), mask_strings(f), mask_strings(f, False)) for f in formulas]

    plain = answers()
    with formula_memo():
        first = answers()
        first[0][0][0]["sheet"] = "changed"
        first[0][1].append("CHANGED")
        first[0][2].clear()
        with formula_memo():  # nested scans share it
            again = answers()
        assert formula_utils._MEMO is not None
    assert formula_utils._MEMO is None
    assert again == plain and answers() == plain

