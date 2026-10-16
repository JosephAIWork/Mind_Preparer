from pathlib import Path

from app.inventory import build_analysis, loop_definitions, looplabels_definitions
from app.models import WorkbookAnalysis


def test_source_never_modified_and_hash_captured(tmp_path, plain_grid_xlsx):
    before = plain_grid_xlsx.read_bytes()
    analysis = build_analysis(plain_grid_xlsx, tmp_path / "work", analysis_id="a1")
    assert plain_grid_xlsx.read_bytes() == before  # source untouched
    assert analysis["source"]["sha256"]
    assert analysis["source"]["copy_path"] != str(plain_grid_xlsx)
    WorkbookAnalysis.model_validate(analysis)  # schema-shape conformance


def test_plain_grid_has_no_vba_and_one_formula(tmp_path, plain_grid_xlsx):
    analysis = build_analysis(plain_grid_xlsx, tmp_path / "work", analysis_id="a2")
    assert analysis["features"]["has_vba"] is False
    formulas = analysis["workbooks"][0]["formulas"]
    assert len(formulas) == 1
    assert formulas[0]["formula"] == "=A4+B3"


def test_macro_xlsm_detected_as_vba(tmp_path, macro_xlsm):
    analysis = build_analysis(macro_xlsm, tmp_path / "work", analysis_id="a3")
    assert analysis["features"]["has_vba"] is True
    assert analysis["risks"][0]["type"] == "VBA_PRESENT"


def test_loop_definitions_grouped_by_name(tmp_path, loops_ok_xlsx):
    analysis = build_analysis(loops_ok_xlsx, tmp_path / "work", analysis_id="a4")
    defs = loop_definitions(analysis)
    assert "Scenario" in defs
    assert len(defs["Scenario"]) == 2
    assert all(d["range_length"] == 5 for d in defs["Scenario"])

    labels = looplabels_definitions(analysis)
    assert labels[0]["name"] == "Scenario"
    assert labels[0]["range_length"] == 5


def test_hidden_sheet_and_merged_cells_inventoried(tmp_path, messy_workbook_xlsx):
    analysis = build_analysis(messy_workbook_xlsx, tmp_path / "work", analysis_id="a5")
    sheets = {s["name"]: s for s in analysis["workbooks"][0]["sheets"]}
    assert sheets["Notes"]["state"] == "hidden"
    assert "B2:C2" in sheets["Main"]["merged_cells"]
    assert len(sheets["Main"]["comments"]) == 1
    assert len(sheets["Main"]["hyperlinks"]) == 1
    assert sheets["Main"]["locked_cell_count"] >= 1


def test_analysis_is_json_serializable(tmp_path, messy_workbook_xlsx):
    """scripts/inventory_workbook.py dumps the whole analysis as JSON; every
    inventoried value (tab colours, grid rows, validations...) must survive."""
    import json

    analysis = build_analysis(messy_workbook_xlsx, tmp_path / "work", analysis_id="a6")
    assert analysis["workbooks"][0]["sheets"][0]["tab_color"] in ("00FF0000", "FF0000")
    json.dumps(analysis)


def test_the_workbook_caches_keep_only_the_latest_workbooks():
    """1.11.2: an unbounded cache kept every analysed version of a 36 MB model in memory (13+ GB)."""
    from app.inventory import _BoundedCache

    closed = []

    class Wb:
        def __init__(self, n):
            self.n = n

        def close(self):
            closed.append(self.n)

    c = _BoundedCache(2)
    c["v1"], c["v2"] = Wb(1), Wb(2)
    assert c.get("v1").n == 1  # v1 is now the most recent
    c["v3"] = Wb(3)
    assert set(c) == {"v1", "v3"} and closed == [2] and c.get("v2") is None
