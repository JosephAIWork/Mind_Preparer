"""1.11.2: Mind's MCP server as an optional extra. The answer texts below are real ones (UC-1, 2026-10-15)."""
from __future__ import annotations

import base64
import json
import time

import pytest

from app import mind_mcp

STRUCTURE = """Note: null
Dimensions[2]{Name,Length,IsLoopInstance,HasLabels}:
  SIM,1,false,false
  Policy_Data,6,false,false
Models[1]:
  - ModelIndex: 0
    Name: model_a7
    Instances[1]{Id}:
      Original_Instance
    GridCount: 117
Workflow: null"""

RESOLVE = """Note: null
ModelIndex: 0
InstanceId: Original_Instance
SpreadsheetName: Calculation
Address: AY35
OutOfGrids: false
GridIndex: 32
GridName: "Untitled(34,51)"
Row: 1
Col: 1
IsHeaderRow: false"""

VALUES = """Queried:
  ModelIndex: 0
TotalValuesCount: 3
ReturnedValuesCount: 3
Truncated: false
Values[4]:
  - [2]: Policy_Data,Values
  - [2]: 1,1
  - [2]: 2,1.5
  - [2]: 3,"a, b\""""

SCALAR = """TotalValuesCount: 1
Values[2]:
  - [1]: Values
  - [1]: 45657"""

HISTORY = """TotalRowsCount: 1
Rows[1]{RunId,Date,User,Result,RunType,ConsistencyStatus,AreResultsLoaded}:
  339fc7d8-d133,"2026-10-15T12:30:26.4059047",A User,Success,Basic,Consistent,true"""


def test_the_answer_format_is_read():
    dims = mind_mcp.table(STRUCTURE, "Dimensions")
    assert dims == [{"Name": "SIM", "Length": 1, "IsLoopInstance": False, "HasLabels": False}, {"Name": "Policy_Data", "Length": 6, "IsLoopInstance": False, "HasLabels": False}]
    assert mind_mcp.field(RESOLVE, "GridIndex") == 32 and mind_mcp.field(RESOLVE, "GridName") == "Untitled(34,51)" and mind_mcp.field(RESOLVE, "OutOfGrids") is False
    assert mind_mcp.values(VALUES) == [{"Policy_Data": 1, "Values": 1}, {"Policy_Data": 2, "Values": 1.5}, {"Policy_Data": 3, "Values": "a, b"}]
    row = mind_mcp.table(HISTORY, "Rows")[0]
    assert row["Result"] == "Success" and row["ConsistencyStatus"] == "Consistent" and row["Date"] == "2026-10-15T12:30:26.4059047"


class FakeMcp:
    def __init__(self, answers):
        self.answers = answers
        self.calls = []

    def call(self, tool, **args):
        self.calls.append((tool, args))
        return self.answers[tool] if not callable(self.answers[tool]) else self.answers[tool](args)


def test_a_cell_is_read_per_instance_and_a_cell_without_the_loop_counts_for_every_instance():
    fake = FakeMcp({"get_project_structure": STRUCTURE, "resolve_excel_address": RESOLVE, "get_cell_values": VALUES})
    got = mind_mcp.cell_instances(fake, "Calculation", "AY35", "Policy_Data")
    assert got["dimensioned"] and got["values"] == {1: 1, 2: 1.5, 3: "a, b"} and got["grid"] == "Untitled(34,51)"
    assert mind_mcp.instance_value(got, 2) == 1.5
    fake = FakeMcp({"get_project_structure": STRUCTURE, "resolve_excel_address": RESOLVE, "get_cell_values": SCALAR})
    got = mind_mcp.cell_instances(fake, "Calculation", "BA35", "Policy_Data")
    assert not got["dimensioned"] and [mind_mcp.instance_value(got, k) for k in (1, 6)] == [45657, 45657]


def test_dates_compare_with_excel_serials():
    assert mind_mcp._same("2024-12-31 00:00:00+00:00", 45657) and not mind_mcp._same(1.0, 2.0) and mind_mcp._same(None, None)


def test_the_run_errors_are_mapped_to_excel():
    log = "Run failed: Error while calculating [model_0_grid_7_R2_C3] ... MIND REFERENCE | Untitled(36,7)(37,9)"
    fake = FakeMcp({"get_run_log": log, "get_project_structure": STRUCTURE,
                    "get_cell_excel_address": "ModelName: model_a\nSpreadsheetName: P20\nAddress: H38"})
    out = mind_mcp.run_errors(fake, "run-1")
    assert {"mind": "model_0_grid_7_R2_C3", "excel": "P20!H38"} in out["cells"] and {"mind": "Untitled(36,7)(37,9)", "excel": "I37"} in out["cells"]
    assert out["lines"] and "Run failed" in out["lines"][0]


def _token(exp: float) -> str:
    body = base64.urlsafe_b64encode(json.dumps({"exp": exp}).encode()).decode().rstrip("=")
    return f"x.{body}.y"


def test_without_a_token_or_with_an_expired_one_the_app_says_so(tmp_path, monkeypatch):
    monkeypatch.delenv(mind_mcp.TOKEN_ENV, raising=False)
    monkeypatch.setattr(mind_mcp, "TOKEN_FILE", tmp_path / "missing")
    ok, why = mind_mcp.available()
    assert not ok and "no Mind MCP token" in why
    monkeypatch.setenv(mind_mcp.TOKEN_ENV, _token(time.time() - 10))
    ok, why = mind_mcp.available()
    assert not ok and "expired" in why
    monkeypatch.setenv(mind_mcp.TOKEN_ENV, _token(time.time() + 3600))
    assert mind_mcp.available() == (True, "")
    with pytest.raises(mind_mcp.MindMcpError):
        mind_mcp.project_id_from("not a link")
    assert mind_mcp.project_id_from("https://shared.milliman-mind.com/init/71/a8c57c8a-bc6c-4a3c-bca6-eca6245d2efa") == "a8c57c8a-bc6c-4a3c-bca6-eca6245d2efa"
