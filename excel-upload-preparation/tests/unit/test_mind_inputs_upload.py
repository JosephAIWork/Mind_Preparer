"""The Mind input files (outside files made Mind inputs) go up through Mind's Inputs window:
app/mind_client.upload_inputs, called by app/mind_loop in every iteration after the model is
added and before the run. No browser and no Mind: a fake page plays the window as it was recorded
in real Mind on 2026-11-08 (an "Inputs" button left of Run; a table model > sheet > grid with
Status and File; Culture and CSV Separator; a folder input; "Match unmatched inputs" for a file no
grid name matches), on a fake clock so nothing really waits."""
from __future__ import annotations

from contextlib import contextmanager
from pathlib import Path

import pytest

from app import mind_client as mc
from app import mind_loop as ml

HEADER = "Element\tDL\tStatus\tFile\tReq.\tRead-Only\tDel."
TAIL = ("Culture\nAfrikaans\nAfrikaans - South Africa\nEnglish - United Kingdom\nZulu - South Africa\n"
        "CSV Separator\nDefault\n;\n,\n:\n~\nTab\nSplit between instances\nNo split\n\n"
        "Upload inputs or Drag & Drop\n\nSupported types are .csv, .txt, .zip and .parquet.\n\nBrowse")
# the page text recorded in real Mind (zzz_mindready_inputs_probe1, grid #CF_Test /Input on sheet Data)
REAL_BEFORE = ("Inputs\nRun\nmind_input_probe\nInputs\nSearch\n" + HEADER + "\n mind_input_probe\n Data\n"
               "CF_Test\t\t\n-\n\t\nNo file\n\t\n\t\n\t\n" + TAIL)
REAL_AFTER = ("Inputs\nRun\nmind_input_probe\nInputs\nSearch\n" + HEADER + "\n mind_input_probe\n Data\n"
              "CF_Test\t\t\nSUCCESS\n\tCF_Test.csv\t\n\t\n\t\n" + TAIL)


class Clock:
    """mind_client's view of time, moved on by the fake page's waits."""

    def __init__(self) -> None:
        self.now = 1_000_000.0

    def time(self) -> float:
        return self.now

    def sleep(self, s: float) -> None:
        self.now += s


@pytest.fixture
def clock(monkeypatch):
    c = Clock()
    monkeypatch.setattr(mc, "time", c)
    return c


class FakeLocator:
    def __init__(self, page: "FakeInputsPage", key: tuple) -> None:
        self.page, self.key = page, key

    @property
    def first(self) -> "FakeLocator":
        return self

    def count(self) -> int:
        return self.page._count(self.key)

    def click(self, timeout: int | None = None) -> None:
        self.page._click(self.key)


class FakeKeyboard:
    def __init__(self, page: "FakeInputsPage") -> None:
        self.page = page

    def press(self, key: str) -> None:
        self.page._press(key)


class FakeInputsPage:
    """A project whose model has the /Input grids `grids` [(sheet, grid)]. A file goes to the grid
    of its name; `delay_s` after the hand-over its row shows `final` (default SUCCESS); a file no
    grid is named after opens the 'Match unmatched inputs' popup (when `popup`). Until then Mind's
    "Inputs import" box is up (when `box`) and the rows show what they showed before (`before`:
    {grid: (status, file)} of an earlier upload -- as in real Mind on 10 November 2026)."""

    def __init__(self, clock: Clock, grids: list[tuple[str, str]], final: dict[str, str] | None = None, delay_s: float = 4.0,
                 inputs_button: bool = True, dom: bool = True, popup: bool = True, model: str = "tokyo_model",
                 culture: str | None = "English - United Kingdom", separator: str | None = "Default",
                 before: dict[str, tuple[str, str]] | None = None, box: bool = True) -> None:
        self.clock, self.grids, self.final, self.delay_s = clock, grids, final or {}, delay_s
        self.before, self.box = before or {}, box
        self.inputs_button, self.dom, self.popup_enabled, self.model = inputs_button, dom, popup, model
        self.culture, self.separator = culture, separator
        self.window = False
        self.popup_files: list[str] = []
        self.uploaded: dict[str, str] = {}
        self.uploaded_at: float | None = None
        self.handed: list[tuple[str, Path, list[str]]] = []
        self.calls: list[tuple] = []
        self.marked: str | None = None
        self.keyboard = FakeKeyboard(self)

    # --- the Playwright surface upload_inputs uses
    def wait_for_timeout(self, ms: int) -> None:
        self.clock.now += ms / 1000

    def get_by_role(self, role: str, name=None) -> FakeLocator:
        return FakeLocator(self, ("role", role, getattr(name, "pattern", name)))

    def locator(self, selector: str) -> FakeLocator:
        return FakeLocator(self, ("css", selector))

    def screenshot(self, path=None, full_page=False) -> None:
        self.calls.append(("screenshot", Path(path).name))

    def set_input_files(self, selector: str, files) -> None:
        folder = Path(files)
        assert folder.is_dir(), "the Inputs window's box takes a folder"
        names = sorted(p.name for p in folder.iterdir())
        self.handed.append((selector, folder, names))
        self.calls.append(("set_input_files", selector))
        self.uploaded_at = self.clock.now
        for f in names:
            if any(g == Path(f).stem for _s, g in self.grids):
                self.uploaded[Path(f).stem] = f
            elif self.popup_enabled:
                self.popup_files.append(f)

    def evaluate(self, js: str, arg=None):
        if js is mc._INPUT_ROWS_JS:
            if not self.dom:
                raise RuntimeError("no table on this page")
            return self._dom_rows() if self.window else []
        if js is mc._BOX_JS:
            marker, mark = arg
            if self.popup_files and mc._has(self._popup_text(), marker):
                self.marked = "popup" if mark else self.marked
                return {"text": self._popup_text(), "marked": "ok" if mark else "no"}
            if self.window and mc._has(self._window_text(), marker):
                self.marked = "window" if mark else self.marked
                return {"text": self._window_text(), "marked": "ok" if mark else "no"}
            return {"text": None, "marked": "nf"}
        if js is mc._INPUT_SETTINGS_JS:
            return {"culture": self.culture, "separator": self.separator} if self.window else {}
        if "document.body.innerText" in js:
            return self._body()
        return None  # the guided-tour mask removal

    # --- the fake Mind
    def _count(self, key: tuple) -> int:
        if key == ("role", "button", mc._INPUTS_BUTTON.pattern):
            return int(self.inputs_button)
        if key == ("css", "[data-mr-close='1']"):
            return int(self.marked is not None)
        return 0

    def _click(self, key: tuple) -> None:
        if key == ("role", "button", mc._INPUTS_BUTTON.pattern):
            self.calls.append(("click", "Inputs"))
            self.window = True
        elif key == ("css", "[data-mr-close='1']"):
            self.calls.append(("close", self.marked))
            if self.marked == "popup":
                self.popup_files = []  # cancelled: those files stay out
            elif self.marked == "window":
                self.window = False
            self.marked = None

    def _press(self, key: str) -> None:
        self.calls.append(("press", key))
        if key == "Escape":
            if self.popup_files:
                self.popup_files = []
            else:
                self.window = False

    def _importing(self) -> bool:
        return self.uploaded_at is not None and self.clock.now < self.uploaded_at + self.delay_s

    def _row(self, grid: str) -> tuple[str, str | None]:
        f = self.uploaded.get(grid)
        if f is None or self._importing():
            return self.before.get(grid, ("-", None))
        return self.final.get(grid, "SUCCESS"), f

    def _dom_rows(self) -> list[list[str]]:
        rows, sheet = [HEADER.split("\t"), [self.model]], None
        for s, g in self.grids:
            if s != sheet:
                rows.append([s])
                sheet = s
            status, f = self._row(g)
            rows.append([g, "", status, f or "No file", "", "", ""])
        return rows

    def _window_text(self) -> str:
        out, sheet = ["Inputs", "Search", HEADER, " " + self.model], None
        for s, g in self.grids:
            if s != sheet:
                out.append(" " + s)
                sheet = s
            status, f = self._row(g)
            out += [f"{g}\t\t", status] + ([f"\t{f}\t", "\t", "\t"] if f else ["\t", "No file", "\t", "\t", "\t"])
        return "\n".join(out) + "\n" + TAIL

    def _popup_text(self) -> str:
        return "Match unmatched inputs\n" + "\n".join(f"{f}\tSelect a grid" for f in self.popup_files) + "\nCancel\nConfirm"

    def _body(self) -> str:
        top = f"Inputs\nRun\n{self.model}"
        if not self.window:
            return top + "\nProjects > zzz_mindready_test"
        importing = "\nInputs import\nUpload\nLinks\nResize\nUploading files ..." if self.box and self._importing() else ""
        return top + "\n" + self._window_text() + importing + ("\n" + self._popup_text() if self.popup_files else "")


def _folder(tmp_path: Path, names: list[str], name: str = "Mind inputs") -> Path:
    folder = tmp_path / name
    folder.mkdir()
    for n in names:
        (folder / n).write_text("Month,保険料\n0,1000.5\n" if n.endswith(".csv") else "Month\t保険料\n0\t1,000.5\n", encoding="utf-8")
    return folder


GRIDS = [("BBA_01", "IN_BBA_01"), ("BBA_02", "IN_BBA_02"), ("VFA_01", "IN_VFA_01"), ("入力", "IN_入力")]
FILES = ["IN_BBA_01.csv", "IN_BBA_02.csv", "IN_VFA_01.txt", "IN_入力.csv"]


# --- reading the window ----------------------------------------------------------------------------
def test_the_window_text_recorded_in_real_mind_reads_back_as_rows():
    before = mc.parse_inputs_text(REAL_BEFORE)
    assert before == [{"model": "mind_input_probe", "sheet": "Data", "grid": "CF_Test", "status": "-", "file": None}]
    after = mc.parse_inputs_text(REAL_AFTER)
    assert after == [{"model": "mind_input_probe", "sheet": "Data", "grid": "CF_Test", "status": "SUCCESS", "file": "CF_Test.csv"}]
    assert not mc._inputs_settled(before, ["CF_Test.csv"]) and mc._inputs_settled(after, ["CF_Test.csv"])
    # one line for the report, without the ~800 cultures of the Culture box
    excerpt = mc._inputs_excerpt(REAL_AFTER)
    assert "CF_Test | SUCCESS | CF_Test.csv" in excerpt and "Afrikaans" not in excerpt and len(excerpt) <= 600


def test_table_rows_name_their_columns_and_headings_are_model_then_sheet():
    rows = [
        ["Element", "DL", "Status", "File", "Req.", "Read-Only", "Del."],
        ["tokyo_model"],  # a heading spans the row
        ["BBA_01", "", "", "", "", "", ""],
        ["IN_BBA_01", "", "SUCCESS", "IN_BBA_01.csv", "", "", ""],
        ["Element", "DL", "Status", "File", "Req.", "Read-Only", "Del."],  # a sticky copy of the header
        ["入力"],
        ["IN_入力", "", "-", "No file", "", "", ""],
    ]
    assert mc.parse_input_rows(rows) == [
        {"model": "tokyo_model", "sheet": "BBA_01", "grid": "IN_BBA_01", "status": "SUCCESS", "file": "IN_BBA_01.csv"},
        {"model": "tokyo_model", "sheet": "入力", "grid": "IN_入力", "status": "-", "file": None},
    ]
    assert mc.parse_input_rows([["no", "header"], ["x", "SUCCESS", "x.csv"]]) == []  # no header: nothing guessed


def test_a_file_is_found_by_its_name_or_by_the_grid_named_after_it():
    rows = [{"grid": "IN_A", "status": "SUCCESS", "file": "inputs/IN_A.csv"},
            {"grid": "IN_B", "status": "-", "file": None},
            {"grid": "IN_C", "status": "SUCCESS", "file": "IN_C_last_quarter.csv"}]
    assert mc._row_for(rows, "IN_A.csv")["grid"] == "IN_A"  # a folder in front of the name
    assert mc._row_for(rows, "in_a.CSV")["grid"] == "IN_A"  # any case
    assert mc._row_for(rows, "IN_B.csv")["grid"] == "IN_B"  # no file yet: the grid of its name
    assert mc._row_for(rows, "IN_C.csv") is None  # that grid shows another file
    assert mc._row_for(rows, "IN_D.csv") is None
    # a file with no grid of its name does not hold the wait; one still at '-' does
    assert mc._inputs_settled(rows[:1], ["IN_A.csv", "IN_D.csv"])
    assert not mc._inputs_settled(rows, ["IN_A.csv", "IN_B.csv"])


def test_input_files_are_the_csv_and_txt_files_of_the_folder_only(tmp_path):
    folder = _folder(tmp_path, ["IN_B.csv", "IN_A.TXT", "Mind inputs.zip", "~$IN_A.csv", "notes.xlsx"])
    (folder / "sub").mkdir()
    (folder / "sub" / "IN_C.csv").write_text("x", encoding="utf-8")
    assert [f.name for f in mc.input_files(folder)] == ["IN_A.TXT", "IN_B.csv"]
    assert mc.input_files(tmp_path / "nowhere") == [] and mc.input_files(None) == []


# --- upload_inputs ---------------------------------------------------------------------------------
def test_every_file_goes_to_the_grid_of_its_name_and_only_the_input_files_go_up(tmp_path, clock):
    folder = _folder(tmp_path, FILES + ["Mind inputs.zip"])
    (folder / "old").mkdir()
    (folder / "old" / "IN_BBA_01.csv").write_text("x", encoding="utf-8")
    page = FakeInputsPage(clock, GRIDS)
    res = mc.upload_inputs(page, folder, shots=tmp_path / "shots")
    assert res["ok"], res
    assert res["statuses"] == {g: "SUCCESS" for _s, g in GRIDS}
    assert res["unmatched"] == [] and res["failed"] == {} and res["waiting"] == [] and res["reason"] is None
    assert not res["timed_out"] and res["closed"] and not page.window
    # one folder holding the input files only (the zip and the subfolder stay out), removed afterwards
    [(selector, handed, names)] = page.handed
    assert selector == "input[type=file][webkitdirectory]" and names == sorted(FILES)
    assert handed != folder and not handed.exists()
    # the Inputs button, then the folder, then the window's own close button
    assert [c for c in page.calls if c[0] in ("click", "set_input_files", "close", "press")] == [
        ("click", "Inputs"), ("set_input_files", "input[type=file][webkitdirectory]"), ("close", "window")]
    assert ("screenshot", "inputs_before.png") in page.calls and ("screenshot", "inputs_after.png") in page.calls
    assert res["culture"] == "English - United Kingdom" and res["separator"] == "Default"  # left as Mind has them
    assert [r["sheet"] for r in res["rows"]] == ["BBA_01", "BBA_02", "VFA_01", "入力"] and res["rows"][0]["model"] == "tokyo_model"
    assert "入力 | IN_入力 | SUCCESS | IN_入力.csv" in res["text"] and "Afrikaans" not in res["text"]
    assert res["seconds"] < 60


def test_a_clean_folder_goes_up_as_it_is_and_the_page_text_serves_when_the_table_cannot_be_read(tmp_path, clock):
    folder = _folder(tmp_path, FILES)
    page = FakeInputsPage(clock, GRIDS, dom=False)
    res = mc.upload_inputs(page, folder)
    assert res["ok"], res
    assert page.handed[0][1] == folder
    assert res["statuses"] == {g: "SUCCESS" for _s, g in GRIDS}
    assert {(r["grid"], r["file"]) for r in res["rows"]} == {(Path(f).stem, f) for f in FILES}


def test_a_file_no_grid_matches_is_reported_and_the_popup_cancelled_never_matched(tmp_path, clock):
    folder = _folder(tmp_path, ["IN_BBA_01.csv", "IN_BBA_02.csv", "BBA_03 next quarter.csv"])
    page = FakeInputsPage(clock, GRIDS[:2])
    res = mc.upload_inputs(page, folder, shots=tmp_path / "shots")
    assert not res["ok"]
    assert res["unmatched"] == ["BBA_03 next quarter.csv"]
    assert res["statuses"] == {"IN_BBA_01": "SUCCESS", "IN_BBA_02": "SUCCESS"}  # the others still went in
    assert "BBA_03 next quarter.csv" in res["reason"] and "Match unmatched inputs" in res["text"]
    closes = [c for c in page.calls if c[0] in ("close", "press")]
    assert closes == [("close", "popup"), ("close", "window")]  # the popup's own Cancel, then the window
    assert ("screenshot", "inputs_unmatched.png") in page.calls and not page.window


def test_a_file_no_grid_matches_without_a_popup_is_reported_without_waiting_for_the_time_limit(tmp_path, clock):
    folder = _folder(tmp_path, ["IN_BBA_01.csv", "Stray.csv"])
    page = FakeInputsPage(clock, GRIDS[:1], popup=False)
    res = mc.upload_inputs(page, folder, timeout_s=300)
    assert not res["ok"] and res["unmatched"] == ["Stray.csv"] and not res["timed_out"]
    assert res["seconds"] < 30


def test_a_grid_mind_does_not_accept_is_not_ok(tmp_path, clock):
    folder = _folder(tmp_path, FILES[:2])
    page = FakeInputsPage(clock, GRIDS[:2], final={"IN_BBA_02": "ERROR"})
    res = mc.upload_inputs(page, folder)
    assert not res["ok"] and res["failed"] == {"IN_BBA_02": "ERROR"}
    assert res["statuses"] == {"IN_BBA_01": "SUCCESS", "IN_BBA_02": "ERROR"}
    assert res["reason"] == "Mind did not accept IN_BBA_02 (ERROR)" and res["closed"] and res["advice"] is None


def test_the_last_upload_s_rows_are_not_read_while_mind_takes_the_new_files(tmp_path, clock):
    """Real Mind, 10 November 2026: files replacing earlier ones of the same names. While the "Inputs import"
    box is up the rows still show the last upload's SUCCESS; this upload ends in an ERROR for IN_BBA_01."""
    folder = _folder(tmp_path, FILES[:2])
    before = {"IN_BBA_01": ("SUCCESS", "IN_BBA_01.csv"), "IN_BBA_02": ("SUCCESS", "IN_BBA_02.csv")}
    page = FakeInputsPage(clock, GRIDS[:2], final={"IN_BBA_01": "ERROR"}, delay_s=16, before=before)
    res = mc.upload_inputs(page, folder)
    assert not res["ok"] and res["failed"] == {"IN_BBA_01": "ERROR"} and res["statuses"]["IN_BBA_02"] == "SUCCESS"
    assert res["seconds"] >= 16  # read only once the box was gone
    # a box too quick to be seen: the rows are read after a while, not at once
    page = FakeInputsPage(clock, GRIDS[:2], final={"IN_BBA_01": "ERROR"}, delay_s=2, before=before, box=False)
    res = mc.upload_inputs(page, _folder(tmp_path, FILES[:2], name="again"))
    assert not res["ok"] and res["failed"] == {"IN_BBA_01": "ERROR"} and res["seconds"] >= mc._IMPORT_START_S


def test_a_raw_file_mind_refuses_for_its_columns_is_said_with_what_to_do(tmp_path, clock):
    """Mind's own words when a raw input file goes where a grid's file is expected (seen 10 November 2026)."""
    folder = _folder(tmp_path, FILES[:1])
    page = FakeInputsPage(clock, GRIDS[:1], final={"IN_BBA_01": "ERROR"})
    body = page._body
    page._body = lambda: body() + ("\nFailed: Input control error: The file IN_BBA_01.csv doesn't have proper headers, there are 1340 columns "
                                   "while grid 'IN_BBA_01' of instance Original_Instance has 93 columns." if page.window else "")
    res = mc.upload_inputs(page, folder)
    assert not res["ok"] and res["failed"] == {"IN_BBA_01": "ERROR"}
    assert res["advice"].startswith("Mind refused IN_BBA_01.csv: it has 1,340 columns, the grid IN_BBA_01 has 93. It looks like a raw input file.")
    assert "Make next quarter's Mind input files" in res["advice"]
    assert res["reason"] == "Mind did not accept IN_BBA_01 (ERROR). " + res["advice"]


class FakeLogTags:
    """The LOG tag of each ERROR row: a click opens Mind's "CSV Logs" box with its reason; Escape closes it."""

    def __init__(self, page: "FakeInputsPage", messages: list[str]) -> None:
        self.page, self.messages = page, messages

    def count(self) -> int:
        return len(self.messages) if self.page.window and not self.page._importing() else 0

    def nth(self, j: int) -> "FakeLogTags":
        tag = FakeLogTags(self.page, self.messages)
        tag.j = j
        return tag

    def is_visible(self) -> bool:
        return True

    def click(self, timeout: int | None = None) -> None:
        self.page.log_open = self.messages[self.j]


def test_mind_s_own_reason_for_a_refused_file_is_read_from_its_log(tmp_path, clock):
    """Real Mind, 10 November 2026: the reason is not on the page; the row's LOG tag opens "CSV Logs"."""
    refusal = ("Failed: Input control error: The file IN_BBA_01.csv doesn't have proper headers, there are 1340 columns "
               "while grid 'IN_BBA_01' of instance Original_Instance has 93 columns.")
    page = FakeInputsPage(clock, GRIDS[:2], final={"IN_BBA_01": "ERROR"})
    page.log_open = None
    tags = FakeLogTags(page, [refusal])
    locator, body, press = page.locator, page._body, page._press
    page.locator = lambda sel: tags if sel == mc._CSV_LOG_TAG else locator(sel)
    page._body = lambda: body() + (f"\nCSV Logs\nClose modal\n{page.log_open}" if page.log_open else "")

    def escape(key: str) -> None:
        if key == "Escape" and page.log_open:
            page.log_open = None  # the log box closes, the window stays
            return
        press(key)

    page._press = escape
    res = mc.upload_inputs(page, _folder(tmp_path, FILES[:2]))
    assert not res["ok"] and res["failed"] == {"IN_BBA_01": "ERROR"} and res["errors"] == [refusal]
    assert res["advice"].startswith("Mind refused IN_BBA_01.csv: it has 1,340 columns, the grid IN_BBA_01 has 93.")
    assert res["reason"] == "Mind did not accept IN_BBA_01 (ERROR). " + res["advice"] and res["closed"] and not page.window


def test_no_inputs_button_means_the_model_has_no_input_grid(tmp_path, clock):
    page = FakeInputsPage(clock, GRIDS, inputs_button=False)
    res = mc.upload_inputs(page, _folder(tmp_path, FILES))
    assert not res["ok"] and "no Inputs button" in res["reason"]
    assert page.handed == [] and res["statuses"] == {}
    assert 30 <= res["seconds"] < 40  # looked for it for half a minute, no more


def test_mind_still_busy_at_the_time_limit(tmp_path, clock):
    page = FakeInputsPage(clock, GRIDS[:2], delay_s=10_000)
    res = mc.upload_inputs(page, _folder(tmp_path, FILES[:2]), timeout_s=60)
    assert not res["ok"] and res["timed_out"]
    assert res["waiting"] == FILES[:2] and res["unmatched"] == []
    assert "had not finished after 60 s" in res["reason"] and "IN_BBA_01.csv" in res["reason"]
    assert res["closed"] and not page.window


@pytest.mark.parametrize("culture,separator", [("German - Germany", "Default"), ("French - France", "Default"),
                                               ("English - South Africa", "Default"), ("English - United Kingdom", ";")])
def test_nothing_goes_up_when_mind_would_read_the_numbers_differently(tmp_path, clock, culture, separator):
    page = FakeInputsPage(clock, GRIDS, culture=culture, separator=separator)
    res = mc.upload_inputs(page, _folder(tmp_path, FILES))
    assert not res["ok"] and culture in res["reason"] and "not uploaded" in res["reason"]
    assert page.handed == [] and res["closed"] and not page.window  # Culture and Separator untouched, window closed


@pytest.mark.parametrize("culture", ["English - United States", "Japanese - Japan", "Hebrew - Israel", None])
def test_cultures_that_read_a_decimal_point_or_an_unreadable_box_go_ahead(tmp_path, clock, culture):
    res = mc.upload_inputs(FakeInputsPage(clock, GRIDS, culture=culture), _folder(tmp_path, FILES))
    assert res["ok"], res


def test_an_empty_folder_uploads_nothing(tmp_path, clock):
    page = FakeInputsPage(clock, GRIDS)
    res = mc.upload_inputs(page, _folder(tmp_path, ["Mind inputs.zip"]))
    assert not res["ok"] and res["reason"].startswith("no .csv or .txt file") and page.calls == []


def test_add_template_also_takes_the_plural_button():
    clicked: list[str] = []

    class Loc:
        def __init__(self, label: str | None) -> None:
            self.label = label

        @property
        def first(self):
            return self

        def click(self, timeout=None):
            if self.label is None:
                raise TimeoutError("no such button")
            clicked.append(self.label)

        def count(self):
            return 1 if clicked else 0

    class Page:
        def get_by_role(self, role, name=None):
            if name.search("Run"):
                return Loc("Run")
            return Loc("Add templates to Milliman Mind" if name.search("Add templates to Milliman Mind") else None)

        def wait_for_timeout(self, ms):
            pass

        def evaluate(self, *_a):
            return None

    assert mc.add_template(Page(), timeout_s=1) and clicked == ["Add templates to Milliman Mind"]
    assert mc._ADD_TEMPLATE.search("Add template to Milliman Mind")


# --- the loop: after the model is added, before the run, in every iteration --------------------------
@pytest.fixture
def fake_mind(monkeypatch):
    """Every Mind step of app/mind_client replaced by a recorder; upload_inputs takes the folder's
    files (or, with state['inputs_ok'] False, finds no grid for the first one)."""
    calls: list[tuple] = []
    state = {"inputs_ok": True}

    class Page:
        def set_viewport_size(self, *_a, **_k):
            pass

    @contextmanager
    def session(headed=False):
        calls.append(("session",))
        yield None, Page()

    def upload_inputs(page, folder, timeout_s=300, shots=None):
        files = [f.name for f in mc.input_files(folder)]
        calls.append(("upload_inputs", tuple(files)))
        state["timeout_s"] = timeout_s
        if state["inputs_ok"]:
            return {"ok": True, "statuses": {Path(f).stem: "SUCCESS" for f in files}, "unmatched": [], "failed": {},
                    "files": files, "closed": True, "reason": None}
        return {"ok": False, "statuses": {}, "unmatched": files[:1], "failed": {}, "files": files, "closed": True,
                "reason": f"Mind found no /Input grid named like {files[0]}"}

    fakes = {
        "mind_session": session,
        "open_manager": lambda page: calls.append(("open_manager",)) or True,
        "ensure_folder": lambda page, folder=None: True,
        "create_blank_project": lambda page, name: calls.append(("create", name)) or f"/init/{name}",
        "open_project": lambda page, init, timeout_s=90: calls.append(("open_project",)) or "empty",
        "upload_and_convert": lambda page, wb, timeout_s=420, shots=None: calls.append(("upload_and_convert", Path(wb).name)) or {
            "success": True, "steps": {"Upload": "Complete", "Convert": "Complete", "Test": "Complete"}, "errors": [], "text": ""},
        "add_template": lambda page, timeout_s=150: calls.append(("add_template",)) or True,
        "upload_inputs": upload_inputs,
        "template_names": lambda page, stem: calls.append(("template_names",)) or [],
        "run_model": lambda page, timeout_s=300, shots=None: calls.append(("run_model",)) or {
            "completed": True, "audit_consistent": True, "seconds": 3.0, "text": ""},
        "export_summary": lambda page: {"opened": False, "elements": None, "tabs": []},
        "screenshot": lambda page, path: None,
        "delete_project": lambda page, name, folder=None: calls.append(("delete", name)) or "deleted",
    }
    for k, v in fakes.items():
        monkeypatch.setattr(ml.mc, k, v)
    return calls, state


def _iteration(tmp_path: Path, inputs_dir: Path | None, n: int = 1):
    cfg = ml.LoopConfig(source=tmp_path / "model.xlsm", work_dir=tmp_path / "loop", inputs_dir=inputs_dir)
    report: dict = {"projects": []}
    events: list[dict] = []
    out = ml._mind_iteration(cfg, tmp_path / f"iter_{n}", tmp_path / "model_prepared.xlsm", {"unnamed": 0, "unnamed_anchors": []},
                             ["CF"], n, "20261109_010101", report, events.append)
    return out, report, events


def _steps(calls):
    return [c[0] for c in calls if c[0] in ("upload_and_convert", "add_template", "upload_inputs", "template_names", "run_model", "delete")]


def _decide(mind):
    it = {"apply": {"status": "APPLIED"}, "numbers": {"ran": True, "match": True}, "names": {"fragmented": False}, "mind": mind}
    return ml.decide(it, [], set(), set())


def test_the_loop_uploads_the_inputs_after_the_template_and_before_the_run(tmp_path, fake_mind):
    calls, state = fake_mind
    folder = _folder(tmp_path, FILES + ["Mind inputs.zip"])
    (folder / "IN_BBA_01.csv").write_bytes(b"0," * 3 * 2**20)  # 6 MB: big inputs get longer
    out, report, events = _iteration(tmp_path, folder)
    assert "error" not in out, out
    assert _steps(calls) == ["upload_and_convert", "add_template", "upload_inputs", "template_names", "run_model"]
    assert ("upload_inputs", tuple(sorted(FILES, key=str.casefold))) in calls
    assert state["timeout_s"] >= 300 + 30 * 6
    assert out["inputs_upload"]["ok"] and report["projects"][0]["ran"]
    assert any("Inputs window" in e["message"] for e in events) and any("SUCCESS" in e["message"] for e in events)
    assert _decide(out)["verdict"] == "converged"


def test_inputs_mind_does_not_take_stop_the_iteration_before_the_run_and_say_which(tmp_path, fake_mind):
    calls, state = fake_mind
    state["inputs_ok"] = False
    out, report, _events = _iteration(tmp_path, _folder(tmp_path, FILES))
    assert out["error"] == "Mind did not take the input files: Mind found no /Input grid named like IN_BBA_01.csv"
    assert "run_model" not in _steps(calls) and "template_names" not in _steps(calls)
    assert _steps(calls)[-1] == "delete" and report["projects"][0]["status"] == "deleted"  # never run: cleaned up
    d = _decide(out)
    assert d["verdict"] == "stuck" and "IN_BBA_01.csv" in d["reason"]


def test_no_inputs_folder_no_upload_and_an_empty_one_is_noted(tmp_path, fake_mind):
    calls, _state = fake_mind
    out, _report, _events = _iteration(tmp_path, None)
    assert "upload_inputs" not in _steps(calls) and "inputs_upload" not in out and "run_model" in _steps(calls)
    calls.clear()
    out, _report, events = _iteration(tmp_path, _folder(tmp_path, ["Mind inputs.zip"]), n=2)
    assert "upload_inputs" not in _steps(calls) and out["inputs_upload"]["skipped"] and "error" not in out
    assert any("no Mind input file" in e["message"] for e in events)


def test_a_missing_inputs_folder_stops_before_mind_is_opened(tmp_path, fake_mind):
    calls, _state = fake_mind
    out, report, _events = _iteration(tmp_path, tmp_path / "gone")
    assert "missing" in out["error"] and calls == [] and report["projects"] == []
    assert _decide(out)["verdict"] == "stuck"
    cfg = ml.LoopConfig(source=tmp_path / "model.xlsm", work_dir=tmp_path / "loop", inputs_dir=tmp_path / "gone")
    rep = ml.run_loop(cfg)
    assert rep["verdict"] == "stuck" and "missing" in rep["reason"] and rep["iterations"] == []


def test_every_iteration_that_uploads_the_model_uploads_the_inputs_again(tmp_path, fake_mind, section_heading_grids_xlsx, monkeypatch):
    calls, _state = fake_mind

    def copy_only(source, work_dir, ops, prefer_excel=True):
        copy, _ = ml.make_immutable_copy(source, work_dir)
        return {"status": "APPLIED", "output_path": copy, "applied": [], "failed": [], "message": "", "verified_opens_in_excel": True}

    monkeypatch.setattr(ml.prep, "apply_operations", copy_only)
    verdicts = iter(["retry", "converged"])
    monkeypatch.setattr(ml, "decide", lambda *a, **k: {"verdict": next(verdicts), "reason": "test", "enable": [], "disable": []})
    cfg = ml.LoopConfig(source=section_heading_grids_xlsx, work_dir=tmp_path / "loop", max_iterations=3, use_assistant=False,
                        check_numbers=False, inputs_dir=_folder(tmp_path, FILES))
    report = ml.run_loop(cfg)
    assert report["verdict"] == "converged" and len(report["iterations"]) == 2
    steps = _steps(calls)
    assert steps == ["upload_and_convert", "add_template", "upload_inputs", "template_names", "run_model"] * 2
    assert len({c[1] for c in calls if c[0] == "create"}) == 2  # a new project each time
    assert all(it["inputs_upload"]["ok"] and it["mind"]["inputs_upload"]["ok"] for it in report["iterations"])
    assert report["config"]["inputs_dir"] == str(cfg.inputs_dir)
    assert "inputs=ok" in ml.summarize(report)


def test_the_upload_has_its_own_stage_only_once_the_estimate_lists_it(monkeypatch):
    from app import estimate

    assert ml._inputs_stage() == ({"stage": "inputs"} if "inputs" in estimate.MIND_STAGE_TITLES else {})
    monkeypatch.setitem(estimate.MIND_STAGE_TITLES, "inputs", "Uploading the input files in Mind")
    assert ml._inputs_stage() == {"stage": "inputs"}


def test_the_summary_says_when_mind_did_not_take_the_inputs():
    report = {"verdict": "stuck", "reason": "x", "source": "m.xlsm", "iterations": [
        {"n": 1, "actions": [], "apply": {"status": "APPLIED"}, "names": {}, "numbers": {"ran": True, "match": True},
         "mind": {"convert": {"success": True}, "inputs_upload": {"ok": False, "reason": "no grid"}}, "decision": {"verdict": "stuck", "reason": "x"}}]}
    assert "inputs=FAIL" in ml.summarize(report)
    report["iterations"][0]["mind"].pop("inputs_upload")
    assert "inputs=" not in ml.summarize(report)
