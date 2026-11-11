"""Comment spacing (tools/comment_spacing.py): the spacing between the words of comments and
docstrings follows a list of gaps given to the tool (never the user's own rule file); stamp touches
only comment text, is idempotent, keeps line endings, a byte order mark, tabs and trailing spaces;
unstamp restores single spaces; check reports the run; the pre-commit hook stamps staged files."""
from __future__ import annotations

import importlib.util
import os
import py_compile
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]
GIT_ROOT = ROOT.parent
TOOL = ROOT / "tools" / "comment_spacing.py"
INSTALLER = ROOT / "tools" / "install_hooks.py"
HOOK = GIT_ROOT / ".githooks" / "pre-commit"


def _load(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


cs = _load(TOOL, "mind_ready_comment_spacing")

GAPS = [1, 2, 3]  # marks at gaps 0, 2, 5, 9, 11, 14, 18, ...

PY_SOURCE = '''"""A module docstring with several words on this line.
And a second line with more words in it here.
"""
import os  # a trailing comment with some words

TEXT = "a string with  two spaces and words that must not change"
OTHER = """not a docstring: this string has words too"""


def f(a, b):
    """The function docstring has words.

    More words on another line of the docstring.
    """
    # a plain comment with several words inside it
    return a + b  # noqa: E501  a directive comment stays as it is


class K:
    """Class docstring words here."""

    x = 1  # type: ignore  another directive


# >>> doctest_lines(are, skipped) in comments too? no: only in docstrings
def g():
    """Doc with a doctest.

    >>> g(1,  2) stays
    ... continued line stays
    Normal words after the doctest.
    """
    return 1
'''

TS_SOURCE = '''// A line comment with several words in it
/** The JSDoc of a function with words on this line
 *  and more words on the next line of the block
 */
export function f(a: string): string {
  const s = "a string with words that must not change";
  const t = `a template with // not a comment and ${a} words`;
  const re = /\\/\\/ not a comment either/g;
  const half = 4 / 2 / 1; // a trailing comment with words
  /* a block comment on one line with words */
  // eslint-disable-next-line no-console  a directive stays as it is
  console.log(s, t, re, half);
  return a.replace(re, "x"); /** another block with words */
}
/// <reference types="node" /> a directive comment stays
'''


def _gap_widths(text: str, suffix: str) -> list[int]:
    lines = cs._split_lines(text)
    return [g.width for g in cs.find_gaps(lines, cs.regions_for(text, suffix))]


# ----------------------------------------------------------------------------------------------
# stamp


def test_stamp_writes_the_marks_at_the_right_gaps_across_comment_lines_python(tmp_path):
    path = tmp_path / "m.py"
    path.write_bytes(PY_SOURCE.encode("utf-8"))
    assert cs.stamp_file(path, GAPS) == "stamped"
    text = path.read_text(encoding="utf-8")
    widths = _gap_widths(text, ".py")
    marks = set(cs.mark_positions(GAPS, len(widths)))
    assert len(widths) > 30 and sorted(marks)[:8] == [0, 2, 5, 9, 11, 14, 18, 20]
    assert all(width == (2 if i in marks else 1) for i, width in enumerate(widths))
    assert widths.count(2) == len(marks)
    # the marks cross the comment lines: gaps 0, 2, 5 on the first docstring line, 9, 11, 14 on the second...
    assert text.startswith('"""A  module docstring  with several words  on this line.\n'
                           'And a  second line  with more words  in it here.\n"""\n')
    assert "import os  # a trailing  comment with  some words\n" in text
    assert "# a plain comment with  several words  inside it\n" in text
    assert "    More words  on another  line of the  docstring.\n" in text
    # code and strings untouched
    assert 'TEXT = "a string with  two spaces and words that must not change"' in text
    assert 'OTHER = """not a docstring: this string has words too"""' in text
    assert "return a + b  # noqa: E501  a directive comment stays as it is" in text
    assert "x = 1  # type: ignore  another directive" in text
    assert ">>> g(1,  2) stays" in text and "... continued line stays" in text
    py_compile.compile(str(path), doraise=True)


def test_stamp_skips_code_strings_templates_and_regexes_typescript(tmp_path):
    path = tmp_path / "m.ts"
    path.write_bytes(TS_SOURCE.encode("utf-8"))
    assert cs.stamp_file(path, GAPS) == "stamped"
    text = path.read_text(encoding="utf-8")
    widths = _gap_widths(text, ".ts")
    marks = set(cs.mark_positions(GAPS, len(widths)))
    assert widths and all(width == (2 if i in marks else 1) for i, width in enumerate(widths))
    assert 'const s = "a string with words that must not change";' in text
    assert "const t = `a template with // not a comment and ${a} words`;" in text
    assert "const re = /\\/\\/ not a comment either/g;" in text
    assert "const half = 4 / 2 / 1;" in text
    assert "// eslint-disable-next-line no-console  a directive stays as it is" in text
    assert '/// <reference types="node" /> a directive comment stays' in text
    assert "// A  line comment" in text  # the first gap of the file is a mark
    assert text.count("  ") - TS_SOURCE.count("  ") == len(marks)


def test_typescript_comment_spans_see_through_strings_templates_and_regexes():
    spans = cs.typescript_comment_spans(TS_SOURCE)
    texts = [TS_SOURCE[a:b] for a, b, _ in spans]
    assert texts[0] == "// A line comment with several words in it"
    assert any(t.startswith("/** The JSDoc") for t in texts)
    assert "/* a block comment on one line with words */" in texts
    assert "/** another block with words */" in texts
    assert not any("not a comment" in t for t in texts)
    assert len(texts) == 7


def test_stamp_is_idempotent_and_restamps_an_edited_file_from_its_start(tmp_path):
    path = tmp_path / "m.py"
    path.write_bytes(PY_SOURCE.encode("utf-8"))
    cs.stamp_file(path, GAPS)
    once = path.read_bytes()
    assert cs.stamp_file(path, GAPS) == "unchanged"
    assert path.read_bytes() == once
    # a comment added at the top shifts every gap: the file is stamped again from its start
    path.write_bytes(b"# a new comment with some words at the top of the file\n" + once)
    assert not cs.check_file(path, GAPS).intact
    assert cs.stamp_file(path, GAPS) == "stamped"
    report = cs.check_file(path, GAPS)
    assert report.intact and report.extra == 0


# ----------------------------------------------------------------------------------------------
# unstamp


@pytest.mark.parametrize("name,source", [("m.py", PY_SOURCE), ("m.ts", TS_SOURCE)])
def test_unstamp_restores_the_file_byte_for_byte(tmp_path, name, source):
    path = tmp_path / name
    original = source.encode("utf-8")
    path.write_bytes(original)
    assert cs.stamp_file(path, GAPS) == "stamped"
    assert path.read_bytes() != original
    assert cs.unstamp_file(path) == "restored"
    assert path.read_bytes() == original
    assert cs.unstamp_file(path) == "unchanged"


# ----------------------------------------------------------------------------------------------
# check


def _long_python(words_per_line: int = 8, lines: int = 60) -> str:
    comment = " ".join(f"w{i}" for i in range(words_per_line))
    return "x = 1\n" + "".join(f"# {comment} line {n}\n" for n in range(lines))


def test_check_finds_the_sequence_and_reports_zero_when_absent(tmp_path):
    path = tmp_path / "long.py"
    path.write_text(_long_python(), encoding="utf-8")
    before = cs.check_file(path, GAPS)
    assert before.gaps == 60 * 9 and before.doubles == 0 and before.matched == 0
    assert before.repetitions == 0 and not before.intact and before.chance == 1.0
    cs.stamp_file(path, GAPS)
    after = cs.check_file(path, GAPS)
    assert after.intact and not after.partial and after.extra == 0
    assert after.matched == after.expected == len(cs.mark_positions(GAPS, after.gaps))
    assert after.repetitions == (after.matched - 1) // len(GAPS) >= 1
    assert 0 < after.chance < 1e-30 and after.chance_text.startswith("1 in ")
    # a stray double space in the middle breaks the run
    text = path.read_text(encoding="utf-8")
    broken = text.replace("w3 w4", "w3  w4", 1) if "w3 w4" in text else text.replace("w3  w4", "w3 w4", 1)
    path.write_text(broken, encoding="utf-8")
    damaged = cs.check_file(path, GAPS)
    assert not damaged.intact
    assert damaged.matched < after.matched or damaged.extra > 0


def test_a_short_file_gets_a_partial_run_and_check_says_so(tmp_path):
    path = tmp_path / "short.py"
    path.write_text("# only four words here\n", encoding="utf-8")  # 3 gaps: marks at 0 and 2 only
    cs.stamp_file(path, GAPS)
    assert path.read_text(encoding="utf-8") == "# only  four words  here\n"
    report = cs.check_file(path, GAPS)
    assert report.partial and report.intact and report.repetitions == 0
    assert report.matched == report.expected == 2
    out = subprocess.run([sys.executable, str(TOOL), "check", "--pattern", "1,2,3", str(path)],
                         capture_output=True, text=True, check=True).stdout
    assert "partial" in out and "0 full repetitions" in out and "chance by accident 1 in 3" in out


def test_mark_positions_follow_the_gaps_and_repeat():
    assert cs.mark_positions([1, 30, 10], 100) == [0, 2, 33, 44, 46, 77, 88, 90]
    assert cs.mark_positions([1, 2, 3], 0) == []


# ----------------------------------------------------------------------------------------------
# byte-exactness


def test_crlf_bom_tabs_and_trailing_spaces_are_kept(tmp_path):
    path = tmp_path / "crlf.py"
    body = ("﻿# first comment with words here   \r\n"
            "x = 1\t# tab\tbefore and after words \r\n"
            '"""Docstring words on a line\r\n\twith a tab and words\r\n"""\r\n')
    path.write_bytes(body.encode("utf-8"))
    assert cs.stamp_file(path, GAPS) == "stamped"
    data = path.read_bytes()
    assert data.startswith(b"\xef\xbb\xbf")
    assert b"\n" not in data.replace(b"\r\n", b"")  # every line end is still CRLF
    assert data.count(b"\r\n") == 5
    assert b"words here   \r\n" in data  # trailing spaces untouched
    assert b"\t# tab\tbefore" in data  # tabs untouched
    assert data.replace(b"  ", b" ") == body.encode("utf-8").replace(b"  ", b" ")
    assert cs.unstamp_file(path) == "restored"
    assert path.read_bytes() == body.encode("utf-8")
    py_compile.compile(str(path), doraise=True)


def test_a_file_that_does_not_parse_or_is_not_utf8_is_left_alone(tmp_path):
    broken = tmp_path / "broken.py"
    broken.write_text("def f(:\n    # comment words here\n", encoding="utf-8")
    assert cs.stamp_file(broken, GAPS).startswith("left alone")
    latin = tmp_path / "latin.py"
    latin.write_bytes(b"# caf\xe9 words here\n")
    assert cs.stamp_file(latin, GAPS) == "left alone: not UTF-8"
    assert latin.read_bytes() == b"# caf\xe9 words here\n"


# ----------------------------------------------------------------------------------------------
# the rule file, the repository rules


def test_without_the_rule_file_the_tool_does_nothing_and_says_so(tmp_path):
    path = tmp_path / "m.py"
    path.write_text(PY_SOURCE, encoding="utf-8")
    env = dict(os.environ, LOCALAPPDATA=str(tmp_path / "empty"))
    out = subprocess.run([sys.executable, str(TOOL), "stamp", str(path)], capture_output=True, text=True,
                         env=env, check=True).stdout
    assert out.strip().startswith("comment spacing: no rule file") and "nothing done" in out
    assert path.read_text(encoding="utf-8") == PY_SOURCE
    # with the rule file in place, the same command stamps
    (tmp_path / "empty" / "MindReady").mkdir(parents=True)
    (tmp_path / "empty" / "MindReady" / cs.RULE_FILE_NAME).write_text('{"gaps": [1, 2, 3]}', encoding="utf-8")
    out = subprocess.run([sys.executable, str(TOOL), "stamp", "--quiet", str(path)], capture_output=True,
                         text=True, env=env, check=True).stdout
    assert out.strip() == "comment spacing: 1 stamped"
    assert cs.check_file(path, GAPS).intact


def test_the_repository_rules_keep_to_the_source_folders(tmp_path):
    root = tmp_path
    for rel in ("excel-upload-preparation/app/x.py", "excel-upload-preparation/scripts/y.py",
                "excel-upload-preparation/tools/z.py", "MindReadyUI/src/app/a.ts"):
        (root / rel).parent.mkdir(parents=True, exist_ok=True)
        (root / rel).write_text("# words\n")
        assert cs.in_scope(root / rel, root), rel
    for rel in ("excel-upload-preparation/tests/unit/test_x.py", "excel-upload-preparation/app/conftest.py",
                "MindReadyUI/src/app/a.spec.ts", "MindReadyUI/dist/a.ts", "MindReadyUI/node_modules/p/a.ts",
                "excel-upload-preparation/app/fixtures/f.py", "excel-upload-preparation/docs/d.py",
                "excel-upload-preparation/app/page.html", "MindReadyUI/src/app/types.d.ts",
                "excel-upload-preparation/app/generated/g.py"):
        (root / rel).parent.mkdir(parents=True, exist_ok=True)
        (root / rel).write_text("# words\n")
        assert not cs.in_scope(root / rel, root), rel
    assert len(cs.repository_files(root)) == 4


# ----------------------------------------------------------------------------------------------
# the hook and its installer


def _git(*args: str, cwd: Path, env: dict) -> subprocess.CompletedProcess:
    return subprocess.run(["git", *args], cwd=str(cwd), env=env, capture_output=True, text=True)


@pytest.mark.skipif(shutil.which("git") is None, reason="git is not on the path")
def test_the_pre_commit_hook_stamps_staged_sources_and_the_installer_switches_it_on(tmp_path):
    clone = tmp_path / "clone"
    (clone / ".githooks").mkdir(parents=True)
    shutil.copy(HOOK, clone / ".githooks" / "pre-commit")
    tool_home = clone / "excel-upload-preparation" / "tools"
    tool_home.mkdir(parents=True)
    shutil.copy(TOOL, tool_home / "comment_spacing.py")
    shutil.copy(INSTALLER, tool_home / "install_hooks.py")
    (clone / "excel-upload-preparation" / "app").mkdir()
    source = clone / "excel-upload-preparation" / "app" / "m.py"
    source.write_text(PY_SOURCE, encoding="utf-8")
    test_file = clone / "excel-upload-preparation" / "tests" / "unit" / "test_m.py"
    test_file.parent.mkdir(parents=True)
    test_file.write_text("# words in a test file stay as they are\n", encoding="utf-8")
    env = dict(os.environ, LOCALAPPDATA=str(tmp_path / "data"), GIT_AUTHOR_NAME="t", GIT_AUTHOR_EMAIL="t@t",
               GIT_COMMITTER_NAME="t", GIT_COMMITTER_EMAIL="t@t", HOME=str(tmp_path))
    (tmp_path / "data" / "MindReady").mkdir(parents=True)
    (tmp_path / "data" / "MindReady" / cs.RULE_FILE_NAME).write_text('{"gaps": [1, 2, 3]}', encoding="utf-8")
    assert _git("init", "-q", cwd=clone, env=env).returncode == 0
    _git("config", "core.autocrlf", "false", cwd=clone, env=env)

    line = subprocess.run([sys.executable, str(tool_home / "install_hooks.py")], capture_output=True, text=True,
                          env=env, check=True).stdout.strip()
    assert line.endswith("core.hooksPath = .githooks")
    assert _git("config", "core.hooksPath", cwd=clone, env=env).stdout.strip() == ".githooks"

    assert _git("add", "-A", cwd=clone, env=env).returncode == 0
    commit = _git("commit", "-q", "-m", "first", cwd=clone, env=env)
    assert commit.returncode == 0, commit.stderr + commit.stdout
    assert "comment spacing:" in commit.stdout + commit.stderr and "stamped" in commit.stdout + commit.stderr
    committed = _git("show", "HEAD:excel-upload-preparation/app/m.py", cwd=clone, env=env).stdout
    assert "A  module docstring" in committed  # the first gap of the file is a mark
    assert source.read_text(encoding="utf-8") == committed  # the working tree holds the same text
    assert cs.check_file(source, GAPS).intact
    assert test_file.read_text(encoding="utf-8") == "# words in a test file stay as they are\n"
    assert _git("status", "--porcelain", cwd=clone, env=env).stdout.strip() == ""

    # a second commit of an unchanged file: the hook finds nothing to do and the commit goes through
    source.write_text(source.read_text(encoding="utf-8") + "\n# one more comment with words\n", encoding="utf-8")
    _git("add", "excel-upload-preparation/app/m.py", cwd=clone, env=env)
    second = _git("commit", "-q", "-m", "second", cwd=clone, env=env)
    assert second.returncode == 0, second.stderr
    assert cs.check_file(source, GAPS).intact
