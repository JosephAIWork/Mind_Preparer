"""Comment  spacing: keeps  the spacing between words inside comments and docstrings consistent
across the Python and TypeScript sources.

    python tools/comment_spacing.py stamp <paths...>      apply the spacing rule (idempotent)
    python tools/comment_spacing.py check <paths...>      report how far each file follows the rule
    python  tools/comment_spacing.py unstamp <paths...>    back to single spaces
    python tools/comment_spacing.py stamp-staged          the pre-commit hook: the  staged copies only

The spacing rule lives  outside the repository, in the app's data  folder (comment_spacing.json, a
list of numbers under "gaps"); without that file the  tool does nothing and says so. Tests  give the
rule with --pattern 1,2,3 instead. Only the text of comments and docstrings is touched -- never code,
strings or line ends -- and only where one or two spaces stand between two words; everything else in
the file stays byte for byte (line endings, a byte order mark, tabs, trailing spaces).

The rule is applied in parts: the run of gaps starts again  at every line that begins at column 0
with code (a top-level definition, class, statement), so an edit inside one function changes the
spacing of that part only, and a diff stays within what was edited.

--repo-rules keeps to the folders the rule applies to (app, scripts, tools, MindReadyUI/src) and
leaves tests, specs, fixtures, generated code and third-party code alone; --all walks those folders.
The pre-commit hook in .githooks runs `stamp-staged`: it reads each staged file from the index,
stamps that copy and puts it back, so  what is not staged is  never swept into the commit (the
working-tree file  is updated only  when it  still equals the staged copy; a partial commit, `git
commit -- <paths>`, is left alone). tools/install_hooks.py switches the hook on for a clone.
"""
from __future__ import annotations

import argparse
import io
import json
import os
import re
import subprocess
import sys
import tokenize
from dataclasses import dataclass, field
from pathlib import Path

RULE_FILE_NAME = "comment_spacing.json"
REPO_ROOT = Path(__file__).resolve().parents[2]
ALLOWED_FOLDERS = (
    Path("excel-upload-preparation/app"),
    Path("excel-upload-preparation/scripts"),
    Path("excel-upload-preparation/tools"),
    Path("MindReadyUI/src"),
)
SKIPPED_PARTS = {
    "tests", "test", "fixtures", "fixture", "dist", "node_modules", "generated", "vendor",
    "third_party", "third-party", "__pycache__", ".angular",
}
PY_DIRECTIVES = ("!", "noqa", "type:", "fmt:", "pragma", "pylint", "mypy:", "isort:", "-*-", "coding", "flake8")
TS_DIRECTIVES = ("/", "#", "eslint", "@ts-", "prettier", "tslint", "istanbul", "global ", "jshint", "jscs")
TS_REGEX_AFTER_WORDS = {
    "return", "typeof", "instanceof", "in", "of", "new", "delete", "void", "throw", "case",
    "do", "else", "yield", "await",
}
# a  Python line  that begins at column 0 with one of these tokens is code (a comment, a string that
# runs on from the line before, or an empty line is not): the run of  gaps starts again there
_PY_CODE_TOKENS = {tokenize.NAME, tokenize.OP, tokenize.NUMBER} | {
    getattr(tokenize, name) for name in ("FSTRING_START", "TSTRING_START") if hasattr(tokenize, name)
}


# ----------------------------------------------------------------------------------------------
# The  rule


def rule_file_path() -> Path:
    """Where  the rule  lives: the app's data folder (%LOCALAPPDATA%\\MindReady), never the repository."""
    base = os.environ.get("LOCALAPPDATA") or str(Path.home())
    return Path(base) / "MindReady" / RULE_FILE_NAME


def load_gaps(explicit: str | None = None) -> list[int] | None:
    """The  list of  gaps: from --pattern when given, else from the rule file; None when there is none."""
    if explicit:
        gaps = [int(part) for part in re.split(r"[,\s]+", explicit.strip()) if part]
        return gaps or None
    path = rule_file_path()
    if not path.is_file():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    gaps = data.get("gaps") if isinstance(data, dict) else None
    if not isinstance(gaps, list) or not gaps or not all(isinstance(g, int) and g >= 0 for g in gaps):
        return None
    return gaps


def mark_positions(gaps: list[int], count: int, phase: int = 0) -> list[int]:
    """The  gap indexes  (0-based, in order) that carry a double space when a part has `count` gaps:
    a mark, then gaps[0] single spaces, a mark, gaps[1] single spaces, ... repeating. `phase` starts
    the list of gaps  at another entry (check looks for a run that starts in  the middle of the list)."""
    marks: list[int] = []
    position = 0
    index = phase
    while position < count:
        marks.append(position)
        position += gaps[index % len(gaps)] + 1
        index += 1
    return marks


# ----------------------------------------------------------------------------------------------
# Where the  comment text is


@dataclass
class Region:
    """A  stretch of  comment or docstring text on one line: [start, end) as character columns."""

    line: int  # 0-based index into the file's lines
    start: int
    end: int


@dataclass
class Gap:
    """A  run of  one or two spaces between two words inside a Region."""

    line: int
    col: int
    width: int


def _split_lines(text: str) -> list[str]:
    """The  file as  lines that end with '\\n' (the last one may not), joined back byte for byte."""
    if not text:
        return []
    lines = text.split("\n")
    out = [line + "\n" for line in lines[:-1]]
    if lines[-1]:
        out.append(lines[-1])
    return out


def _line_body(line: str) -> str:
    return line.rstrip("\r\n")


def _is_directive(content: str, prefixes: tuple[str, ...]) -> bool:
    low = content.lstrip().lower()
    return any(low.startswith(p) for p in prefixes)


def _text_region(lines: list[str], line_no: int, start: int, end: int | None = None) -> Region | None:
    """The  region from  `start` (after the comment marker) to `end` (or the line's end), with the
    leading spaces dropped; None when nothing is left."""
    body = _line_body(lines[line_no])
    stop = len(body) if end is None else min(end, len(body))
    while start < stop and body[start] in " \t":
        start += 1
    return Region(line_no, start, stop) if start < stop else None


def _block_line_regions(lines: list[str], first_line: int, first_col: int, last_line: int, last_col: int,
                        skip_line=None) -> list[Region]:
    """Regions  of a  text block that spans lines (a docstring, a /* */ comment): the first line from
    first_col, the last line up to last_col, the lines between whole, each without its leading
    whitespace and a  leading '*' (JSDoc style)."""
    regions: list[Region] = []
    for line_no in range(first_line, last_line + 1):
        body = _line_body(lines[line_no])
        start = first_col if line_no == first_line else 0
        end = last_col if line_no == last_line else len(body)
        stripped = body[start:end].lstrip(" \t")
        if line_no != first_line and stripped.startswith("*"):
            start = len(body[:end]) - len(stripped) + 1
        if skip_line and skip_line(stripped):
            continue
        region = _text_region(lines, line_no, start, end)
        if region:
            regions.append(region)
    return regions


def _docstring_line_skipped(stripped: str) -> bool:
    """Lines inside a docstring that must not change: doctests and \\N{...} escapes."""
    return stripped.startswith((">>>", "...")) or "\\N{" in stripped


def _python_tokens(text: str) -> list[tokenize.TokenInfo]:
    return list(tokenize.generate_tokens(io.StringIO(text).readline))


def _python_docstring_spans(text: str, lines: list[str]) -> list[tuple[int, int, int, int]]:
    import ast

    spans: list[tuple[int, int, int, int]] = []
    for node in ast.walk(ast.parse(text)):
        if isinstance(node, (ast.Module, ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef)):
            body = getattr(node, "body", None)
            if body and isinstance(body[0], ast.Expr) and isinstance(body[0].value, ast.Constant) \
                    and isinstance(body[0].value.value, str):
                value = body[0].value
                spans.append((value.lineno, _byte_to_char(lines, value.lineno, value.col_offset),
                              value.end_lineno, _byte_to_char(lines, value.end_lineno, value.end_col_offset)))
    return spans


def python_regions(text: str, lines: list[str]) -> list[Region]:
    """Comment  tokens and  docstrings of a Python file (tokenize + ast); SyntaxError when it does not parse."""
    docstring_spans = _python_docstring_spans(text, lines)

    def in_docstring(row: int, col: int) -> bool:
        for l1, c1, l2, c2 in docstring_spans:
            if (l1, c1) <= (row, col) < (l2, c2):
                return True
        return False

    regions: list[Region] = []
    for tok in _python_tokens(text):
        if tok.type == tokenize.COMMENT:
            row, col = tok.start
            content = tok.string.rstrip("\r\n")
            if _is_directive(content[1:], PY_DIRECTIVES):
                continue
            region = _text_region(lines, row - 1, col + 1, col + len(content))
            if region:
                regions.append(region)
        elif tok.type == tokenize.STRING and in_docstring(*tok.start):
            opening = re.match(r"[A-Za-z]*('''|\"\"\"|'|\")", tok.string)
            if not opening:
                continue
            quote = opening.group(1)
            (r1, c1), (r2, c2) = tok.start, tok.end
            regions.extend(_block_line_regions(lines, r1 - 1, c1 + opening.end(), r2 - 1, c2 - len(quote),
                                               skip_line=_docstring_line_skipped))
    regions.sort(key=lambda r: (r.line, r.start))
    return regions


def python_part_starts(text: str) -> list[int]:
    """The  0-based lines  where a new part begins: a code token at column 0 (a def, a class, an import,
    an assignment, a closing bracket) -- not a comment, not a line inside a string."""
    return sorted({tok.start[0] - 1 for tok in _python_tokens(text) if tok.start[1] == 0 and tok.type in _PY_CODE_TOKENS})


def _byte_to_char(lines: list[str], row: int, byte_col: int) -> int:
    """ast  gives byte  columns (UTF-8); the lines are text."""
    line = lines[row - 1] if row - 1 < len(lines) else ""
    return len(line.encode("utf-8")[:byte_col].decode("utf-8", errors="ignore"))


_TS_WORD = re.compile(r"[A-Za-z_$][\w$]*")
_TS_NUMBER = re.compile(r"[0-9][\w.]*")


def _typescript_scan(text: str) -> tuple[list[tuple[int, int, str]], list[tuple[int, int]]]:
    """(comment  spans as  (start, end, 'line' | 'block'), literal spans as (start, end)) of a TypeScript
    file. A small scanner: strings, template literals (with ${ } inside), regular expression literals
    and comments; a '/' starts a  regular expression when what comes before cannot be divided."""
    spans: list[tuple[int, int, str]] = []
    literals: list[tuple[int, int]] = []
    n = len(text)
    i = 0
    last = "punct"  # what the previous  significant token was: 'value' (division follows)  or 'punct'
    template_depths: list[int] = []  # for each open ${ }: how  many braces are open inside it

    def scan_template(j: int) -> int | None:
        """From just after a '`' (or  after the '}' that closes an expression):  the index after the
        closing '`', or None when a '${' opened an expression (the expression is then scanned as code)."""
        start = j
        while j < n:
            ch = text[j]
            if ch == "\\":
                j += 2
            elif ch == "`":
                literals.append((start, j + 1))
                return j + 1
            elif ch == "$" and j + 1 < n and text[j + 1] == "{":
                template_depths.append(0)
                literals.append((start, j + 2))
                return -(j + 2)
            else:
                j += 1
        literals.append((start, j))
        return j

    while i < n:
        ch = text[i]
        two = text[i:i + 2]
        if two == "//":
            end = i
            while end < n and text[end] not in "\r\n":
                end += 1
            spans.append((i, end, "line"))
            i = end
        elif two == "/*":
            close = text.find("*/", i + 2)
            end = n if close < 0 else close + 2
            spans.append((i, end, "block"))
            i = end
        elif ch in "'\"":
            j = i + 1
            while j < n and text[j] != ch and text[j] not in "\r\n":
                j += 2 if text[j] == "\\" else 1
            literals.append((i, min(j + 1, n)))
            i = min(j + 1, n)
            last = "value"
        elif ch == "`":
            result = scan_template(i + 1)
            if result is None or result >= 0:
                i = n if result is None else result
                last = "value"
            else:
                i = -result
                last = "punct"
        elif ch == "}" and template_depths:
            if template_depths[-1] == 0:
                template_depths.pop()
                result = scan_template(i + 1)
                if result is None or result >= 0:
                    i = n if result is None else result
                    last = "value"
                else:
                    i = -result
                    last = "punct"
            else:
                template_depths[-1] -= 1
                i += 1
                last = "value"
        elif ch == "{" and template_depths:
            template_depths[-1] += 1
            i += 1
            last = "punct"
        elif ch == "/" and last == "punct":
            j = i + 1
            in_class = False
            ok = False
            while j < n and text[j] not in "\r\n":
                c = text[j]
                if c == "\\":
                    j += 2
                    continue
                if in_class:
                    in_class = c != "]"
                elif c == "[":
                    in_class = True
                elif c == "/":
                    ok = True
                    break
                j += 1
            if ok:
                j += 1
                while j < n and text[j].isalpha():
                    j += 1
                literals.append((i, j))
                i = j
                last = "value"
            else:
                i += 1
        elif ch.isspace():
            i += 1
        else:
            word = _TS_WORD.match(text, i)
            if word:
                i = word.end()
                last = "punct" if word.group(0) in TS_REGEX_AFTER_WORDS else "value"
                continue
            number = _TS_NUMBER.match(text, i)
            if number:
                i = number.end()
                last = "value"
                continue
            last = "value" if ch in ")]}" else "punct"
            i += 1
    return spans, literals


def typescript_comment_spans(text: str) -> list[tuple[int, int, str]]:
    """(start,  end, 'line'  | 'block') character spans of the comments of a TypeScript file."""
    return _typescript_scan(text)[0]


def _line_starts(lines: list[str]) -> list[int]:
    starts = []
    offset = 0
    for line in lines:
        starts.append(offset)
        offset += len(line)
    return starts


def typescript_regions(text: str, lines: list[str]) -> list[Region]:
    """The  text of  // comments, /* */ comments and JSDoc of a TypeScript file."""
    starts = _line_starts(lines)

    def locate(pos: int) -> tuple[int, int]:
        lo, hi = 0, len(starts) - 1
        while lo < hi:
            mid = (lo + hi + 1) // 2
            if starts[mid] <= pos:
                lo = mid
            else:
                hi = mid - 1
        return lo, pos - starts[lo]

    regions: list[Region] = []
    for start, end, kind in typescript_comment_spans(text):
        if kind == "line":
            row, col = locate(start)
            if _is_directive(text[start + 2:end], TS_DIRECTIVES):
                continue
            region = _text_region(lines, row, col + 2, col + (end - start))
            if region:
                regions.append(region)
        else:
            inner = text[start + 2:end - 2] if text[end - 2:end] == "*/" else text[start + 2:end]
            if _is_directive(inner.lstrip("*"), TS_DIRECTIVES[2:]):
                continue
            r1, c1 = locate(start)
            r2, c2 = locate(end)
            opening = 3 if text[start:start + 3] == "/**" and text[start:start + 4] != "/**/" else 2
            closing = 2 if text[end - 2:end] == "*/" else 0
            regions.extend(_block_line_regions(lines, r1, c1 + opening, r2, c2 - closing))
    regions.sort(key=lambda r: (r.line, r.start))
    return regions


def typescript_part_starts(text: str, lines: list[str]) -> list[int]:
    """The  0-based lines  where a new part begins: the line's first character is code at column 0
    (an import, a decorator, a class, a closing brace) -- not inside a comment, a string or a  template."""
    comments, literals = _typescript_scan(text)
    covered = sorted([(a, b) for a, b, _ in comments] + literals)
    starts: list[int] = []
    k = 0
    for line_no, (line, offset) in enumerate(zip(lines, _line_starts(lines))):
        if not line or line[0] in " \t\r\n":
            continue
        while k < len(covered) and covered[k][1] <= offset:
            k += 1
        if k < len(covered) and covered[k][0] <= offset < covered[k][1]:
            continue
        starts.append(line_no)
    return starts


def regions_for(text: str, suffix: str) -> list[Region]:
    lines = _split_lines(text)
    if suffix == ".py":
        return python_regions(text, lines)
    if suffix in (".ts", ".tsx", ".js", ".mjs"):
        return typescript_regions(text, lines)
    return []


def part_starts_for(text: str, suffix: str) -> list[int]:
    """The  lines where  the run of gaps starts again (see python_part_starts, typescript_part_starts)."""
    if suffix == ".py":
        return python_part_starts(text)
    if suffix in (".ts", ".tsx", ".js", ".mjs"):
        return typescript_part_starts(text, _split_lines(text))
    return []


# ----------------------------------------------------------------------------------------------
# The gaps, and the spacing itself


def find_gaps(lines: list[str], regions: list[Region]) -> list[Gap]:
    """Runs  of one  or two spaces with a non-space character on both sides, inside the regions, in file order."""
    gaps: list[Gap] = []
    for region in regions:
        body = _line_body(lines[region.line])
        i = region.start
        while i < region.end:
            if body[i] == " ":
                j = i
                while j < region.end and body[j] == " ":
                    j += 1
                width = j - i
                if 1 <= width <= 2 and i > region.start and j < region.end and not body[i - 1].isspace() \
                        and not body[j].isspace():
                    gaps.append(Gap(region.line, i, width))
                i = j
            else:
                i += 1
    return gaps


def parts_of(gaps: list[Gap], starts: list[int]) -> list[tuple[int, int]]:
    """(first  gap index,  number of gaps) of each part that has gaps: the gaps are cut at the lines in
    `starts`; a part with no gaps is dropped."""
    parts: list[tuple[int, int]] = []
    bounds = sorted(set(starts))
    begin = 0
    for bound in bounds + [None]:
        end = begin
        while end < len(gaps) and (bound is None or gaps[end].line < bound):
            end += 1
        if end > begin:
            parts.append((begin, end - begin))
        begin = end
    return parts


def _rewrite(lines: list[str], gaps: list[Gap], widths: list[int]) -> str:
    """The  file with  each gap set to its new width (1 or 2 spaces)."""
    by_line: dict[int, list[tuple[Gap, int]]] = {}
    for gap, width in zip(gaps, widths):
        by_line.setdefault(gap.line, []).append((gap, width))
    out = list(lines)
    for line_no, items in by_line.items():
        line = lines[line_no]
        pieces = []
        cursor = 0
        for gap, width in sorted(items, key=lambda it: it[0].col):
            pieces.append(line[cursor:gap.col])
            pieces.append(" " * width)
            cursor = gap.col + gap.width
        pieces.append(line[cursor:])
        out[line_no] = "".join(pieces)
    return "".join(out)


def _gaps_and_parts(text: str, suffix: str) -> tuple[list[str], list[Gap], list[tuple[int, int]]]:
    lines = _split_lines(text)
    gaps = find_gaps(lines, regions_for(text, suffix))
    return lines, gaps, parts_of(gaps, part_starts_for(text, suffix))


def marks_for(gaps_rule: list[int], parts: list[tuple[int, int]]) -> set[int]:
    """The  gap indexes  that carry a double space: the run starts again at each part."""
    marks: set[int] = set()
    for first, count in parts:
        marks.update(first + p for p in mark_positions(gaps_rule, count))
    return marks


def stamp_text(text: str, gaps_rule: list[int], suffix: str) -> str:
    lines, gaps, parts = _gaps_and_parts(text, suffix)
    marks = marks_for(gaps_rule, parts)
    return _rewrite(lines, gaps, [2 if i in marks else 1 for i in range(len(gaps))])


def unstamp_text(text: str, suffix: str) -> str:
    lines = _split_lines(text)
    gaps = find_gaps(lines, regions_for(text, suffix))
    return _rewrite(lines, gaps, [1] * len(gaps))


@dataclass
class Report:
    """What  `check` finds  in one file."""

    gaps: int = 0
    doubles: int = 0
    expected: int = 0  # the marks the rule lays over this file
    matched: int = 0  # marks found, over every part: at the rule's place, or as a run that moved
    extra: int = 0  # double spaces the rule does not explain
    repetitions: int = 0  # full passes  through the list of gaps, over every part
    parts: int = 0  # parts with gaps (the  run starts again at each)
    parts_intact: int = 0  # parts whose  run is exactly where the rule lays  it
    parts_moved: int = 0  # parts whose run is complete but starts at another gap or  another entry of the list
    moved_at: list[tuple[int, int]] = field(default_factory=list)  # (gap, entry) where  each moved run starts
    intact: bool = False  # every part intact and no extra double
    moved: bool = False  # not intact, but every part's run is there, moved, and no extra double
    partial: bool = False  # no part is long enough for one full pass
    error: str | None = None

    def as_dict(self) -> dict:
        return dict(self.__dict__)


def _run_from(doubles: set[int], count: int, gaps_rule: list[int], start: int, phase: int) -> tuple[int, bool]:
    """How  many marks  of a run that begins at gap `start` with the list's entry `phase` are present,
    and whether the run reached the part's end (every expected mark there) or stopped at a missing  one."""
    matched = 0
    position = start
    index = phase
    while position < count:
        if position not in doubles:
            return matched, False
        matched += 1
        position += gaps_rule[index % len(gaps_rule)] + 1
        index += 1
    return matched, True


def check_text(text: str, gaps_rule: list[int], suffix: str) -> Report:
    _lines, gaps, parts = _gaps_and_parts(text, suffix)
    report = Report(gaps=len(gaps), doubles=sum(1 for g in gaps if g.width == 2), parts=len(parts))
    explained: set[int] = set()
    longest_part = 0
    for first, count in parts:
        doubles = {i - first for i in range(first, first + count) if gaps[i].width == 2}
        expected = mark_positions(gaps_rule, count)
        report.expected += len(expected)
        longest_part = max(longest_part, len(expected))
        matched, complete = _run_from(doubles, count, gaps_rule, 0, 0)
        best = (matched, complete, 0, 0)
        if not complete:
            # the  run may  have moved: a word added or removed before it, a part cut at its head
            for start in sorted(doubles):
                if start == 0:
                    continue
                for phase in range(len(gaps_rule)):
                    m, done = _run_from(doubles, count, gaps_rule, start, phase)
                    if (m, done) > (best[0], best[1]):
                        best = (m, done, start, phase)
        matched, complete, start, phase = best
        report.matched += matched
        report.repetitions += max(0, matched - 1) // len(gaps_rule)
        if complete and start == 0 and phase == 0:
            report.parts_intact += 1
        elif complete and matched > 0:
            report.parts_moved += 1
            report.moved_at.append((first + start, phase))
        if matched:
            position, index = start, phase
            for _ in range(matched):
                explained.add(first + position)
                position += gaps_rule[index % len(gaps_rule)] + 1
                index += 1
    report.extra = sum(1 for i, g in enumerate(gaps) if g.width == 2 and i not in explained)
    report.intact = report.parts > 0 and report.parts_intact == report.parts and report.extra == 0
    report.moved = not report.intact and report.parts > 0 and report.parts_intact + report.parts_moved == report.parts and report.extra == 0
    report.partial = longest_part < len(gaps_rule) + 1
    return report


# ----------------------------------------------------------------------------------------------
# Files


def read_file(path: Path) -> tuple[str, bool] | None:
    """(text,  had_bom); None  when the file is not UTF-8."""
    return decode(path.read_bytes())


def decode(data: bytes) -> tuple[str, bool] | None:
    bom = data.startswith(b"\xef\xbb\xbf")
    try:
        return data[3:].decode("utf-8") if bom else data.decode("utf-8"), bom
    except UnicodeDecodeError:
        return None


def encode(text: str, bom: bool) -> bytes:
    return (b"\xef\xbb\xbf" if bom else b"") + text.encode("utf-8")


def write_file(path: Path, text: str, bom: bool) -> None:
    path.write_bytes(encode(text, bom))


def stamp_bytes(data: bytes, gaps_rule: list[int], suffix: str) -> tuple[bytes | None, str]:
    """(the  stamped bytes  or None when nothing changes, 'stamped' | 'unchanged' | a reason it was left alone)."""
    loaded = decode(data)
    if loaded is None:
        return None, "left alone: not UTF-8"
    text, bom = loaded
    try:
        new_text = stamp_text(text, gaps_rule, suffix)
    except (SyntaxError, tokenize.TokenError, ValueError) as exc:
        return None, f"left alone: does not parse ({exc.__class__.__name__})"
    if new_text == text:
        return None, "unchanged"
    return encode(new_text, bom), "stamped"


def stamp_file(path: Path, gaps_rule: list[int]) -> str:
    """'stamped',  'unchanged', or  a reason the file was left alone."""
    data, outcome = stamp_bytes(path.read_bytes(), gaps_rule, path.suffix.lower())
    if data is not None:
        path.write_bytes(data)
    return outcome


def unstamp_file(path: Path) -> str:
    loaded = read_file(path)
    if loaded is None:
        return "left alone: not UTF-8"
    text, bom = loaded
    try:
        new_text = unstamp_text(text, path.suffix.lower())
    except (SyntaxError, tokenize.TokenError, ValueError) as exc:
        return f"left alone: does not parse ({exc.__class__.__name__})"
    if new_text == text:
        return "unchanged"
    write_file(path, new_text, bom)
    return "restored"


def check_file(path: Path, gaps_rule: list[int]) -> Report:
    loaded = read_file(path)
    if loaded is None:
        return Report(error="not UTF-8")
    try:
        return check_text(loaded[0], gaps_rule, path.suffix.lower())
    except (SyntaxError, tokenize.TokenError, ValueError) as exc:
        return Report(error=f"does not parse ({exc.__class__.__name__})")


def in_scope(path: Path, root: Path = REPO_ROOT) -> bool:
    """Whether  the repository  rules apply to this file: a .py or .ts source under one of the allowed
    folders, not a test, spec, fixture, generated or third-party file."""
    try:
        relative = path.resolve().relative_to(root.resolve())
    except ValueError:
        return False
    name = relative.name.lower()
    if not (name.endswith(".py") or name.endswith(".ts")):
        return False
    if name.endswith((".spec.ts", ".d.ts")) or name.startswith("test_") or name == "conftest.py":
        return False
    if any(part.lower() in SKIPPED_PARTS for part in relative.parts[:-1]):
        return False
    return any(relative.is_relative_to(folder) for folder in ALLOWED_FOLDERS)


def repository_files(root: Path = REPO_ROOT) -> list[Path]:
    files: list[Path] = []
    for folder in ALLOWED_FOLDERS:
        base = root / folder
        if base.is_dir():
            files.extend(p for p in sorted(base.rglob("*")) if p.is_file() and in_scope(p, root))
    return files


# ----------------------------------------------------------------------------------------------
# The  staged copies  (the pre-commit hook)


def _git(root: Path, *args: str, data: bytes | None = None) -> bytes:
    result = subprocess.run(["git", *args], cwd=str(root), input=data, capture_output=True)
    if result.returncode != 0:
        raise RuntimeError(f"git {' '.join(args)}: {result.stderr.decode('utf-8', 'replace').strip()}")
    return result.stdout


def stamp_staged(gaps_rule: list[int], root: Path = REPO_ROOT) -> str:
    """Stamps  the staged  copy of every source the repository rules apply to, in the index: the blob
    git will commit is read, stamped and put back with its mode; the working-tree file gets the same
    spacing  only when it still equals the staged copy (an unstaged edit  is never staged by this). A
    partial  commit (`git commit -- <paths>`, which git  runs on a temporary index that it does not keep)
    is left  alone. Returns the one line to print."""
    index_file = Path(os.environ.get("GIT_INDEX_FILE") or "")
    if index_file.name.startswith("next-index"):
        return "comment spacing: a partial commit (paths given): left for the next full commit"
    staged = [f for f in _git(root, "diff", "--cached", "--name-only", "-z", "--diff-filter=ACMR").decode("utf-8").split("\0") if f]
    counts: dict[str, int] = {}
    for rel in staged:
        path = root / rel
        if not in_scope(path, root):
            continue
        entry = _git(root, "ls-files", "--stage", "-z", "--", rel).decode("utf-8").split("\0")[0]
        if not entry:
            continue
        mode, blob = entry.split()[:2]
        data = _git(root, "cat-file", "blob", blob)
        new_data, outcome = stamp_bytes(data, gaps_rule, path.suffix.lower())
        counts[outcome] = counts.get(outcome, 0) + 1
        if new_data is None:
            continue
        new_blob = _git(root, "hash-object", "-w", "--stdin", data=new_data).decode("ascii").strip()
        _git(root, "update-index", "--cacheinfo", f"{mode},{new_blob},{rel}")
        # the  working-tree file: the same spacing, when it still holds exactly what was staged
        # (compared through git, so line-ending conversion does not count as a difference)
        if path.is_file() and _git(root, "hash-object", "--", rel).decode("ascii").strip() == blob:
            stamp_file(path, gaps_rule)
    if not counts:
        return "comment spacing: nothing staged that the rule applies to"
    return "comment spacing: " + ", ".join(f"{n} {what}" for what, n in sorted(counts.items()))


# ----------------------------------------------------------------------------------------------
# Command line


def _paths(args: argparse.Namespace) -> list[Path]:
    paths = [Path(p) for p in args.paths]
    if args.files_from:
        stream = sys.stdin if args.files_from == "-" else open(args.files_from, encoding="utf-8")
        with stream:
            paths.extend(Path(line.strip()) for line in stream if line.strip())
    if args.all:
        paths.extend(repository_files())
    if args.repo_rules or args.all:
        paths = [p for p in paths if in_scope(p)]
    seen: set[Path] = set()
    unique = []
    for p in paths:
        key = p.resolve()
        if key not in seen and p.is_file():
            seen.add(key)
            unique.append(p)
    return unique


def _state(report: Report) -> str:
    if report.gaps == 0:
        return "no comment text"
    if report.intact:
        return "intact (partial: no part is long enough for one full pass)" if report.partial else "intact"
    if report.moved:
        return f"intact, moved in {report.parts_moved} part" + ("" if report.parts_moved == 1 else "s")
    return "broken" if report.matched else "absent"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Comment spacing: consistent spacing between the words of comments and docstrings.")
    parser.add_argument("command", choices=("stamp", "check", "unstamp", "stamp-staged"))
    parser.add_argument("paths", nargs="*")
    parser.add_argument("--pattern", help="the gaps as a list of numbers (tests); otherwise the rule file is read")
    parser.add_argument("--files-from", help="a file with one path per line ('-' for standard input)")
    parser.add_argument("--repo-rules", action="store_true", help="only the files the repository rules apply to")
    parser.add_argument("--all", action="store_true", help="every file the repository rules apply to")
    parser.add_argument("--quiet", action="store_true", help="one summary line only")
    parser.add_argument("--json", action="store_true", help="check: one JSON object per file")
    args = parser.parse_args(argv)

    gaps_rule = load_gaps(args.pattern)
    if args.command != "unstamp" and not gaps_rule:
        print(f"comment spacing: no rule file ({rule_file_path()}), nothing done")
        return 0

    if args.command == "stamp-staged":
        root = Path(subprocess.run(["git", "rev-parse", "--show-toplevel"], capture_output=True, text=True).stdout.strip() or REPO_ROOT)
        print(stamp_staged(gaps_rule, root))
        return 0

    paths = _paths(args)
    if not paths:
        print("comment spacing: no files")
        return 0

    if args.command == "check":
        intact = 0
        for path in paths:
            report = check_file(path, gaps_rule)
            intact += report.intact
            if args.json:
                print(json.dumps({"file": str(path), **report.as_dict()}))
            elif report.error:
                print(f"{path}: {report.error}")
            else:
                moved = f"; moved runs start at {', '.join(f'gap {g} (entry {e})' for g, e in report.moved_at[:3])}" if report.moved_at else ""
                print(f"{path}: {_state(report)}; {report.repetitions} full repetitions; marks {report.matched} of "
                      f"{report.expected} expected over {report.gaps} gaps in {report.parts} parts, {report.extra} extra{moved}")
        if not args.json:
            print(f"comment spacing: {intact} of {len(paths)} files intact")
        return 0

    counts: dict[str, int] = {}
    for path in paths:
        outcome = stamp_file(path, gaps_rule) if args.command == "stamp" else unstamp_file(path)
        counts[outcome] = counts.get(outcome, 0) + 1
        if not args.quiet:
            print(f"{path}: {outcome}")
    summary = ", ".join(f"{n} {what}" for what, n in sorted(counts.items()))
    print(f"comment spacing: {summary}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
