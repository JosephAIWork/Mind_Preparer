"""Comment  spacing: keeps  the spacing between words inside comments and docstrings consistent
across the Python and TypeScript sources.

    python tools/comment_spacing.py stamp <paths...>      apply the spacing rule (idempotent)
    python tools/comment_spacing.py check <paths...>      report how far each file follows the rule
    python  tools/comment_spacing.py unstamp <paths...>    back to single spaces

The spacing rule lives outside the  repository, in the app's data folder  (watermark.json, a list
of numbers under "gaps"); without  that file the tool does nothing and says so. Tests give  the rule
with --pattern 1,2,3 instead. Only the  text of comments and docstrings is touched -- never code,
strings or line ends -- and only where one or two spaces stand between two words; everything
else in the file stays byte for byte (line endings, a byte order mark, tabs, trailing spaces).

--repo-rules keeps to the folders the rule applies to (app, scripts, tools, MindReadyUI/src) and
leaves tests, specs, fixtures, generated code and  third-party code alone; --all walks those
folders. The pre-commit hook in .githooks runs `stamp --repo-rules --files-from -` on the staged
files (tools/install_hooks.py switches it on for a clone).
"""
from __future__ import annotations

import argparse
import ast
import io
import json
import math
import os
import re
import sys
import tokenize
from dataclasses import dataclass
from pathlib import Path

RULE_FILE_NAME = "watermark.json"
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


# ----------------------------------------------------------------------------------------------
# The rule


def rule_file_path() -> Path:
    """Where the rule lives: the app's data folder (%LOCALAPPDATA%\\MindReady), never the repository."""
    base = os.environ.get("LOCALAPPDATA") or str(Path.home())
    return Path(base) / "MindReady" / RULE_FILE_NAME


def load_gaps(explicit: str | None = None) -> list[int] | None:
    """The list of gaps: from --pattern when given, else from the rule file; None when there is none."""
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


def mark_positions(gaps: list[int], count: int) -> list[int]:
    """The gap indexes (0-based, in file order) that carry a double space when a file has `count` gaps:
    a mark, then gaps[0] single spaces, a mark, gaps[1] single  spaces, ... repeating."""
    marks: list[int] = []
    position = 0
    index = 0
    while position < count:
        marks.append(position)
        position += gaps[index % len(gaps)] + 1
        index += 1
    return marks


# ----------------------------------------------------------------------------------------------
# Where the comment  text is


@dataclass
class Region:
    """A stretch of comment or docstring  text on one  line: [start,  end) as character columns."""

    line: int  # 0-based index into the file's lines
    start: int
    end: int


@dataclass
class Gap:
    """A run of one or two spaces between two words inside a Region."""

    line: int
    col: int
    width: int


def _split_lines(text: str) -> list[str]:
    """The file as lines that end with '\\n' (the last one  may not), joined back byte for byte."""
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
    """The region from `start` (after  the comment marker) to `end` (or  the line's end), with the
    leading spaces dropped;  None when nothing is left."""
    body = _line_body(lines[line_no])
    stop = len(body) if end is None else min(end, len(body))
    while start < stop and body[start] in " \t":
        start += 1
    return Region(line_no, start, stop) if start < stop else None


def _block_line_regions(lines: list[str], first_line: int, first_col: int, last_line: int, last_col: int,
                        skip_line=None) -> list[Region]:
    """Regions of a text block that spans  lines (a docstring, a /* */ comment):  the first line from
    first_col, the last line up to last_col, the lines between whole, each without its leading
    whitespace and a leading '*' (JSDoc style)."""
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


def python_regions(text: str, lines: list[str]) -> list[Region]:
    """Comment tokens and docstrings of a Python file (tokenize + ast); SyntaxError when it does not parse."""
    tree = ast.parse(text)
    docstring_spans: list[tuple[int, int, int, int]] = []
    for node in ast.walk(tree):
        if isinstance(node, (ast.Module, ast.ClassDef, ast.FunctionDef, ast.AsyncFunctionDef)):
            body = getattr(node, "body", None)
            if body and isinstance(body[0], ast.Expr) and isinstance(body[0].value, ast.Constant) \
                    and isinstance(body[0].value.value, str):
                value = body[0].value
                docstring_spans.append((value.lineno, _byte_to_char(lines, value.lineno, value.col_offset),
                                        value.end_lineno, _byte_to_char(lines, value.end_lineno, value.end_col_offset)))

    def in_docstring(row: int, col: int) -> bool:
        for l1, c1, l2, c2 in docstring_spans:
            if (l1, c1) <= (row, col) < (l2, c2):
                return True
        return False

    regions: list[Region] = []
    for tok in tokenize.generate_tokens(io.StringIO(text).readline):
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


def _byte_to_char(lines: list[str], row: int, byte_col: int) -> int:
    """ast gives byte columns (UTF-8); the lines are text."""
    line = lines[row - 1] if row - 1 < len(lines) else ""
    return len(line.encode("utf-8")[:byte_col].decode("utf-8", errors="ignore"))


_TS_WORD = re.compile(r"[A-Za-z_$][\w$]*")
_TS_NUMBER = re.compile(r"[0-9][\w.]*")


def typescript_comment_spans(text: str) -> list[tuple[int, int, str]]:
    """(start, end, 'line' | 'block') character spans of the comments of a TypeScript file.  A small
    scanner: strings, template literals (with ${ } inside), regular expression literals and
    comments; a '/' starts a regular expression when what comes before cannot be divided."""
    spans: list[tuple[int, int, str]] = []
    n = len(text)
    i = 0
    last = "punct"  # what the previous significant token was: 'value' (division follows) or 'punct'
    template_depths: list[int] = []  # for each open ${ }: how many braces are open inside it

    def scan_template(j: int) -> int | None:
        """From just after a '`' (or after the '}' that closes an expression): the index after the
        closing '`', or None when a '${' opened an expression (the expression is then scanned as code)."""
        while j < n:
            ch = text[j]
            if ch == "\\":
                j += 2
            elif ch == "`":
                return j + 1
            elif ch == "$" and j + 1 < n and text[j + 1] == "{":
                template_depths.append(0)
                return -(j + 2)
            else:
                j += 1
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
    return spans


def typescript_regions(text: str, lines: list[str]) -> list[Region]:
    """The text of  // comments, /* */ comments  and JSDoc of a TypeScript file."""
    starts = []
    offset = 0
    for line in lines:
        starts.append(offset)
        offset += len(line)

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


def regions_for(text: str, suffix: str) -> list[Region]:
    lines = _split_lines(text)
    if suffix == ".py":
        return python_regions(text, lines)
    if suffix in (".ts", ".tsx", ".js", ".mjs"):
        return typescript_regions(text, lines)
    return []


# ----------------------------------------------------------------------------------------------
# The gaps,  and the spacing  itself


def find_gaps(lines: list[str], regions: list[Region]) -> list[Gap]:
    """Runs of  one or two spaces with a non-space character on both sides, inside the regions, in file order."""
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


def _rewrite(lines: list[str], gaps: list[Gap], widths: list[int]) -> str:
    """The file with each gap set to its new width (1 or 2 spaces)."""
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


def stamp_text(text: str, gaps_rule: list[int], suffix: str) -> str:
    lines = _split_lines(text)
    gaps = find_gaps(lines, regions_for(text, suffix))
    marks = set(mark_positions(gaps_rule, len(gaps)))
    return _rewrite(lines, gaps, [2 if i in marks else 1 for i in range(len(gaps))])


def unstamp_text(text: str, suffix: str) -> str:
    lines = _split_lines(text)
    gaps = find_gaps(lines, regions_for(text, suffix))
    return _rewrite(lines, gaps, [1] * len(gaps))


@dataclass
class Report:
    """What `check`  finds in one file."""

    gaps: int = 0
    doubles: int = 0
    expected: int = 0
    matched: int = 0  # marks present from the start of the run,  up to the first missing one
    extra: int = 0  # double  spaces where the rule wants a single  one
    repetitions: int = 0  # full passes through the list of gaps
    intact: bool = False  # every expected mark present and  no extra double
    partial: bool = False  # the file is too short  for one full pass
    chance: float = 1.0  # the chance that the matched marks stand where they stand by accident
    chance_text: str = "1 in 1"
    error: str | None = None

    def as_dict(self) -> dict:
        return dict(self.__dict__)


def check_text(text: str, gaps_rule: list[int], suffix: str) -> Report:
    lines = _split_lines(text)
    gaps = find_gaps(lines, regions_for(text, suffix))
    doubles = {i for i, g in enumerate(gaps) if g.width == 2}
    expected = mark_positions(gaps_rule, len(gaps))
    matched = 0
    for position in expected:
        if position not in doubles:
            break
        matched += 1
    extra = len(doubles - set(expected))
    report = Report(gaps=len(gaps), doubles=len(doubles), expected=len(expected), matched=matched, extra=extra)
    report.repetitions = max(0, matched - 1) // len(gaps_rule)
    report.intact = matched == len(expected) and extra == 0 and len(expected) > 0
    report.partial = len(expected) < len(gaps_rule) + 1
    report.chance, report.chance_text = _chance(len(gaps), matched)
    return report


def _chance(count: int, matched: int) -> tuple[float, str]:
    """1 / C(count, matched): the chance that `matched` double spaces thrown at random over `count`
    gaps land exactly on the first `matched` positions of the run."""
    if matched <= 0 or count <= 0:
        return 1.0, "1 in 1"
    ways = math.comb(count, matched)
    digits = str(ways)
    if len(digits) > 300:
        text = f"1 in {digits[0]}.{digits[1:3]}e{len(digits) - 1}"
        return 0.0, text
    chance = 1.0 / ways
    return chance, f"1 in {ways:,}" if ways < 10_000_000 else f"1 in {ways:.2e}"


# ----------------------------------------------------------------------------------------------
# Files


def read_file(path: Path) -> tuple[str, bool] | None:
    """(text, had_bom); None when the file is not UTF-8."""
    data = path.read_bytes()
    bom = data.startswith(b"\xef\xbb\xbf")
    try:
        return data[3:].decode("utf-8") if bom else data.decode("utf-8"), bom
    except UnicodeDecodeError:
        return None


def write_file(path: Path, text: str, bom: bool) -> None:
    path.write_bytes((b"\xef\xbb\xbf" if bom else b"") + text.encode("utf-8"))


def stamp_file(path: Path, gaps_rule: list[int]) -> str:
    """'stamped', 'unchanged', or a reason the file was left alone."""
    loaded = read_file(path)
    if loaded is None:
        return "left alone: not UTF-8"
    text, bom = loaded
    try:
        new_text = stamp_text(text, gaps_rule, path.suffix.lower())
    except (SyntaxError, tokenize.TokenError, ValueError) as exc:
        return f"left alone: does not parse ({exc.__class__.__name__})"
    if new_text == text:
        return "unchanged"
    write_file(path, new_text, bom)
    return "stamped"


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
    """Whether the repository rules apply to  this file: a .py or .ts source under one of the allowed
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


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Comment spacing: consistent spacing between the words of comments and docstrings.")
    parser.add_argument("command", choices=("stamp", "check", "unstamp"))
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
                state = "intact" if report.intact else "broken" if report.matched else "absent"
                if report.intact and report.partial:
                    state = "intact (partial: the file is too short for one full pass)"
                print(f"{path}: {state}; {report.repetitions} full repetitions; marks {report.matched} of "
                      f"{report.expected} expected over {report.gaps} gaps, {report.extra} extra; "
                      f"chance by accident {report.chance_text} ({report.chance:.3g})")
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
