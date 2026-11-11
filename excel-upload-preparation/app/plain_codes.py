"""Plain  English next  to every code the app shows (docs/PLAIN_CODES.md).

Three kinds of code reach the user, and each gets its plain words here, in one place:

* **Rule ids** (FRM-009, INP-010 ...): one short name per  rule, from rules/plain_names.yaml, read
  once. The server adds it to every finding  (`rule_name`), to the readiness lists and  to the Prep
  actions (`rule_names`); a text that  cites a rule says "FRM-009 (an empty cell is compared with  a
  number)" (`cite`).
* **Excel's error values** (#REF!, #N/A ...):  one glossary, `ERROR_VALUES`. The screens get the same
  meanings from GET /api/plain-codes; a text that quotes error values gets the meaning of each once
  (`error_meanings`), never after every occurrence.
* **Excel / VBA / COM errors** handed on to the user ("Subscript out of range", or pywin32's
  "(-2147352567, 'Exception occurred.', (0, None, 'Member not found.', None, 0, -2146788248), None)"):
  `plain_com_error` says the common ones in one plain  sentence and keeps Excel's own words in brackets.

Nothing here reads or changes a workbook: it only puts words next to codes.
"""
from __future__ import annotations

import ast
import re
from functools import lru_cache
from pathlib import Path
from typing import Any, Iterable

import yaml

PLAIN_NAMES_PATH = Path(__file__).resolve().parent.parent / "rules" / "plain_names.yaml"
MAX_NAME_WORDS = 8  # a  name is  read at a glance, next to the code


# --- rule ids ----------------------------------------------------------------------------------
@lru_cache(maxsize=1)
def rule_names() -> dict[str, str]:
    """rule  id ->  its plain name, read once from rules/plain_names.yaml. Without the file the codes
    show alone: nothing else depends on the names."""
    if not PLAIN_NAMES_PATH.exists():
        return {}
    data = yaml.safe_load(PLAIN_NAMES_PATH.read_text(encoding="utf-8")) or {}
    return {str(k).strip(): str(v).strip() for k, v in data.items() if v}


def rule_name(rule_id: Any) -> str | None:
    """The  plain name  of a rule id, None for an id that is not a rule ("UI", "CHAT", "RECALC")."""
    return rule_names().get(str(rule_id or "").strip())


def cite(rule_id: Any) -> str:
    """A  rule id  as a text says it: 'FRM-009 (an empty cell is compared with a number)'; the id
    alone when it has no name."""
    rid = str(rule_id or "").strip()
    name = rule_name(rid)
    return f"{rid} ({name})" if name else rid


def cite_all(rule_ids: Iterable[Any], sep: str = ", ") -> str:
    return sep.join(cite(r) for r in rule_ids)


def named(item: dict[str, Any]) -> dict[str, Any]:
    """`item`  with `rule_name`  beside its `rule_id` (changed in place, and returned)."""
    item["rule_name"] = rule_name(item.get("rule_id"))
    return item


# --- Excel's error values ----------------------------------------------------------------------
# Each meaning reads after "means": "#REF! means the formula points at a cell that no longer exists".
# The screens show the  same words (GET /api/plain-codes): keep them short and plain.
ERROR_VALUES: dict[str, str] = {
    "#REF!": "the formula points at a cell that no longer exists",
    "#N/A": "the value is not available (often: a lookup found nothing)",
    "#VALUE!": "a value is the wrong kind, such as text in a sum",
    "#NAME?": "Excel does not know a function or a name in the formula",
    "#DIV/0!": "a division by zero or by an empty cell",
    "#NUM!": "a number Excel cannot compute (impossible, or too big)",
    "#NULL!": "two ranges in the formula do not meet",
    "#SPILL!": "the result needs more cells, and some of them are not empty",
    "#CALC!": "Excel could not compute an array (an empty result, say)",
    "#GETTING_DATA": "Excel is still getting the value",
    "#FIELD!": "the field asked for is not in the linked data",
    "#BLOCKED!": "Excel blocked something the formula needs",
    "#UNKNOWN!": "a data type this Excel does not know",
    "#CONNECT!": "Excel could not reach a data source",
    "#BUSY!": "Excel is still working on the value",
}


def error_meaning(value: Any) -> str | None:
    """'#REF!'  -> what  it means; None for anything that is not one of Excel's error values."""
    return ERROR_VALUES.get(str(value or "").strip().upper())


def error_values_in(text: Any) -> list[str]:
    """The  error values  a text quotes, each once, in the order they first appear."""
    s = str(text or "")
    return [v for _, v in sorted((s.find(v), v) for v in ERROR_VALUES if v in s)]


def error_meanings(values_or_text: str | Iterable[Any]) -> str:
    """'#REF!  means the  formula points at a cell that no longer exists; #N/A means ...' for the error
    values given -- or quoted in a text -- each once; '' when there is none."""
    if isinstance(values_or_text, str):
        values = error_values_in(values_or_text)
    else:
        values = list(dict.fromkeys(str(v).strip().upper() for v in values_or_text if error_meaning(v)))
    return "; ".join(f"{v} means {ERROR_VALUES[v]}" for v in values)


# --- Excel  / VBA / COM errors ------------------------------------------------------------------
def _hresult(code: int) -> int:
    """An  HRESULT as  pywin32 hands it over: a signed 32-bit integer."""
    return code - (1 << 32) if code >= 1 << 31 else code


def _vba(number: int) -> int:
    """The  code a  VBA run-time error reaches Python with: 0x800A0000 + its number (9: Subscript out of range)."""
    return _hresult(0x800A0000 | number)


# (Excel's words, lower case: any of them; the codes that mean the same in any language --  a Windows
# in Hebrew says "The RPC server is unavailable" in Hebrew;  the plain sentence). Every place that  uses
# a sentence puts it after a colon  ("Excel did not accept it: ..."), so it starts in lower  case unless
# its first word is Excel.
COM_ERRORS: list[tuple[tuple[str, ...], tuple[int, ...], str]] = [
    (("member not found",), (_hresult(0x80020003),), "Excel does not have the command the app asked for"),
    (("subscript out of range",), (_vba(9),), "the macro looked for a sheet, workbook or list item that is not there"),
    (("type mismatch",), (_vba(13), _hresult(0x80020005)), "a value was of the wrong kind, such as text where a number was expected"),
    (("object required",), (_vba(424),), "the macro used a sheet, range or workbook that is not there"),
    (("application-defined or object-defined error",), (), "Excel refused the action (a protected sheet, a merged cell or a value it does not accept, for example)"),
    (("rpc server is unavailable",), (_hresult(0x800706BA),), "Excel closed or stopped answering while the app was using it"),
    (("call was rejected by callee", "application is busy"), (_hresult(0x80010001), _hresult(0x8001010A)), "Excel was busy (a message box may be open in Excel) and refused the request: try again"),
    (("disconnected from its clients",), (_hresult(0x80010108),), "Excel closed while the app was using it"),
    (("server execution failed",), (_hresult(0x80080005),), "Excel could not be started"),
    (("file not found", "cannot find the file", "couldn't find"), (_vba(53), _hresult(0x80070002)), "a file that is needed is not there (moved, renamed or deleted?)"),
    (("path not found", "cannot find the path"), (_vba(76), _hresult(0x80070003)), "a folder that is needed is not there"),
    (("permission denied", "access is denied"), (_vba(70), _hresult(0x80070005)), "Windows refused access to a file (open in another program, or read-only?)"),
]
# Excel's  catch-all codes  say only that something failed: the words Excel gave with them ("Cannot change
# part of an array.") say more than any sentence, so these apply only when it gave none.
GENERIC_CODES: dict[int, str] = {
    _vba(1004): "Excel refused the action (a protected sheet, a merged cell or a value it does not accept, for example)",  # 0x800A03EC
    _hresult(0x80020009): "Excel stopped with an error and did not say why",  # DISP_E_EXCEPTION,  "Exception occurred."
}
# pywin32's  wrapper around  Excel's own error: last of all, so the code inside it (a VBA "Subscript out
# of range" that came without its words) says more first
GENERIC_PHRASES: dict[str, str] = {"exception occurred": "Excel stopped with an error and did not say why"}

# a  failure HRESULT  is negative and long (-2147352567): "(3, 'Sheet1')" in another error's text is no COM error
_COM_START_RE = re.compile(r"\(\s*(-\d{6,})\s*,\s*['\"]")
_STR = r"'(?:[^'\\]|\\.)*'|\"(?:[^\"\\]|\\.)*\""
_MESSAGE_RE = re.compile(rf"\(\s*-?\d+\s*,\s*({_STR})")
_EXCEPINFO_RE = re.compile(rf"\(\s*-?\d+\s*,\s*(?:None|{_STR})\s*,\s*({_STR})")  # (wcode,  source, description
_CODE_RE = re.compile(r"-?\d{6,}")


def _literal(text: str) -> Any:
    try:
        return ast.literal_eval(text)
    except (ValueError, SyntaxError, TypeError, MemoryError, RecursionError):
        return None


def _com_tuple(text: str) -> tuple[int, int, Any] | None:
    """(start,  end, the  tuple) of the first pywin32 error written in a text, e.g.
    "(-2147352567, 'Exception occurred.', (0, None, 'Member not found.', None, 0, -2146788248), None)".
    A text cut before the tuple closes (a message kept  to 300 characters) is closed and read as far as
    it goes;  the tuple is None when even  that cannot be read."""
    m = _COM_START_RE.search(text)
    if not m:
        return None
    depth, quote, i = 0, None, m.start()
    while i < len(text):
        ch = text[i]
        if quote:
            if ch == "\\":
                i += 2
                continue
            if ch == quote:
                quote = None
        elif ch in "'\"":
            quote = ch
        elif ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
            if depth == 0:
                return m.start(), i + 1, _literal(text[m.start():i + 1])
        i += 1
    return m.start(), len(text), _literal(text[m.start():] + (quote or "") + ")" * max(depth, 0))


def _parts(hresult: Any, message: Any, info: Any) -> tuple[int | None, str | None, int | None, str | None]:
    """(the  HRESULT, its  message, Excel's own code, Excel's own words) from pywin32's pieces;
    `info` is its excepinfo: (wcode, source, description, helpfile, helpcontext, scode)."""
    info = info if isinstance(info, tuple) else ()
    scode = info[5] if len(info) > 5 and isinstance(info[5], int) else None
    words = str(info[2]).strip() if len(info) > 2 and info[2] else None
    return (hresult if isinstance(hresult, int) else None), (str(message).strip() if message else None), scode, words


def is_com_error(error: Any) -> bool:
    """A  pywin32 com_error,  or an error that carries one (excepinfo / hresult)."""
    return isinstance(error, BaseException) and (getattr(error, "excepinfo", None) is not None or isinstance(getattr(error, "hresult", None), int))


def _com_parts(error: Any) -> tuple[int | None, str | None, int | None, str | None] | None:
    """The  pieces of  the COM error in `error` -- a pywin32 com_error, or a text that holds one -- or None."""
    if isinstance(error, BaseException):
        if is_com_error(error):
            args = tuple(getattr(error, "args", ()) or ())
            hresult = getattr(error, "hresult", None)
            message = getattr(error, "strerror", None)
            if hresult is None and len(args) >= 2 and isinstance(args[0], int):
                hresult, message = args[0], args[1]
            return _parts(hresult, message, getattr(error, "excepinfo", None))
        return _com_parts(str(error))
    text = str(error or "")
    found = _com_tuple(text)
    if found is None:
        return None
    start, end, parsed = found
    if isinstance(parsed, tuple) and len(parsed) >= 2:
        return _parts(parsed[0], parsed[1], parsed[2] if len(parsed) > 2 else None)
    # not readable as a whole: its pieces one by one -- the HRESULT, its  message, Excel's words (the
    # string after the source in the inner tuple)  and Excel's code (the last long  number)
    fragment = text[start:end]
    head = _COM_START_RE.match(fragment)
    message = _MESSAGE_RE.match(fragment)
    inner = _EXCEPINFO_RE.search(fragment, message.end() if message else 0)
    codes = [int(c) for c in _CODE_RE.findall(fragment)]
    words = _literal(inner.group(1)) if inner else None
    return (
        int(head.group(1)) if head else None,
        _literal(message.group(1)) if message else None,
        codes[-1] if len(codes) > 1 else None,
        str(words).strip() if words else None,
    )


def _sentence(words: str | None, codes: Iterable[int | None], excel_words: bool) -> str | None:
    """The  plain sentence  for an error: by Excel's words first, then by its codes, Excel's own code
    before the wrapper's (they mean the same whatever the language); Excel's catch-all codes only when
    it gave no words  of its own; pywin32's "Exception occurred." last."""
    low = (words or "").lower()
    for phrases, _, sentence in COM_ERRORS:
        if low and any(p in low for p in phrases):
            return sentence
    codes = [c for c in codes if isinstance(c, int)]
    for c in codes:
        for _, ids, sentence in COM_ERRORS:
            if c in ids:
                return sentence
    if not excel_words:
        for c in codes:
            if c in GENERIC_CODES:
                return GENERIC_CODES[c]
    return next((sentence for p, sentence in GENERIC_PHRASES.items() if p in low), None)


def plain_com_error(error: Any) -> str:
    """An  Excel /  VBA / COM error in one plain sentence, Excel's own words kept in brackets:
    'the macro looked for a sheet, workbook or list item that is not there (Excel said: "Subscript
    out of  range")'. `error` is a pywin32 com_error, a text holding one (only  the error is kept, as
    the old  _vba_error_text did), or Excel's words themselves. An  error this does not know comes back
    as Excel's own words (never  the raw tuple); a text with no  error in it comes back as it is."""
    parts = _com_parts(error)
    if parts is None:
        text = re.sub(r"\s+", " ", str(error or "")).strip()
        # a text already said plainly (it quotes Excel's words) is not said again around itself
        sentence = _sentence(text, (), True) if text and "(Excel said: " not in text else None
        return f'{sentence} (Excel said: "{text}")' if sentence else text
    hresult, message, scode, words = parts
    said = words or message
    sentence = _sentence(said, (scode, hresult), bool(words))
    if sentence is None:
        return said or re.sub(r"\s+", " ", str(error or "")).strip()
    # pywin32's own "Exception occurred." is not Excel's words: nothing worth quoting
    wrapper = not words and any(p in (said or "").lower() for p in GENERIC_PHRASES)
    return f'{sentence} (Excel said: "{said}")' if said and not wrapper else sentence


def plain_com_errors_in(text: Any) -> str:
    """A  longer text  with every pywin32 error tuple in it said in plain words, the rest kept:
    "analysis failed: (-2147352567, ...)" -> 'analysis failed: Excel does not have the command the app
    asked for (Excel said:  "Member not found.")'."""
    rest, out = str(text or ""), []
    while True:
        found = _com_tuple(rest)
        if found is None:
            break
        start, end, _ = found
        out.append(rest[:start])
        out.append(plain_com_error(rest[start:end]))
        rest = rest[end:]
    out.append(rest)
    return "".join(out)


def plain_error(error: BaseException, with_type: bool = False) -> str:
    """An  error the  app caught, in words for a screen: an Excel / VBA error (pywin32's com_error) as a
    plain sentence with Excel's words in brackets; any other error as its text, with every pywin32  error
    tuple it quotes said plainly. `with_type` puts the type of a  non-Excel error in front ("KeyError:
    'Sheet1'"): for  an error no step expected, whose text  alone can say nothing.

    The messages that reach a screen inside a  normal answer (a recalculation that could not  run, the
    numbers check, the report's warnings) need this where they are made: the server's error handler,
    which says the errors of a failed request plainly, never sees them."""
    if is_com_error(error):
        return plain_com_error(error)
    return plain_com_errors_in(f"{type(error).__name__}: {error}" if with_type else str(error))


def glossary() -> dict[str, Any]:
    """What  GET /api/plain-codes  sends the screens: every rule's plain name and Excel's error values."""
    return {"rules": dict(rule_names()), "errors": dict(ERROR_VALUES)}
