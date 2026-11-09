"""Plain English next to every code (app/plain_codes.py, docs/PLAIN_CODES.md): the rules' plain
names, Excel's error values, and Excel / VBA / COM errors said in a plain sentence."""
from app.plain_codes import (
    ERROR_VALUES,
    cite,
    cite_all,
    error_meaning,
    error_meanings,
    error_values_in,
    glossary,
    is_com_error,
    named,
    plain_com_error,
    plain_com_errors_in,
    rule_name,
)
from app.rules_engine import RulesEngine
from app.validators import run_rule

# pywin32's own text for the errors Excel and VBA raise (str(com_error)); the source is None or a name
MEMBER_NOT_FOUND = "(-2147352567, 'Exception occurred.', (0, None, 'Member not found.', None, 0, -2146788248), None)"
SUBSCRIPT = "(-2147352567, 'Exception occurred.', (0, 'Microsoft Visual Basic', 'Subscript out of range', None, 0, -2146828279), None)"


# --- rule ids ----------------------------------------------------------------------------------
def test_a_rule_is_cited_with_its_plain_name_and_an_unknown_id_alone():
    assert rule_name("FRM-009") == "an empty cell is compared with a number"
    assert cite("FRM-009") == "FRM-009 (an empty cell is compared with a number)"
    assert cite_all(["REF-002", "STR-008"]) == "REF-002 (a formula reads another workbook), STR-008 (a formula reads a lone text cell)"
    # the screens' own changes ("UI"), a recalculation error ("RECALC"): no rule, no name
    assert rule_name("UI") is None and cite("UI") == "UI" and cite(None) == ""
    assert named({"rule_id": "INP-010", "status": "ERROR"}) == {"rule_id": "INP-010", "status": "ERROR", "rule_name": "macros bring in data from other files"}


def test_every_finding_carries_its_rule_name():
    engine = RulesEngine()
    rule = engine.get("READY-001")
    finding = run_rule(rule, {"source": {}, "workbooks": [], "features": {}}, {})
    assert finding["rule_id"] == "READY-001" and finding["rule_name"] == "Excel recalculation not run or not clean"


# --- Excel's error values ----------------------------------------------------------------------
def test_the_glossary_has_every_error_value_excel_shows():
    for value in ("#REF!", "#N/A", "#VALUE!", "#NAME?", "#DIV/0!", "#NUM!", "#NULL!", "#SPILL!", "#CALC!", "#GETTING_DATA",
                  "#FIELD!", "#BLOCKED!", "#UNKNOWN!", "#CONNECT!", "#BUSY!"):
        assert error_meaning(value), value
    assert error_meaning("#REF!") == "the formula points at a cell that no longer exists"
    assert error_meaning(" #div/0! ") == ERROR_VALUES["#DIV/0!"]  # as a user may type it
    assert error_meaning("#Assumptions") is None and error_meaning(None) is None  # a grid title is no error value
    # short and plain: read after "means", one line each
    assert all(len(m.split()) <= 14 and not m.endswith(".") for m in ERROR_VALUES.values())


def test_a_text_gets_the_meaning_of_each_error_value_once():
    text = "Sheet!B2 shows #REF! and C3 #N/A; D4 #REF! again, read by =IFERROR(B2,0)"
    assert error_values_in(text) == ["#REF!", "#N/A"]
    assert error_meanings(text) == (
        "#REF! means the formula points at a cell that no longer exists; #N/A means the value is not available (often: a lookup found nothing)"
    )
    assert error_meanings(["#div/0!", "a value", "#DIV/0!"]) == "#DIV/0! means a division by zero or by an empty cell"
    assert error_meanings("no error value here") == "" and error_values_in(None) == []


def test_the_screens_get_the_same_words():
    g = glossary()
    assert g["errors"] == ERROR_VALUES and g["rules"]["FRM-009"] == rule_name("FRM-009")
    assert len(g["rules"]) == len(RulesEngine().rules)


# --- Excel / VBA / COM errors ------------------------------------------------------------------
def test_the_common_com_and_vba_errors_become_plain_sentences_with_excels_words_in_brackets():
    assert plain_com_error(MEMBER_NOT_FOUND) == 'Excel does not have the command the app asked for (Excel said: "Member not found.")'
    assert plain_com_error(SUBSCRIPT) == 'the macro looked for a sheet, workbook or list item that is not there (Excel said: "Subscript out of range")'
    vba = "(-2147352567, 'Exception occurred.', (0, 'Microsoft Visual Basic', '{}', None, 0, {}), None)"
    assert plain_com_error(vba.format("Type mismatch", -2146828275)).startswith("a value was of the wrong kind, such as text where a number was expected (Excel said:")
    assert plain_com_error(vba.format("Object required", -2146827864)).startswith("the macro used a sheet, range or workbook that is not there")
    assert plain_com_error(vba.format("Application-defined or object-defined error", -2146827284)).startswith("Excel refused the action (a protected sheet")
    assert plain_com_error(vba.format("File not found", -2146828235)).startswith("a file that is needed is not there")
    assert plain_com_error(vba.format("Path not found", -2146828212)).startswith("a folder that is needed is not there")
    assert plain_com_error(vba.format("Permission denied", -2146828218)).startswith("Windows refused access to a file")
    assert plain_com_error("(-2147023174, 'The RPC server is unavailable.', None, None)") == (
        'Excel closed or stopped answering while the app was using it (Excel said: "The RPC server is unavailable.")'
    )
    assert plain_com_error("(-2147418111, 'Call was rejected by callee.', None, None)").startswith("Excel was busy (a message box may be open in Excel)")
    # Excel's words alone (a refused write, a macro's message) are said the same way
    assert plain_com_error("Subscript out of range") == 'the macro looked for a sheet, workbook or list item that is not there (Excel said: "Subscript out of range")'


def test_the_code_says_it_when_the_words_are_in_another_language_or_missing():
    # a Windows in Hebrew words the RPC error in Hebrew: its code says the same
    assert plain_com_error("(-2147023174, 'שרת ה-RPC אינו זמין.', None, None)").startswith("Excel closed or stopped answering")
    # VBA's code without its words: the code inside wins over pywin32's "Exception occurred."
    assert plain_com_error("(-2147352567, 'Exception occurred.', (0, None, None, None, 0, -2146828279), None)") == (
        "the macro looked for a sheet, workbook or list item that is not there"
    )
    assert plain_com_error("(-2147352567, 'Exception occurred.', None, None)") == "Excel stopped with an error and did not say why"


def test_excels_own_specific_words_and_unknown_errors_stay_as_they_are():
    # 0x800A03EC is Excel's catch-all: its own words say more than any sentence
    assert plain_com_error("(-2147352567, 'Exception occurred.', (0, 'Microsoft Excel', 'Cannot change part of an array.', 'xlmain11.chm', 0, -2146827284), None)") == (
        "Cannot change part of an array."
    )
    assert plain_com_error("(-2147467259, 'Unspecified error', None, None)") == "Unspecified error"  # never the raw tuple
    assert plain_com_error("the import macro ran longer than 900 s and was stopped") == "the import macro ran longer than 900 s and was stopped"
    assert plain_com_error("") == "" and plain_com_error(None) == ""
    # a text already said plainly is not said again around itself
    once = plain_com_error(SUBSCRIPT)
    assert plain_com_error(once) == once


def test_a_message_cut_short_is_read_as_far_as_it_goes():
    assert plain_com_error(MEMBER_NOT_FOUND[:-12]) == 'Excel does not have the command the app asked for (Excel said: "Member not found.")'
    cut = SUBSCRIPT[: SUBSCRIPT.index("range")]  # "... 'Microsoft Visual Basic', 'Subscript out of "
    assert plain_com_error(cut) == "Subscript out of"  # Excel's words as far as they go, never its source


def test_a_com_error_object_and_a_longer_text_holding_one():
    class ComError(Exception):  # pywin32's com_error, as Excel raises it
        def __init__(self, hresult, message, excepinfo):
            super().__init__(hresult, message, excepinfo, None)
            self.hresult, self.strerror, self.excepinfo, self.argerror = hresult, message, excepinfo, None

    err = ComError(-2147352567, "Exception occurred.", (0, "Microsoft Visual Basic", "Subscript out of range", None, 0, -2146828279))
    assert is_com_error(err) and not is_com_error(ValueError("x"))
    assert plain_com_error(err) == plain_com_error(SUBSCRIPT)
    assert plain_com_error(ValueError("merged range B5:D5")) == "merged range B5:D5"
    assert plain_com_errors_in(f"analysis failed: {MEMBER_NOT_FOUND}") == (
        'analysis failed: Excel does not have the command the app asked for (Excel said: "Member not found.")'
    )
    assert plain_com_errors_in("no COM error here") == "no COM error here"
    assert plain_com_errors_in("KeyError: (3, 'Sheet1')") == "KeyError: (3, 'Sheet1')"  # a tuple that is no COM error


def test_excels_refusals_of_a_write_are_said_plainly():
    """Prep's and the fixer's "Excel did not accept this change" reasons (app/prep.py _com_message,
    app/autofix.py _com_text) go through the same translator; the app's own refusals stay as they are."""
    from app.autofix import _com_text
    from app.prep import _com_message

    class Refused(Exception):
        excepinfo = (0, "Microsoft Excel", "Application-defined or object-defined error", None, 0, -2146827284)

    assert _com_message(Refused()).startswith("Excel refused the action (a protected sheet")
    assert _com_text(Refused()).startswith("Excel refused the action (a protected sheet")
    assert _com_message(ValueError("S!A2: inside the merged range A2:C2")) == "S!A2: inside the merged range A2:C2"
