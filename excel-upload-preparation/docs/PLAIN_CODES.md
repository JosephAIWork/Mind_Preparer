# Plain English next to every code

Joseph asked: "whenever it says the error code it also translates the error to plain english".
Three kinds of code reach the user. Each one now comes with its plain words, from one place on
the server (`app/plain_codes.py`); the screens get the same words from it.

## 1. Rule codes (FRM-009, INP-010 ...)

Every rule has a short plain name in `rules/plain_names.yaml`: what is wrong, in the user's
words, at most 8 words. It was written from each rule's description **and** its validator.

    FRM-009: "an empty cell is compared with a number"
    REF-002: "a formula reads another workbook"
    INP-010: "macros bring in data from other files"

- Rules that only list what the workbook holds (they always pass) say "list of ...";
  advice that never blocks says "tip: ...".
- **Every rule needs one**, draft rules included: `tests/unit/test_rules_engine.py` fails on a
  rule without a name, on a name longer than 8 words, and on a name for a rule that does not exist.
- Where it shows:
  - The server adds `rule_name` to every finding, to the readiness lists (`blocking`, the next
    step, what an Apply did) and to the "since the previous analysis" lists; each Prep action has
    `rule_names` in the order of its `rule_ids`.
  - A text that cites a code says it with its name in brackets: "Fix FRM-002 (a function Mind may
    not support) with the assistant"; "Apply the blocking repairs in Prep: REF-001 (a broken
    reference to a cell or name), STR-008 (a formula reads a lone text cell)"; the scan's progress
    line "Checking rule 34/109: FRM-009 (an empty cell is compared with a number)" (`cite`).
  - The screens show "FRM-009 · an empty cell is compared with a number", the code small and grey
    (`MindReadyUI/src/app/core/code-name.ts`, `<mr-code>`). A chip in the workbook map shows the
    code and has the name as its tooltip.
  - The Excel report has a column "In plain words" next to "Rule ID".
  - The assistant sees each finding with its name and is asked to name codes the same way. Its
    context names each rule once (at its first mention), and the names come on top of the
    context's size limit: they never push out the end of it (the last findings, the Prep actions).
- An id that is not a rule ("UI" for the screens' own changes, "RECALC") shows alone.

## 2. Excel's error values (#REF!, #N/A ...)

One glossary, `ERROR_VALUES`, read after "means":

| Value | Means |
|---|---|
| #REF! | the formula points at a cell that no longer exists |
| #N/A | the value is not available (often: a lookup found nothing) |
| #VALUE! | a value is the wrong kind, such as text in a sum |
| #NAME? | Excel does not know a function or a name in the formula |
| #DIV/0! | a division by zero or by an empty cell |
| #NUM! | a number Excel cannot compute (impossible, or too big) |
| #NULL! | two ranges in the formula do not meet |
| #SPILL! | the result needs more cells, and some of them are not empty |
| #CALC! | Excel could not compute an array (an empty result, say) |
| #GETTING_DATA | Excel is still getting the value |
| #FIELD! | the field asked for is not in the linked data |
| #BLOCKED! | Excel blocked something the formula needs |
| #UNKNOWN! | a data type this Excel does not know |
| #CONNECT! | Excel could not reach a data source |
| #BUSY! | Excel is still working on the value |

The meaning is given **once**, never after every occurrence:

- under a finding's message that quotes error values (Findings table, the "By hand" list in Prep,
  the Excel report's Message column);
- once above the Recalculate panel's error list (and as each error cell's tooltip);
- once under the Model map's cell window (and in each error cell's tooltip, under its formula; a
  '#Title' cell, which the window marks too, has no meaning);
- once under the numbers-check samples (Deliver);
- once after the cells "Fix all issues" left for a person (main screen).

The screens get the glossary from `GET /api/plain-codes` (`{"rules": {...}, "errors": {...}}`),
loaded once at start (`MindReadyUI/src/app/core/plain-codes.ts`): they keep no list of their own.

## 3. Excel / VBA / COM errors

`plain_com_error` turns the error Excel, VBA or Windows hands on into one plain sentence and keeps
Excel's own words in brackets:

    (-2147352567, 'Exception occurred.', (0, None, 'Member not found.', None, 0, -2146788248), None)
    -> Excel does not have the command the app asked for (Excel said: "Member not found.")

| Excel's words (or code) | Plain sentence |
|---|---|
| Member not found | Excel does not have the command the app asked for |
| Subscript out of range | the macro looked for a sheet, workbook or list item that is not there |
| Type mismatch | a value was of the wrong kind, such as text where a number was expected |
| Object required | the macro used a sheet, range or workbook that is not there |
| Application-defined or object-defined error | Excel refused the action (a protected sheet, a merged cell or a value it does not accept, for example) |
| The RPC server is unavailable | Excel closed or stopped answering while the app was using it |
| Call was rejected by callee / application is busy | Excel was busy (a message box may be open in Excel) and refused the request: try again |
| File not found / couldn't find | a file that is needed is not there (moved, renamed or deleted?) |
| Path not found | a folder that is needed is not there |
| Permission denied / Access is denied | Windows refused access to a file (open in another program, or read-only?) |
| Exception occurred (and nothing more) | Excel stopped with an error and did not say why |

- Excel's words are matched first, then the error's code -- so a Windows or an Office in another
  language (a Hebrew "RPC server is unavailable") is translated too.
- Excel's catch-all code 0x800A03EC says only that something failed: when Excel gave specific words
  with it ("Cannot change part of an array.") those are kept as they are.
- An error this does not know comes back as Excel's own words, never as pywin32's raw tuple.
- Where it is used: the import macro's error (it replaces `_vba_error_text`), the reason Excel
  refused a change (Prep, Fix all issues), the plain sentence of the feedback box, and the error
  text of a failed scan, Apply or background job.
- And where a message reaches the screens inside a normal answer, which the server's error handler
  never sees (`plain_error`, where the message is made): a recalculation that could not run (its
  prefix "Excel COM recalculation failed" stays: the server reads it), a numbers check that could not
  compare, an .xlsb conversion Excel refused, and the report's warnings (Deliver).
- An error no step expected gets "Excel stopped with an error: " in front only when its sentence does
  not already say what Excel did -- never "Excel stopped with an error: Excel stopped with an error ...".
- The screens show the reason of a refused change whole (a plain sentence and Excel's words, often
  160 characters): up to 200 characters, cut after a whole word; changes refused for the same reason
  share it ("S!B5, S!B6: ...").
- The message to the Tel Aviv team (`app/feedback.py`) keeps its codes as they are: the decoder
  reads them.

## Adding a rule

Give it a line in `rules/plain_names.yaml` (the test says so if you forget). Nothing else: every
screen, text and report picks the name up from there.
