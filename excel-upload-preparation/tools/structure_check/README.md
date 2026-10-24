# Structure check (prototype, not wired into the app)

Reports where a Mind loop or Mind instances would fit in a workbook. It checks; it
does not choose. Design, results and the next step: `docs/LOOP_INSTANCE_CHECK.md`.

Nothing in `app/` imports this folder; it imports the app's R1C1 parser read-only.

    python build_control.py                          # control workbook with known answers -> control.xlsx
    python extract.py control=<full path>\control.xlsx tag2=<full path to a workbook>
                                                     # Excel reads FormulaR1C1 -> data/<tag>.pkl
    python analyze.py control tag2                   # -> data/<tag>.result.pkl
    python report.py control tag2                    # -> data/<tag>.report.txt

Set PYTHONIOENCODING=utf-8 for workbooks with non-ASCII sheet names.
`data/` holds client formulas: it is git-ignored and must stay out of the repo.
