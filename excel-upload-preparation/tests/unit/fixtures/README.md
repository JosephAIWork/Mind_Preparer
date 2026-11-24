# Fixtures

- `vba_japanese_cp932.xlsm` (15 KB): a small row-per-record model (the loop builder's shape: Policy Data,
  Calculation, OUTPUT) whose VBA module holds Japanese comments and string literals stored in cp932
  with the project's code page recorded as 0, as a workbook saved on a Japanese PC has. Made once on
  10 November 2026: written through Excel (DispatchEx + VBProject) with a placeholder module, then
  the module stream's compressed source replaced at the same length by the cp932 bytes (MS-OVBA
  literal tokens; the p-code cache before the text offset kept). For olevba / model_map.read_vba;
  Excel itself is not expected to run its macro. Read by tests/unit/test_model_map_vba_text.py.
