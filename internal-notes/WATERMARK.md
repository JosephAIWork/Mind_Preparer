# Watermark in the code (11 November 2026)

Joseph's work is watermarked with spaces: `excel-upload-preparation/tools/comment_spacing.py`
("comment spacing" in the shipped repository, on purpose) walks the word gaps of the comment and
docstring text of a file in order (Python `#` comments and docstrings; TypeScript `//`, `/* */` and
JSDoc) and writes a double space at the gaps his pattern names -- a mark, then 1 normal gap, a mark,
30 normal gaps, a mark, 10, 5, 6, 10, 6, 60, 80, 4, 6, 2 normal gaps, repeating until the file's
comment text runs out. The pattern itself is not in the repository: the tool reads it from
`%LOCALAPPDATA%\MindReady\watermark.json` (`{"gaps": [1,30,10,5,6,10,6,60,80,4,6,2]}`); without that
file it does nothing and says so. `stamp` is idempotent (a file already carrying the mark is
unchanged; a file edited since is re-stamped from its start -- every one-or-two-space gap in comment
text is set fresh, so prose double spaces such as `a  ->  b` become single), `check` reports per file
how many full repetitions are present, whether the run is intact and the chance that the double
spaces stand where they stand by accident (`1 / C(gaps, marks)`), `unstamp` goes back to single
spaces. Code, strings, template literals, regular expressions, line ends, tabs, trailing spaces,
CRLF/LF and a byte order mark are never touched; tool directives (`# noqa`, `# type:`, `# fmt:`,
`// eslint`, `/// <reference>`...) and doctest lines are left alone. The pre-commit hook
`.githooks/pre-commit` stamps every staged `.py`/`.ts` under `excel-upload-preparation/app`,
`scripts`, `tools` and `MindReadyUI/src` (tests, specs, fixtures, dist, node_modules, generated and
third-party files skipped) and re-adds them, never blocking a commit; `tools/install_hooks.py` sets
`core.hooksPath = .githooks` for a clone. The shipped docs say nothing about this; the tests
(`tests/unit/test_comment_spacing.py`) give the tool their own pattern. The built screens (`dist`) are
minified, so the mark lives in the sources only.
