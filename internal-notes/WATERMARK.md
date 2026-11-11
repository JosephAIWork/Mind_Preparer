# Watermark in the code (11-12 November 2026)

Joseph's work is watermarked with spaces: `excel-upload-preparation/tools/comment_spacing.py`
("comment spacing" in the shipped repository, on purpose -- the shipped tool, its docs and the
changelog never use the word watermark) walks the word gaps of the comment and docstring text of a
file in order (Python `#` comments and docstrings; TypeScript `//`, `/* */` and JSDoc) and writes a
double space at the gaps his pattern names: a mark, then n1 normal gaps, a mark, n2 normal gaps, ...
repeating. **The pattern is not in the repository, not in this note, and not in git's history from
commit 4e0ac69 on** (see "Before the push" below): the tool reads it from
`%LOCALAPPDATA%\MindReady\comment_spacing.json` (`{"gaps": [...]}`, a list of twelve numbers on
Joseph's laptop only); without that file it does nothing and says so.

The run starts again at every line that begins at column 0 with code (a top-level def, class,
import, assignment; a TypeScript import, decorator, class or closing brace): one part per top-level
block. So an edit inside one function re-lays the marks of that function only, and a commit's diff
stays within what was edited (12 November: before this, every edit shifted every later mark and a
commit rewrote comment lines across the whole file). A copied function carries its own complete run
from its first gap.

- `stamp` is idempotent (a file already carrying the mark is unchanged; an edited part is stamped
  again from its start -- every one-or-two-space gap in comment text is set fresh, so prose double
  spaces such as `a  ->  b` become single).
- `check` reports per file: intact (every part's run where the rule lays it, no extra double),
  "intact, moved in N parts" (a run is complete but starts at another gap or another entry of the
  list -- a word added before it, or the head of the file cut off; it says where each moved run
  starts), broken (some marks, not the whole run) or absent; full repetitions over all parts; marks
  found of expected over the gaps; extra doubles. It prints no probability (12 November: the shipped
  tool gives nothing away). The chance that the marks stand where they stand by accident is
  1 / C(gaps, marks): `%LOCALAPPDATA%\MindReady\watermark_odds.py <files...>` prints it per file
  (outside the repository; it imports the tool from the repo).
- `unstamp` goes back to single spaces (a reviewer needs it to read a diff; it is shipped, and a
  reader who understands the tool could run it -- the mark proves authorship of what was copied
  before that, not after).
- `stamp-staged` is what the pre-commit hook runs: it reads each staged `.py`/`.ts` the rules apply
  to from the index, stamps that copy and puts it back (`git hash-object -w` + `git update-index
  --cacheinfo`), and stamps the working-tree file only when it still equals the staged copy. An
  unstaged edit is never swept into the commit (12 November: the first hook ran `git add` on the
  working-tree file and did exactly that). A partial commit (`git commit -- <paths>`, run by git on a
  temporary index it then throws away) is left alone with a line saying so; the next full commit of
  the file stamps it.

Code, strings, template literals, regular expressions, line ends, tabs, trailing spaces, CRLF/LF and
a byte order mark are never touched; tool directives (`# noqa`, `# type:`, `# fmt:`, `// eslint`,
`/// <reference>`...) and doctest lines are left alone. The rules apply under
`excel-upload-preparation/app`, `scripts`, `tools` and `MindReadyUI/src` (tests, specs, fixtures,
dist, node_modules, generated and third-party files skipped). `tools/install_hooks.py` sets
`core.hooksPath = .githooks` for a clone; the hook never blocks a commit. The tests
(`tests/unit/test_comment_spacing.py`) give the tool their own pattern (1, 2, 3). The built screens
(`dist`) are minified, so the mark lives in the sources only.

## Before the push

The first version of this note (commit 4e0ac69, 11 November, not pushed as of 12 November 02:00)
printed the pattern. The note was corrected in the review-fixes commit of 12 November, but the old
text stays in the history of branch `iteration-1-14-0` until that history is rewritten. Joseph
decides before pushing: either rewrite the branch from 4e0ac69 (an interactive rebase editing that
one commit, or `git filter-repo --replace-text` with the twelve numbers), or accept that the pattern
is in the Milliman repository's history (internal-notes/ ships with every push; only the `git archive`
hand-over to the Mind team excludes it).
