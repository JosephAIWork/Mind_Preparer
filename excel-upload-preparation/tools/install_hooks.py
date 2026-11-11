"""Switches  the repository's  git hooks on for this clone: `git config core.hooksPath .githooks`.

    python tools/install_hooks.py [--root <git root>]

The hooks live in <git root>/.githooks (the pre-commit hook keeps the comment spacing of the staged
Python and TypeScript  sources consistent, see tools/comment_spacing.py). Git does not run hooks from a
folder  of the repository by itself; this  one-line setting, per clone, makes it do  so.
"""
from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path

HOOKS_FOLDER = ".githooks"


def install(root: Path) -> str:
    """Sets  core.hooksPath for  the clone at `root`; returns one line to print."""
    hooks = root / HOOKS_FOLDER
    if not (hooks / "pre-commit").is_file():
        return f"no hooks folder at {hooks}, nothing done"
    result = subprocess.run(["git", "config", "core.hooksPath", HOOKS_FOLDER], cwd=str(root),
                            capture_output=True, text=True)
    if result.returncode != 0:
        return f"git config failed: {(result.stderr or result.stdout).strip()}"
    return f"git hooks on for {root}: core.hooksPath = {HOOKS_FOLDER}"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Switch the repository's git hooks on for this clone.")
    parser.add_argument("--root", help="the git root (default: two folders above tools/)")
    args = parser.parse_args(argv)
    root = Path(args.root).resolve() if args.root else Path(__file__).resolve().parents[2]
    print(install(root))
    return 0


if __name__ == "__main__":
    sys.exit(main())
