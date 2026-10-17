"""Compare two scan_fingerprint.py outputs: every difference by path, then the seconds of each.

    python scripts/compare_fingerprints.py <before.json> <after.json>     (exit 1 when they differ)
"""
from __future__ import annotations

import json
import sys

IGNORED = {"seconds", "report_id"}


def diff(x, y, path, out):
    if type(x) is not type(y):
        out.append((path, str(x)[:160], str(y)[:160]))
    elif isinstance(x, dict):
        for k in sorted(set(x) | set(y)):
            if k in IGNORED:
                continue
            if k not in x or k not in y:
                out.append((f"{path}/{k}", str(x.get(k))[:160], str(y.get(k))[:160]))
            else:
                diff(x[k], y[k], f"{path}/{k}", out)
    elif isinstance(x, list):
        if len(x) != len(y):
            out.append((f"{path} (length)", len(x), len(y)))
        for i, (p, q) in enumerate(zip(x, y)):
            diff(p, q, f"{path}[{i}]", out)
    elif x != y:
        out.append((path, str(x)[:160], str(y)[:160]))
    return out


def main() -> int:
    a = json.load(open(sys.argv[1], encoding="utf-8"))
    b = json.load(open(sys.argv[2], encoding="utf-8"))
    d = diff(a, b, "", [])
    print(f"seconds: {a.get('seconds')} -> {b.get('seconds')}; {len(d)} difference(s)")
    for row in d[:40]:
        print(ascii(row))
    return 1 if d else 0


if __name__ == "__main__":
    sys.exit(main())
