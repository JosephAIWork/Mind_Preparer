"""Stage 3: turn the analysis into a check report. For each repetition: what is there, whether the
copies are parallel or a chain, and for each option (loop / instances / leave as is) whether it
fits and why. It never picks one."""
from __future__ import annotations

import pickle
import sys
from pathlib import Path

SMALL_GAIN = 200        # formula cells: below this a change buys little
CHAIN_SHARE = 0.10      # share of a copy's formulas reading an earlier copy that makes it a chain


def gain(f):
    return (f["copies"] - 1) * f["per_copy"]


def assess(f):
    """f: copies, per_copy, chain, kinds, numbers_differ, size_differs, adding, whole_sheets."""
    k = f["kinds"]
    logic = k.get("different logic", 0)
    driver = k.get("reads a different input", 0)
    typed = k.get("differs only in a typed-in number", 0)
    lines = []
    if f["chain"]:
        lines.append("CHAIN: each copy reads an earlier one (a sequence such as periods). Not a dimension: "
                     "neither a loop nor instances express it; leave as is.")
        return lines
    g = gain(f)
    loop_need, loop_block, inst_need, inst_block = [], [], [], []
    if driver:
        loop_need.append(f"{driver} cells differ only in which input they read: they become the loop index")
    if typed:
        loop_need.append(f"{typed} cells differ only in a typed-in number: move those numbers into a table")
    if f["numbers_differ"]:
        loop_need.append(f"{f['numbers_differ']} assumption numbers differ between copies: one table row per copy")
    if f["adding"]:
        loop_need.append(f"{f['adding']} cells add the copies up: they become MM_RESULT(cell)")
    if logic:
        loop_block.append(f"{logic} cells have different logic per copy: a loop needs IF(MM_DIMINDEX(...)) for each")
    if f["size_differs"]:
        loop_block.append(f"{f['size_differs']} cells exist in some copies only (different sizes): needs MM_LOOPINSTANCE or instances")
    if f["whole_sheets"]:
        inst_need.append("instances copy a whole workbook: these sheets would move into a workbook of their own")
    else:
        inst_need.append("instances copy a whole workbook: this block would move into a workbook of its own")
    if logic:
        inst_need.append(f"the {logic} cells of different logic stay as per-copy overrides")
    if f["numbers_differ"]:
        inst_need.append(f"the {f['numbers_differ']} differing assumptions stay per copy")
    if f["adding"]:
        inst_need.append(f"the {f['adding']} adding cells become a plain link (Mind adds the copies)")
    inst_block.append("Mind: instances use more RAM and CPU than loops")
    lines.append(f"PARALLEL: {f['copies']} copies x {f['per_copy']} formulas; a single copy would remove {g:,} formula cells"
                 + (" (small)" if g < SMALL_GAIN else ""))
    lines.append("  Loop      " + ("does not fit as is" if loop_block else "fits") + (": " + "; ".join(loop_block) if loop_block else ""))
    for x in loop_need:
        lines.append("            - " + x)
    lines.append("  Instances fits" + ("" if not inst_need else ""))
    for x in inst_need + inst_block:
        lines.append("            - " + x)
    lines.append("  Leave     fits" + (" (little to gain)" if g < SMALL_GAIN else f" (keeps {g:,} formula cells to maintain)"))
    return lines


def main(tag):
    r = pickle.load(open(Path(__file__).parent / "data" / f"{tag}.result.pkl", "rb"))
    out = [f"=== {tag}: {r['sheets']} sheets, {r['formulas']:,} formula cells",
           f"Mind constructs already there: {r['mind'] or 'none'}"]
    n = 0
    for g in r["sheet_copies"]:
        n += 1
        f = {"copies": len(g["members"]), "per_copy": max(g["per_copy_formulas"]),
             "chain": bool(g["copies_read_each_other"]), "kinds": g["diff_kinds"],
             "numbers_differ": g["numbers_differ"], "size_differs": g["cells_in_some_copies_only"],
             "adding": g["adding_cells"], "whole_sheets": True}
        out.append(f"\n[{n}] Sheet copies: {', '.join(g['members'])}  (similarity {g['min_similarity']})")
        out.append(f"    differences: {g['diff_kinds']}; numbers differ {g['numbers_differ']}"
                   + (f" in rows {g['differing_numbers_area'][0]}-{g['differing_numbers_area'][2]}" if g['differing_numbers_area'] else ""))
        if g["copies_read_each_other"]:
            out.append(f"    copies read: {g['copies_read_each_other']}")
        if g["adding_example"]:
            out.append(f"    adding cell e.g. {g['adding_example'][0]}: {g['adding_example'][1]}")
        out += ["    " + x for x in assess(f)]
    axes = {"time axis": 0, "independent": []}
    for b in r["inside"]:
        chain = b["chain_share"] >= CHAIN_SHARE
        if b.get("fill"):
            if chain or b["header"] == "numbers":
                axes["time axis"] += 1
            elif b["copies"] * b["formulas_per_copy"] >= 500:
                axes["independent"].append(f"{b['sheet']}!{b['where']} ({b['copies']} {b['axis']} x {b['formulas_per_copy']}, header {b['header']})")
            continue
        n += 1
        f = {"copies": b["copies"], "per_copy": b["formulas_per_copy"], "chain": chain, "kinds": b["diff_kinds"],
             "numbers_differ": b.get("numbers_differ", 0), "size_differs": 0, "adding": b["adding_cells"], "whole_sheets": False}
        out.append(f"\n[{n}] Repeated block in {b['sheet']}!{b['where']}: {b['copies']} copies, {b['width']} {b['axis']} wide"
                   f" (copies read earlier copies in {b['chain_share']:.0%} of sampled cells; header {b['header']})")
        out.append(f"    differences: {b['diff_kinds']}; numbers differ {b.get('numbers_differ', 0)}")
        out += ["    " + x for x in assess(f)]
    out.append(f"\nFilled-across/down areas that read their own earlier columns/rows (time axes): {axes['time axis']} -- not candidates")
    if axes["independent"]:
        out.append("Large filled areas whose columns/rows do not read each other (a list of items, or periods without recurrence):")
        out += ["  - " + x for x in axes["independent"][:15]]
        out.append("  Each is already one formula repeated; a loop would replace the repetition only if the columns/rows are items, which the formulas alone cannot tell.")
    if r.get("switches"):
        out.append("\nSwitch cells: typed-in cells that steer many lookups -- the model computes one choice at a time:")
        for w in r["switches"]:
            out.append(f"  - {w['cell']}{' = name ' + w['name'] if w.get('name') else ''} ({w['kind']}{', label at ' + w['label_cell'] if w['label_cell'] else ''}):"
                       f" read by {w['formula_cells']:,} formula cells via {w['via']}")
        out.append("  For each: a loop over its choices fits if every choice uses the same formulas (all results in one run);"
                   " instances fit if a choice needs its own assumptions or formulas; leave fits if one choice per run is what is wanted."
                   " Which choices exist (a dropdown list, a table column, a macro) is not in the formulas -- the app's model map reads dropdowns and macros.")
    if n == 0:
        out.append("\nNo parallel copies found: nothing here for a loop or instances to replace.")
    text = "\n".join(out)
    (Path(__file__).parent / "data" / f"{tag}.report.txt").write_text(text, encoding="utf-8")
    return text


if __name__ == "__main__":
    for tag in sys.argv[1:]:
        print(main(tag))
        print()
