import { TestBed } from "@angular/core/testing";
import { ChangeStack } from "./change-stack";
import { Operation, Version } from "./models";
import { SessionStore } from "./session.store";

const version = (id: string, n: number): Version => ({ id, major: n, minor: 0, label: `v${n} — test`, file_name: `v${n}.xlsx`, source: "test" });
const set = (cell: string, value: unknown, extra: Partial<Operation> = {}): Operation => ({ op: "set_value", sheet: "Sheet1", cell, value, ...extra });

describe("ChangeStack", () => {
  let store: SessionStore;
  let stack: ChangeStack;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    store = TestBed.inject(SessionStore);
    stack = TestBed.inject(ChangeStack);
    store.sessionId.set("s1");
    store.currentVersion.set(version("v2", 2));
  });

  it("keeps the last copy of a repeated change, so a value chosen again wins", () => {
    stack.add("assistant", "A", [set("B2", 0.03)]);
    stack.add("assistant", "B", [set("B2", 0.035)]);
    stack.add("assistant", "C", [set("B2", 0.03)]);
    expect(stack.operations().map((o) => o["value"])).toEqual([0.035, 0.03]);
  });

  it("stamps the version a proposal was made on, so an older one is reported", () => {
    stack.add("assistant", "A", [set("B12", 1)], version("v1", 1));
    expect(stack.warnings().some((w) => w.kind === "stale")).toBe(true);
    stack.add("prep", "B", [set("C3", 1)]);
    expect(stack.entries()[1].versionId).toBe("v2");
  });

  it("removes only the applied entries", () => {
    stack.add("prep", "A", [set("A1", 1)]);
    const keys = new Set(stack.entries().map((e) => e.key));
    stack.add("assistant", "B", [set("A2", 1)]);
    stack.removeKeys(keys);
    expect(stack.entries().map((e) => e.label)).toEqual(["B"]);
  });

  it("does not compare a cell planned after the inserts with one planned before them", () => {
    stack.add("prep", "Titles", [{ op: "insert_row", sheet: "Sheet1", row: 5 }, set("A5", "Title", { after_inserts: true, insert_at: 5 })]);
    stack.add("prep", "Header", [set("A5", "Rate")]);
    expect(stack.warnings().filter((w) => w.kind === "overlap")).toEqual([]);
  });

  it("says a cell write wins over a pivot table whatever the order", () => {
    stack.add("assistant", "Constant", [set("B3", 7)]);
    stack.add("prep", "Pivot", [{ op: "pivot_to_formulas", sheet: "Sheet1", range: "A1:C10" }]);
    const w = stack.warnings().filter((x) => x.kind === "overlap");
    expect(w.length).toBe(1);
    expect(w[0].text).toContain('so "Constant" is written over it');
  });

  it("warns when Prep and the assistant both insert on one sheet", () => {
    stack.add("prep", "Titles", [{ op: "insert_row", sheet: "Sheet1", row: 6 }, set("D6", "Title", { after_inserts: true, insert_at: 6 })]);
    stack.add("assistant", "Column", [{ op: "insert_column", sheet: "Sheet1", column: "B" }]);
    expect(stack.warnings().some((w) => w.kind === "overlap" && w.text.includes("both insert rows or columns on Sheet1"))).toBe(true);
  });
});
