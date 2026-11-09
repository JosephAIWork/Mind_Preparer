import { TestBed } from "@angular/core/testing";
import { ApiService } from "./api.service";
import { ApplyResult, PlainCodesGlossary, WindowCell, clipAtWord, failedText } from "./models";
import { PlainCodes, cellTip, errorCellValues, errorLegend, errorValuesIn, meaningsText, ruleLabel } from "./plain-codes";

// the server's own words (GET /api/plain-codes), a few of them
const errors = {
  "#REF!": "the formula points at a cell that no longer exists",
  "#N/A": "the value is not available (often: a lookup found nothing)",
  "#DIV/0!": "a division by zero or by an empty cell",
};
const rules = { "FRM-009": "an empty cell is compared with a number", "REF-002": "a formula reads another workbook" };

describe("plain codes", () => {
  it("shows a code with its plain name, or the code alone when it has none", () => {
    expect(ruleLabel("FRM-009", rules["FRM-009"])).toBe("FRM-009 · an empty cell is compared with a number");
    expect(ruleLabel("UI", null)).toBe("UI");
  });

  it("finds the error values a text quotes, each once, in the order they appear", () => {
    expect(errorValuesIn("B2 shows #N/A, C3 shows #REF! and D4 #N/A again", errors)).toEqual(["#N/A", "#REF!"]);
    expect(errorValuesIn("#Assumptions is a grid title, not an error", errors)).toEqual([]);
    expect(errorValuesIn(null, errors)).toEqual([]);
  });

  it("gives what each error value means once, and leaves out what is not an error value", () => {
    expect(errorLegend(["#REF!", "12.5", "#ref!", null, "#DIV/0!"], errors)).toEqual([
      { value: "#REF!", meaning: errors["#REF!"] },
      { value: "#DIV/0!", meaning: errors["#DIV/0!"] },
    ]);
    expect(meaningsText(errorLegend(["#REF!", "#N/A"], errors))).toBe(
      "#REF! means the formula points at a cell that no longer exists; #N/A means the value is not available (often: a lookup found nothing)",
    );
    expect(meaningsText([])).toBe("");
  });

  describe("the Model map's cell window", () => {
    // as the server sends it (GET .../cells): every value that starts with '#' is marked, a grid title too
    const cell = (ref: string, value: unknown, formula: string | null = null): WindowCell => ({
      ref, value, formula, array: null, error: typeof value === "string" && value.startsWith("#"), focus: false,
    });
    const cells = [cell("A1", "#Rates /Input"), cell("B1", 1.5), cell("C1", "#N/A", "=VLOOKUP(A1,T,2,0)"), cell("A2", "#REF!", "=#REF!+1"), cell("B2", "#REF!", "=A2*2")];

    it("says what each error value in the window means, once each, in the order shown; a '#Title' has no meaning", () => {
      expect(errorCellValues(cells)).toEqual(["#Rates /Input", "#N/A", "#REF!", "#REF!"]);
      expect(errorLegend(errorCellValues(cells), errors)).toEqual([
        { value: "#N/A", meaning: errors["#N/A"] },
        { value: "#REF!", meaning: errors["#REF!"] },
      ]);
    });

    it("puts the meaning in an error cell's tooltip, under its formula", () => {
      expect(cellTip("=A2*2", "#REF!", errors)).toBe("=A2*2\n#REF! means the formula points at a cell that no longer exists");
      expect(cellTip(null, "#N/A", errors)).toBe("#N/A means the value is not available (often: a lookup found nothing)");
      expect(cellTip("=B1*2", 3, errors)).toBe("=B1*2"); // a value: the formula alone
      expect(cellTip(null, "#Rates /Input", errors)).toBe(""); // a grid title: nothing to explain
      expect(cellTip(null, null, errors)).toBe("");
    });
  });

  describe("the reason Excel refused a change", () => {
    // what the server sends now: a plain sentence with Excel's own words in brackets (app/plain_codes.py)
    const refused = 'Excel refused the action (a protected sheet, a merged cell or a value it does not accept, for example) (Excel said: "Application-defined or object-defined error")';
    const missing = 'Excel does not have the command the app asked for (Excel said: "Member not found.")';
    const result = (...failed: [string, string][]): ApplyResult => ({
      status: "PARTIAL",
      applied: 1,
      failed: failed.map(([cell, error]) => ({ op: "set_formula", sheet: "S", cell, error })),
    });

    it("is shown whole, Excel's own words included", () => {
      expect(failedText(result(["B5", refused]))).toBe(`S!B5: ${refused}`);
      expect(failedText(result(["B5", missing]))).toBe(`S!B5: ${missing}`);
    });

    it("is said once for the changes refused the same way, and the count of the rest stays", () => {
      const text = failedText(result(["B5", refused], ["B6", refused], ["C9", missing], ["D1", refused]));
      expect(text).toBe(`S!B5, S!B6: ${refused}; S!C9: ${missing}; … 1 more`);
    });

    it("is cut after a whole word when it is very long, never in the middle of one", () => {
      const long = `${refused} ${"and then some more words ".repeat(6)}`.trim();
      const shown = failedText(result(["B5", long])).slice("S!B5: ".length);
      expect(shown.endsWith("…")).toBe(true);
      expect(shown.length).toBeLessThanOrEqual(201);
      expect(long.startsWith(shown.slice(0, -1))).toBe(true);
      expect(long[shown.length - 1]).toBe(" "); // the cut falls between two words
      expect(clipAtWord("short", 200)).toBe("short");
    });
  });

  describe("PlainCodes", () => {
    let calls = 0;
    let answer: () => Promise<PlainCodesGlossary> = () => Promise.resolve({ rules, errors });
    let codes: PlainCodes;

    beforeEach(() => {
      calls = 0;
      answer = () => Promise.resolve({ rules, errors });
      TestBed.configureTestingModule({
        providers: [{ provide: ApiService, useValue: { plainCodes: () => (calls++, answer()) } }],
      });
      codes = TestBed.inject(PlainCodes);
    });

    it("shows codes alone until the glossary has loaded, then with their words", async () => {
      expect(codes.label("FRM-009")).toBe("FRM-009");
      expect(codes.cellTip("=A2*2", "#REF!")).toBe("=A2*2");
      await codes.load();
      expect(codes.cellTip("=A2*2", "#REF!")).toBe(`=A2*2\n#REF! means ${errors["#REF!"]}`);
      expect(codes.cellsLegend([{ value: "#REF!", error: true }, { value: 2, error: false }, { value: "#ref!", error: true }])).toEqual([
        { value: "#REF!", meaning: errors["#REF!"] },
      ]);
      expect(codes.label("FRM-009")).toBe("FRM-009 · an empty cell is compared with a number");
      expect(codes.name("#ref!")).toBe(errors["#REF!"]); // an error value has words too
      expect(codes.name("UI")).toBeNull();
      expect(codes.meaning("#DIV/0!")).toBe(errors["#DIV/0!"]);
      expect(codes.meaning("42")).toBeNull();
      expect(codes.meaningsIn("A1 =#REF!+1 and A2 =#REF!*2").map((m) => m.value)).toEqual(["#REF!"]);
    });

    it("prefers the name the server sent with the code", async () => {
      await codes.load();
      expect(codes.name("FRM-009", "sent with the finding")).toBe("sent with the finding");
      expect(codes.name("FRM-009", null)).toBe(rules["FRM-009"]);
    });

    it("asks the server once, and again only after a failure", async () => {
      await Promise.all([codes.load(), codes.load()]);
      expect(calls).toBe(1);

      TestBed.resetTestingModule();
      answer = () => Promise.reject(new Error("offline"));
      TestBed.configureTestingModule({ providers: [{ provide: ApiService, useValue: { plainCodes: () => (calls++, answer()) } }] });
      const other = TestBed.inject(PlainCodes);
      calls = 0;
      await other.load();
      expect(other.label("FRM-009")).toBe("FRM-009"); // nothing else changes: the code shows alone
      answer = () => Promise.resolve({ rules, errors });
      await other.load();
      expect(calls).toBe(2);
      expect(other.label("REF-002")).toBe("REF-002 · a formula reads another workbook");
    });
  });
});
