import { TestBed } from "@angular/core/testing";
import { ApiService } from "./api.service";
import { PlainCodesGlossary } from "./models";
import { PlainCodes, errorLegend, errorValuesIn, meaningsText, ruleLabel } from "./plain-codes";

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
      await codes.load();
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
