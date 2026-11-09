import { Injectable, inject, signal } from "@angular/core";
import { ApiService } from "./api.service";
import { PlainCodesGlossary } from "./models";

/** Plain English next to every code the screens show (excel-upload-preparation/docs/PLAIN_CODES.md).
 *
 *  The words come from the server, GET /api/plain-codes: the same glossary the server's texts and the
 *  Excel report use (rules/plain_names.yaml, app/plain_codes.py), so the screens keep no list of their
 *  own. Until it has loaded -- or when it cannot -- a code shows alone and nothing else changes. */

/** "FRM-009 · an empty cell is compared with a number"; the code alone when it has no name. */
export function ruleLabel(code: string, name?: string | null): string {
  return name ? `${code} · ${name}` : code;
}

/** The error values a text quotes, each once, in the order they first appear. */
export function errorValuesIn(text: string | null | undefined, errors: Record<string, string>): string[] {
  const t = text ?? "";
  return Object.keys(errors)
    .map((value) => [t.indexOf(value), value] as const)
    .filter(([at]) => at >= 0)
    .sort((a, b) => a[0] - b[0])
    .map(([, value]) => value);
}

export interface ErrorMeaning {
  value: string;
  meaning: string;
}

/** What each error value means, once per value, for the values that are Excel's error values
 *  (anything else -- a number, a text -- is left out). */
export function errorLegend(values: Iterable<string | null | undefined>, errors: Record<string, string>): ErrorMeaning[] {
  const out: ErrorMeaning[] = [];
  const seen = new Set<string>();
  for (const raw of values) {
    const value = String(raw ?? "").trim().toUpperCase();
    const meaning = errors[value];
    if (meaning && !seen.has(value)) {
      seen.add(value);
      out.push({ value, meaning });
    }
  }
  return out;
}

/** "#REF! means the formula points at a cell that no longer exists; #N/A means ..." ("" when none). */
export function meaningsText(legend: ErrorMeaning[]): string {
  return legend.map((m) => `${m.value} means ${m.meaning}`).join("; ");
}

/** The glossary as signals: the app shell loads it once at start (`load`); every screen asks it. */
@Injectable({ providedIn: "root" })
export class PlainCodes {
  private api = inject(ApiService);
  readonly glossary = signal<PlainCodesGlossary>({ rules: {}, errors: {} });
  private loading: Promise<void> | null = null;

  /** Fetches the glossary once; a failed fetch can be tried again. */
  load(): Promise<void> {
    this.loading ??= this.api
      .plainCodes()
      .then((g) => this.glossary.set({ rules: g?.rules ?? {}, errors: g?.errors ?? {} }))
      .catch(() => {
        this.loading = null;
      });
    return this.loading;
  }

  /** A code's plain name: the one the server sent with it, else the glossary's -- a rule's name or what
   *  an error value means (null: a code with no words, such as the screens' own "UI"). */
  name(code: string | null | undefined, given?: string | null): string | null {
    if (given) return given;
    if (!code) return null;
    const g = this.glossary();
    return g.rules[code] ?? g.errors[code.trim().toUpperCase()] ?? null;
  }

  /** "FRM-009 · an empty cell is compared with a number", for a tooltip or a line of text. */
  label(code: string, given?: string | null): string {
    return ruleLabel(code, this.name(code, given));
  }

  /** What one error value means (null: not one of Excel's error values). */
  meaning(value: string | null | undefined): string | null {
    return errorLegend([value], this.glossary().errors)[0]?.meaning ?? null;
  }

  /** The meaning of each error value among `values`, once each. */
  legend(values: Iterable<string | null | undefined>): ErrorMeaning[] {
    return errorLegend(values, this.glossary().errors);
  }

  /** The error values a text quotes, each once. */
  valuesIn(text: string | null | undefined): string[] {
    return errorValuesIn(text, this.glossary().errors);
  }

  /** The meanings of the error values a text quotes, each once: shown under a message, never after
   *  every occurrence. */
  meaningsIn(text: string | null | undefined): ErrorMeaning[] {
    return this.legend(this.valuesIn(text));
  }
}
