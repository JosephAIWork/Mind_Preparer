/** The  backend's plan  texts with their jargon said plainly, for the Mind loop and the Instances tabs alike
 *  (the loop plan's `why` of a loop, the instances plan's loop options and notes). "count cell" ->  "a cell
 *  for the number of passes", "the driver" -> "the cell  the macro changes", "MM_LOOP needs" ->  "Mind
 *  needs", "needs a decision" -> "needs one  answer", "Mind loop panel" -> "Mind loop tab". */
export function plainWhy(why: string): string {
  return why
    .replace(/\s*\(count cell\)/gi, "")
    .replace(/\bMM_LOOP needs a cell that holds the count\b/g, "Mind needs a cell that holds the number of passes")
    .replace(/\bthe count cell\b/gi, "the cell for the number of passes")
    .replace(/\bcount cell\b/gi, "a cell for the number of passes")
    .replace(/\bthe driver cell\b/gi, "the cell the macro changes")
    .replace(/\b(the|its) driver\b/gi, (_m: string, w: string) => `${w[0] === w[0].toUpperCase() ? "The" : "the"} cell the macro changes`)
    .replace(/\bneeds a decision\b/gi, "needs one answer")
    .replace(/\bneed a decision\b/gi, "need one answer")
    .replace(/\bMind loop panel\b/g, "Mind loop tab")
    .replace(/\s{2,}/g, " ")
    .trim();
}

/** The  one plan  note that must be read before the build: Prep still has blocking items that are not
 *  formulas (a lone text cell, a pivot table), which Mind's converter refuses or computes wrong. */
const PREP_NOTE = /^Prep still has \d+ blocking items? that (?:is|are) not (?:a )?formulas?\b/;

/** A  plan's notes  in two parts: Prep's blockers (shown above Build it) and the rest (behind "Notes"). */
export function splitNotes(notes: string[] | null | undefined): { prep: string | null; rest: string[] } {
  const all = notes ?? [];
  const prep = all.find((n) => PREP_NOTE.test(n)) ?? null;
  return { prep, rest: all.filter((n) => n !== prep) };
}
