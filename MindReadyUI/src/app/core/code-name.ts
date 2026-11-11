import { Component, Input, inject } from "@angular/core";
import { PlainCodes } from "./plain-codes";

/** A  code with  its plain words, the code small and grey: "FRM-009 · an empty cell is compared with a
 *  number", or "#REF! · the formula points at a cell that no longer exists". The  one way every screen
 *  shows a rule id or an Excel error  value; the code alone when it  has no words. */
@Component({
  selector: "mr-code",
  standalone: true,
  template: `<span class="code">{{ code }}</span>@if (words(); as w) {<span class="sep"> · </span><span class="words">{{ w }}</span>}`,
  styleUrl: "./code-name.scss",
})
export class CodeName {
  private codes = inject(PlainCodes);
  /** a  rule id  (FRM-009) or an Excel error value (#REF!) */
  @Input({ required: true }) code = "";
  /** the words the server sent with the code (a finding's `rule_name`); the glossary's otherwise */
  @Input() name: string | null | undefined = null;

  words(): string | null {
    return this.codes.name(this.code, this.name);
  }
}
