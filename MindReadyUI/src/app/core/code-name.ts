import { Component, Input, inject } from "@angular/core";
import { PlainCodes } from "./plain-codes";

/** A  code in  plain words (16 Nov 2026, Joseph: no technical codes on screen, they stay in the bug
 *  reports): a rule shows its words only, "an empty cell is compared with a number"; an  Excel error
 *  value shows its words with the value in brackets, "the  formula points at a cell that  no longer
 *  exists (#REF!)", since that is what  the user sees in Excel. The code alone when it has  no words. */
@Component({
  selector: "mr-code",
  standalone: true,
  template: `@if (words(); as w) {<span class="words" [attr.title]="null">{{ w }}</span>@if (isExcelError()) {<span class="code"> ({{ code }})</span>}} @else {<span class="code">{{ code }}</span>}`,
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

  /** #N/A, #REF!, #VALUE! ...: Excel's own error values, which the user sees  in the workbook */
  isExcelError(): boolean {
    return this.code.trim().startsWith("#");
  }
}
