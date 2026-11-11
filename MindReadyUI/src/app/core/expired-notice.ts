import { Component, Input } from "@angular/core";
import { noticeParts } from "./expiry";

/** The  whole page  once this copy stopped working (its end date passed, app/expiry.py): the server's
 *  message, calmly, instead of the app. No button: nothing is left to do here but ask for a new copy. */
@Component({
  selector: "mr-expired-notice",
  standalone: true,
  templateUrl: "./expired-notice.html",
  styleUrl: "./expired-notice.scss",
})
export class ExpiredNotice {
  @Input() message = "";

  parts(): { title: string; text: string } {
    return noticeParts(this.message);
  }
}
