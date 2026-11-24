import { Component, input } from "@angular/core";

/** The  assistant is  thinking: three dots that pulse one after the other (16 Nov 2026, Joseph: "some sort of
 *  motion which represents that the model is thinking", wherever the AI works; a neutral motion, no  logo).
 *  It moves gently and stops for people who ask their system  for less motion. */
@Component({
  selector: "mr-thinking",
  standalone: true,
  template: `<span class="thinking" role="status" [attr.aria-label]="label()"><span class="thinking__dot"></span><span class="thinking__dot"></span><span class="thinking__dot"></span>@if (text()) {<span class="thinking__text">{{ text() }}</span>}</span>`,
  styles: [
    `
      :host {
        display: inline-block;
        vertical-align: middle;
      }
      .thinking {
        display: inline-flex;
        align-items: center;
        gap: 0.3rem;
      }
      .thinking__dot {
        width: 0.5rem;
        height: 0.5rem;
        border-radius: 50%;
        background: var(--cds-interactive, #0f62fe);
        opacity: 0.25;
        animation: mr-thinking-pulse 1.2s ease-in-out infinite;
      }
      .thinking__dot:nth-child(2) {
        animation-delay: 0.2s;
      }
      .thinking__dot:nth-child(3) {
        animation-delay: 0.4s;
      }
      .thinking__text {
        margin-left: 0.4rem;
        color: var(--cds-text-secondary, #525252);
        font-size: 0.875rem;
      }
      @keyframes mr-thinking-pulse {
        0%,
        80%,
        100% {
          opacity: 0.25;
          transform: scale(0.8);
        }
        40% {
          opacity: 1;
          transform: scale(1);
        }
      }
      @media (prefers-reduced-motion: reduce) {
        .thinking__dot {
          animation: none;
          opacity: 0.6;
        }
      }
    `,
  ],
})
export class Thinking {
  /** a  short line  beside the dots ("The assistant is thinking…"); none by default */
  readonly text = input("");
  /** what a screen reader says */
  readonly label = input("The assistant is thinking");
}
