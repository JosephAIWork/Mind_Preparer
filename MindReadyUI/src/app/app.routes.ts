import { Routes } from "@angular/router";
import { AdvancedPage } from "./pages/advanced.page";
import { HomePage } from "./pages/home.page";

/** 1.15.0: three screens. Basic ("") and Custom ("custom") are the same page in two modes (route data `mode`). */
export const routes: Routes = [
  { path: "", component: HomePage, data: { mode: "basic" } },
  { path: "custom", component: HomePage, data: { mode: "custom" } },
  { path: "advanced", component: AdvancedPage },
  { path: "**", redirectTo: "" },
];
