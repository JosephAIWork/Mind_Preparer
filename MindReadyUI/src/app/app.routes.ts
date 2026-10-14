import { Routes } from "@angular/router";
import { AdvancedPage } from "./pages/advanced.page";
import { HomePage } from "./pages/home.page";

export const routes: Routes = [
  { path: "", component: HomePage },
  { path: "advanced", component: AdvancedPage },
  { path: "**", redirectTo: "" },
];
