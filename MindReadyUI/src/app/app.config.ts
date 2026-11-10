import { provideHttpClient, withFetch, withInterceptors } from "@angular/common/http";
import { ApplicationConfig, provideBrowserGlobalErrorListeners } from "@angular/core";
import { provideRouter, withComponentInputBinding } from "@angular/router";
import { routes } from "./app.routes";
import { expiryInterceptor } from "./core/expiry";
import { feedbackInterceptor } from "./core/feedback.service";
import { restartInterceptor } from "./core/restart.interceptor";

export const appConfig: ApplicationConfig = {
  providers: [provideBrowserGlobalErrorListeners(), provideRouter(routes, withComponentInputBinding()), provideHttpClient(withFetch(), withInterceptors([feedbackInterceptor, expiryInterceptor, restartInterceptor]))],
};
