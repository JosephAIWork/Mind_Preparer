# UC-6 -- Shlomo's IFRS risk model

What the app learned from this use case (moved out of the app's comments and messages;
the app itself names no client). Grouped by the app file the insight belongs to.

## app/prep.py

- **titles and moved labels change values**: `=COUNTA(CoverageDetails!A:A)` counted a title written into column A and the model picked a different policy; `=Information!A1` on five sheets read 0 after the company name was "moved" into a title. Hence the reference index every write site consults.

## mind_client.py

- **selectors**: The Mind project-operation selectors were proven by the Shlomo closed loop.
