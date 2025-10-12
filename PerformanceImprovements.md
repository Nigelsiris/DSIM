# Performance Improvement Opportunities

## Cache the spreadsheet handle
Multiple functions call `SpreadsheetApp.getActiveSpreadsheet()` before every operation, including `updateAllEpjStatuses`, `loginAndGetUserView`, and several admin helpers.【F:code.js†L52-L84】【F:code.js†L176-L205】【F:code.js†L293-L360】 Because Apps Script re-evaluates globals on cold starts, caching the active spreadsheet in a top-level variable (for example, `const ss = SpreadsheetApp.getActive();`) lets each invocation reuse the same handle and removes dozens of repeated lookups.

## Read only populated rows on login
`loginAndGetUserView` currently loads the entire `A:D` block from the Users sheet, which can include tens of thousands of empty rows and increases response latency.【F:code.js†L176-L205】 Switching to `getRange(2, 1, lastRow - 1, 4)` (and skipping the loop if `lastRow <= 1`) keeps the data volume proportional to the actual user count.

## Incremental status recomputation
`updateAllEpjStatuses` and admin overrides call `getDataRange().getValues()` on the log sheet, then walk the entire dataset to refresh status and locate trips.【F:code.js†L52-L84】【F:code.js†L353-L370】 Persisting the last processed row in `PropertiesService` (or precomputing a map in cache) would let the script process only the newest rows and avoid re-reading historical data on every submission.

## Cache credential lookups
After reading the Users sheet, `loginAndGetUserView` linearly searches for the username on each login.【F:code.js†L176-L205】 Building a username→record map once (and storing it in `CacheService` or `PropertiesService`) would turn the lookup into an `O(1)` operation and reduce spreadsheet traffic during peak login periods.
