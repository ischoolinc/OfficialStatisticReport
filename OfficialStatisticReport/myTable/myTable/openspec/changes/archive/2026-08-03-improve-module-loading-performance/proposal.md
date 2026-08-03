## Why

This module (`myTable` / 公務統計報表 — 新生入學方式統計表) is loaded by the 1Campus Desktop host on every application startup via its `[FISCA.MainMethod]` entry point (`Progress.main()`), and its report window (`Form2`) is opened synchronously on the UI thread. Two concrete bottlenecks were found by reading the code:

1. `Progress.main()` unconditionally calls `TagConfig.SelectAll()` (a full server round-trip fetching every tag for the school) on **every** application startup, for every school using this module — even though the follow-up `TagConfig.Insert()` calls it guards against are only ever needed once, the first time the module runs after install. There is no "already initialized" check, so the network round-trip is paid repeatedly forever.
2. `Form2`'s constructor performs synchronous, blocking data access (`_Q.Select("select * from tag where category='Student' ...")` in `Column3Prepare()`, and `K12.Data.School.Configuration[...]` in `LoadConfigXml()`) before the window is shown, so the report dialog appears to hang for however long those round-trips take, with no feedback to the user despite a `loading.gif` resource already existing in the project.

Together these make both "the module loading into the desktop shell" and "opening the report" feel slow, and the cost scales with server latency and tag-table size rather than being fixed.

## What Changes

- Guard the one-time tag bootstrap in `Progress.main()` with a persisted "already initialized" marker (using the same `K12.Data.School.Configuration` mechanism already used elsewhere in this module) so `TagConfig.SelectAll()` and the conditional `TagConfig.Insert()` calls run at most once per school, not on every startup.
- Move the remaining startup-time work in `Progress.main()` (ribbon/menu registration) to stay synchronous and cheap; keep any first-run tag seeding off the critical path where possible.
- Defer `Form2`'s blocking data loads (`Column3Prepare()`'s SQL query, `LoadConfigXml()`'s config lookup) so the window can display immediately and populate asynchronously, reusing the project's existing `loading.gif` / `BackgroundWorker` pattern (already used elsewhere in `Form2.cs`) to show progress instead of freezing the UI thread.
- No change to the report's business logic, calculations, or output — this is purely about when/how data is fetched relative to when the UI becomes visible.

## Capabilities

### New Capabilities
- `module-loading-performance`: Testable requirements for how this module bootstraps at application startup and how its report window loads data, so redundant network calls and UI-blocking loads don't regress in the future.

### Modified Capabilities
(none — no existing spec'd capabilities in this project yet; the report still shows the same data and options, only faster and with feedback while loading)

## Impact

- `Progress.cs` — `main()` gains a persisted-flag guard before the `TagConfig.SelectAll()`/`Insert()` block.
- `Form2.cs` — constructor and `LoadConfigXml()`/`Column3Prepare()` reworked to load data asynchronously (likely via the existing `BackgroundWorker` field `_BGWClassStudentAbsenceDetail` pattern or a new one) with the window shown immediately and `loading.gif` shown while data populates.
- No database schema, external API, or configuration format changes. `K12.Data.School.Configuration` gains one new key for the initialization marker.
- Risk: must ensure the async load in `Form2` completes (and disables relevant controls) before the user can interact with fields that depend on that data, to avoid race conditions.
