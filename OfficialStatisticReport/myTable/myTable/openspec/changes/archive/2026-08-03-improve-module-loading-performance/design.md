## Context

`myTable.dll` is a plugin module for the 1Campus Desktop host application (built on the FISCA framework). The host discovers and loads all installed modules at application startup and invokes each module's `[FISCA.MainMethod]`-attributed entry point synchronously as part of that startup sequence — in this module, that's `Progress.main()`. Any blocking work done there directly extends how long the whole desktop application takes to become usable, for every school that has this module installed.

Separately, when a user clicks the "新生入學方式統計表" ribbon button, `Progress.main()`'s click handler constructs `Form2` and calls `form.ShowDialog()`. `Form2`'s constructor currently does blocking data access before the window paints, so the dialog appears frozen.

Both paths currently use synchronous, unconditional server round-trips:
- `Progress.main()`: `TagConfig.SelectAll()` (always) + `TagConfig.Insert()` (per missing tag, first run only) — added in commit `2b396b7` to seed default tag categories once per school.
- `Form2()` constructor: `Column3Prepare()` runs `QueryHelper.Select("select * from tag where category='Student' ...")`, and `LoadConfigXml()` reads `K12.Data.School.Configuration[...]`.

The project already has a `loading.gif` resource and a `BackgroundWorker` field (`_BGWClassStudentAbsenceDetail`) demonstrating the existing async pattern used elsewhere in `Form2.cs`, so we're extending an established pattern rather than introducing a new one.

## Goals / Non-Goals

**Goals:**
- Eliminate the unconditional `TagConfig.SelectAll()`/`Insert()` round-trip from every application startup; it should run at most once per school (first run after this module is installed/updated).
- Make `Form2` display its window before blocking data loads complete, with a visible loading indicator, instead of freezing the UI thread.
- Preserve all existing business logic, computed values, and report output exactly as-is.

**Non-Goals:**
- No change to the report's statistical calculations, Excel export logic, or UI layout.
- No change to the underlying `TagConfig`, `K12.Data`, or `QueryHelper` APIs themselves — this is a call-site/timing change only.
- Not attempting to batch or optimize the `TagConfig.Insert()` calls themselves (still one-by-one on first run); the goal is to make that cost a one-time event, not to make the one-time event itself faster.

## Decisions

**1. Guard the tag bootstrap with a persisted marker, not a "try/catch and ignore" or a version check against tag count.**
Use `K12.Data.School.Configuration` (the same config store `LoadConfigXml()` already uses in this module) with a new key, e.g. `新生入學統計報表_預設類別已初始化`, set to a marker value after the bootstrap block completes successfully. On `main()` entry, check this key first; only run `TagConfig.SelectAll()`/`Insert()` if it's absent.
- *Alternative considered*: Infer "already done" by checking if all expected tags exist (still requires the `SelectAll()` every time — doesn't remove the recurring round-trip, only removes the inserts). Rejected because the `SelectAll()` itself is the main recurring cost.
- *Alternative considered*: Run the whole bootstrap block on a background thread inside `main()` so it doesn't block startup even when it does run. Kept as a secondary mitigation (see below) but not a replacement for the marker, since an unguarded background call still repeats a needless server round-trip on every single startup.

**2. Move `Form2`'s blocking data loads to run after the window is shown, using a `BackgroundWorker` (or `Task.Run` + `Invoke` marshal-back, consistent with existing `System.Threading.Tasks` usage already imported in `Form2.cs`).**
Show the window immediately with the loading indicator visible and dependent controls (`Column3`, `dataGridViewComboBoxExColumn2/4`, the config grids) disabled; populate them when the background load completes, then enable controls and hide the indicator.
- *Alternative considered*: Prefetch data before `ShowDialog()` inside `Progress.main()`'s click handler (e.g., a splash/spinner before the dialog even opens). Rejected — same total wait, just moved earlier, and duplicates the loading-indicator logic that could live once inside `Form2`.

**3. Keep the fix scoped to this module's own code paths.**
No changes to `Lib/FISCA*.dll`, `K12.Data.dll`, or other referenced binaries — those are external, versioned dependencies (`<Private>False</Private>` in the csproj) and out of scope.

## Risks / Trade-offs

- **[Risk]** If `Form2` controls are enabled before the async load finishes, the user could interact with incomplete data (e.g., pick a source item that hasn't loaded into a combo box yet) → **Mitigation**: disable the dependent controls up front and only re-enable them in the background worker's `RunWorkerCompleted` handler.
- **[Risk]** A school with an existing partial/inconsistent tag setup (e.g., manually deleted some seeded tags) will no longer get them re-added automatically once the "initialized" marker is set → **Mitigation**: this matches the intent of the original code (seed once, then leave school data alone); document the config key so support staff can clear it manually to re-trigger seeding if ever needed.
- **[Trade-off]** Config key check adds one small `K12.Data.School.Configuration[...]` read to every startup in place of the `TagConfig.SelectAll()` — strictly cheaper (single key lookup vs. full tag table scan) but not literally free. Accepted as the standard cost of a persisted "already done" flag.

## Migration Plan

- No data migration required. On first startup after deployment, `main()` finds no marker, runs the existing seed logic exactly as before (same behavior as today for a school seeing this for the first time), then writes the marker.
- Schools that have already had tags seeded under the old code will simply get the marker set on their next startup after upgrade, without re-inserting anything (existing `Contains`/`Remove` de-dup logic in the bootstrap block already prevents duplicate inserts, so running it one extra time post-upgrade is harmless).
- Rollback: reverting this module's DLL removes the marker check; worst case is a return to the previous behavior (redundant `SelectAll()` per startup) — no data loss either direction.

## Open Questions

- Should the tag-bootstrap block in `Progress.main()` also move to a background thread (in addition to the marker guard) so that even a school's *first* startup after install isn't blocked by the insert round-trips? Leaning yes since it's low-risk and consistent with Decision 2, but confirming isn't required to unblock implementation — can be included in tasks as a stretch item.
