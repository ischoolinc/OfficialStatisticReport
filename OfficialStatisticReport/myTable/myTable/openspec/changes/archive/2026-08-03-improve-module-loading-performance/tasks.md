## 1. Module-load bootstrap guard (`Progress.cs`)

- [x] 1.1 Add a config-key constant for the "tag bootstrap complete" marker (e.g. `新生入學統計報表_預設類別已初始化`) alongside the existing `新生入學統計報表_來源目標設定Config` key usage pattern.
- [x] 1.2 In `Progress.main()`, read the marker via `K12.Data.School.Configuration[...]` before the `#region 加入預設的入學方式、入學身分、原住民類別` block; skip the entire block (the `TagConfig.SelectAll()` call and all conditional `TagConfig.Insert()` calls) when the marker is already set.
- [x] 1.3 After the bootstrap block completes successfully (whether it inserted anything or not), persist the marker so it is not run again.
- [x] 1.4 Verify the ribbon/menu registration (`RibbonBarItems`, `RoleAclSource` permission registration) still runs unconditionally on every startup, unaffected by the new guard.

## 2. Non-blocking `Form2` data load

- [x] 2.1 Identify the constructor calls that hit the server (`Column3Prepare()`'s `QueryHelper.Select(...)`, `LoadConfigXml()`'s `K12.Data.School.Configuration[...]`) and move the actual data fetch (not the UI-only setup like `Column2Prepare()`, `dataGridViewComboBoxExColumn2/3Prepare()`) out of the synchronous constructor path.
- [x] 2.2 Add a `BackgroundWorker` (or `Task.Run` with UI marshalling) to `Form2` that performs the tag query and config lookup off the UI thread, following the existing pattern already used by `_BGWClassStudentAbsenceDetail` elsewhere in `Form2.cs`.
- [x] 2.3 Disable `Column3`, `dataGridViewComboBoxExColumn2`, `dataGridViewComboBoxExColumn4`, and the `dataGridViewX1`/`X2`/`X3` config grids at construction time, and show the existing `loading.gif` resource, until the background load completes. (Implemented by disabling the parent `dataGridViewX1`/`X2`/`X3` grids, since `DataGridViewColumn`-derived `Column3`/`dataGridViewComboBoxExColumn2`/`4` have no `Enabled` property of their own.)
- [x] 2.4 On `RunWorkerCompleted` (or task continuation marshalled to the UI thread), populate the controls exactly as `Column3Prepare()`/`LoadConfigXml()` do today, then re-enable the controls and hide the loading indicator.
- [x] 2.5 Ensure `Form2()`'s constructor still calls `InitializeComponent()` and the non-blocking UI setup methods (`Column2Prepare()`, `dataGridViewComboBoxExColumn2Prepare()`, `dataGridViewComboBoxExColumn3Prepare()`, `SchoolYearItem()`, `LoadClassTypeCodeDic()`) synchronously as before, since those don't touch the server.

## 3. Verification

- [ ] 3.1 Manually test first-run behavior: clear/simulate an empty config marker, confirm tags are seeded once and the marker is written.
- [ ] 3.2 Manually test steady-state behavior: with the marker already set, confirm `Progress.main()` no longer issues a `TagConfig.SelectAll()` call (e.g. via logging or a debugger breakpoint) on subsequent app startups.
- [ ] 3.3 Manually open the 新生入學方式統計表 report and confirm the window appears immediately with a loading indicator, then populates its dropdowns/grids and becomes interactive once data arrives, with no change in the final displayed options/values compared to current behavior.
- [ ] 3.4 Confirm no regression in the report's existing print/export (`buttonX1_Click`) and saved-mapping (`SaveMappingXmlRecord`) functionality after these timing changes.
