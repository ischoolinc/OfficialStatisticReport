## ADDED Requirements

### Requirement: One-time tag bootstrap on module load
The module's `[FISCA.MainMethod]` entry point SHALL perform the default tag-category bootstrap (`TagConfig.SelectAll()` plus any conditional `TagConfig.Insert()` calls) at most once per school, persisting a completion marker so subsequent application startups skip the bootstrap entirely.

#### Scenario: First startup after install seeds tags
- **WHEN** the module loads and no "bootstrap complete" marker is present in `K12.Data.School.Configuration`
- **THEN** the module fetches existing tags via `TagConfig.SelectAll()`, inserts any missing default 入學方式/入學身分/原住民 tags, and persists a completion marker

#### Scenario: Subsequent startups skip the bootstrap
- **WHEN** the module loads and a "bootstrap complete" marker is already present in `K12.Data.School.Configuration`
- **THEN** the module SHALL NOT call `TagConfig.SelectAll()` or `TagConfig.Insert()` as part of startup

### Requirement: Non-blocking report window load
The report window (`Form2`) SHALL become visible to the user without waiting for its data-dependent controls (tag/category lists, saved source-target mapping config) to finish loading from the server.

#### Scenario: Window opens immediately with loading indicator
- **WHEN** the user opens the 新生入學方式統計表 report
- **THEN** the window is shown right away with a loading indicator visible and the data-dependent controls disabled

#### Scenario: Controls become usable once data arrives
- **WHEN** the background data load (tag list query and saved config lookup) completes
- **THEN** the loading indicator is hidden and the previously-disabled controls are populated and re-enabled, with no change to the values or options the user would have seen under the previous synchronous behavior
