# TODO - 新生入學統計設定來源同名自動帶入

## Goal

Adjust `Form2.cs` so that Source loading is consistent for:

- `dataGridViewX1`
- `dataGridViewX2`
- `dataGridViewX3`

Required rule:

```text
Config has target and source has a valid value
    -> use Config source

Config has no target
OR
Config has target but source is empty / whitespace
    -> check Source list for an exact same-name item

Exact same-name Source exists
    -> automatically fill it

No exact same-name Source exists
    -> keep Source blank
```

After modification, update:

`新生入學統計調整0915.md`

## Current Issue

Current `BindConfigMappingsToGrid()` treats any existing Config target as valid even when `source` is empty.

Current pattern:

```csharp
if (mappings.ContainsKey(target))
{
    source = mappings[target];
}
else if (sourceSet.Contains(target))
{
    source = target;
}
```

Example:

```xml
<item
    target="入學身份:外加錄取--原住民生"
    source="" />
```

Even if Source contains:

```text
入學身份:外加錄取--原住民生
```

the UI still shows blank because `mappings.ContainsKey(target)` is already true.

## Required Change

Modify only the decision logic inside `BindConfigMappingsToGrid()`.

Recommended implementation:

```csharp
string source = "";

if (mappings.ContainsKey(target) &&
    !string.IsNullOrWhiteSpace(mappings[target]))
{
    // Config has a valid source, use it first.
    source = mappings[target];
}
else if (sourceSet.Contains(target))
{
    // Config has no valid source.
    // If Source contains exact same text as Target, auto-fill it.
    source = target;
}
```

Keep existing row creation logic:

```csharp
DataGridViewRow row = new DataGridViewRow();
row.CreateCells(grid);

row.Cells[0].Value = target;
row.Cells[1].Value = source;

grid.Rows.Add(row);
```

## Apply to All Three DataGridViews

Do not duplicate the logic.

Continue using the shared `BindConfigMappingsToGrid(...)`.

### dataGridViewX1

Target:

```csharp
Column2.Items
```

Source:

```csharp
Column3.Items
```

### dataGridViewX2

Target:

```csharp
dataGridViewComboBoxExColumn1.Items
```

Source:

```csharp
dataGridViewComboBoxExColumn2.Items
```

### dataGridViewX3

Target:

```csharp
dataGridViewComboBoxExColumn3.Items
```

Source:

```csharp
dataGridViewComboBoxExColumn4.Items
```

## Exact Match Rule

Use exact string comparison only.

Keep:

```csharp
HashSet<string> sourceSet =
    new HashSet<string>(sourceItems, StringComparer.Ordinal);
```

Then use:

```csharp
sourceSet.Contains(target)
```

Do NOT use partial, fuzzy, `StartsWith`, `EndsWith`, or case-insensitive matching.

## Expected Behavior

### Config source has value

Use Config source even if a same-name Source also exists.

### Config source is empty

If same-name Source exists, auto-fill it.

### Config target does not exist

If same-name Source exists, auto-fill it.

### No Config source and no same-name Source

Keep Source blank.

## Preserve Existing Behavior

Do not change previous fixes:

- `dataGridViewX1` always displays all predefined targets.
- `dataGridViewX2` always displays all predefined targets.
- `dataGridViewX3` always displays all predefined targets.
- Clear DataGridView rows before reload.
- Config does not control row count.
- Non-empty Config source remains highest priority.

## Do Not Change Other Logic

Do not modify:

- `Column2Prepare()`
- `dataGridViewComboBoxExColumn2Prepare()`
- `dataGridViewComboBoxExColumn3Prepare()`
- `Column3Prepare()`
- `SaveMappingXmlRecord()`
- `ReadXMLMappingData()`
- `SetXMLMappingDataKey()`
- Tag ID lookup
- student filtering
- report generation
- Excel export
- admission method statistics
- admission identity statistics
- Aboriginal identity statistics

This task only changes default Source loading behavior.

## Validation

1. X1: Config source empty + exact same-name Source exists -> auto-fill.
2. X2: Config source empty + exact same-name Source exists -> auto-fill.
3. X3: Config source empty + exact same-name Source exists -> auto-fill.
4. Config source has value -> keep Config source.
5. Config source is whitespace -> treat as unset and try same-name Source.
6. Only similar/partial Source exists -> remain blank.
7. Reload -> no duplicate rows; same-name fallback still works.
8. Save and reopen -> saved Source remains correct; no report/statistics regression.

## Completion Record

After implementation, update `新生入學統計調整0915.md` with:

1. Modified file(s).
2. Original issue.
3. Updated Source loading priority.
4. Empty Config source fallback behavior.
5. Exact same-name matching rule.
6. Confirmation that X1/X2/X3 use the same shared logic.
7. Validation results.
8. Confirmation that report/statistics logic was not intentionally changed.

## Acceptance Criteria

- [x] X1 auto-fills same-name Source when Config source is missing/blank.
- [x] X2 auto-fills same-name Source when Config source is missing/blank.
- [x] X3 auto-fills same-name Source when Config source is missing/blank.
- [x] Existing non-empty Config source remains highest priority.
- [x] Empty/whitespace Config source is treated as not configured.
- [x] Matching is exact text only.
- [x] No partial/fuzzy matching is introduced.
- [x] Existing default-target loading behavior remains unchanged.
- [x] Existing save/mapping/report/statistics logic remains unchanged.
- [x] Changes are documented in `新生入學統計調整0915.md`.
