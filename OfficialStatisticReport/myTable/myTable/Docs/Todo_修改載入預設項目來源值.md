# TODO - 新生入學統計來源自動同名比對

## Goal

Modify `Form2.cs` so that when loading:

- `dataGridViewX1`
- `dataGridViewX2`
- `dataGridViewX3`

the existing Config mapping still has the highest priority.

If Config does NOT contain a mapping for a target, check whether the available Source list contains an item whose text is exactly the same as the Target text.

If an exact same-name Source exists, automatically fill it into the Source cell.

If no exact same-name Source exists, keep the Source cell blank.

After completing the modification, document the changes in:

`新生入學統計調整0915.md`

---

## Current Behavior

Current loading logic in `BindConfigMappingsToGrid()` is:

```csharp
row.Cells[0].Value = target;
row.Cells[1].Value = mappings.ContainsKey(target)
    ? mappings[target]
    : "";
```

Current behavior:

```text
Config has matching target
    -> load Config source

Config does not have matching target
    -> source is blank
```

This does not check whether the Source list contains the exact same text as the Target.

---

## Required Behavior

Change the logic to:

```text
1. Config contains this target
   -> use Config source

2. Config does not contain this target
   -> check available Source items

3. Source list contains an item exactly equal to target
   -> automatically use that Source

4. No exact same-name Source exists
   -> source remains blank
```

Config must remain the first priority.

Do NOT overwrite a Config source just because an exact same-name Source exists.

---

## Example 1 - Config Has Mapping

Target:

```text
入學方式:免試入學--校內直升
```

Config:

```xml
<item
    target="入學方式:免試入學--校內直升"
    source="招生方式:校內直升" />
```

Available Source list also contains:

```text
入學方式:免試入學--校內直升
```

Expected UI:

```text
標記：入學方式:免試入學--校內直升
來源：招生方式:校內直升
```

Reason:

Config has higher priority.

Do NOT replace it with the same-name Source.

---

## Example 2 - Config Has No Mapping, Same-name Source Exists

Target:

```text
入學方式:免試入學--優先免試
```

Config has no matching target.

Available Source list contains:

```text
入學方式:免試入學--優先免試
```

Expected:

```text
標記：入學方式:免試入學--優先免試
來源：入學方式:免試入學--優先免試
```

The Source is automatically filled.

---

## Example 3 - Config Has No Mapping, No Same-name Source

Target:

```text
入學方式:免試入學--完全免試
```

Config has no matching target.

Available Source list does not contain the exact same text.

Expected:

```text
標記：入學方式:免試入學--完全免試
來源：
```

Keep the Source blank.

---

## Scope

Apply the same behavior to all three grids.

### 1. `dataGridViewX1`

Target source:

```csharp
Column2.Items
```

Source column options:

```csharp
Column3.Items
```

---

### 2. `dataGridViewX2`

Target source:

```csharp
dataGridViewComboBoxExColumn1.Items
```

Source column options:

```csharp
dataGridViewComboBoxExColumn2.Items
```

---

### 3. `dataGridViewX3`

Target source:

```csharp
dataGridViewComboBoxExColumn3.Items
```

Source column options:

```csharp
dataGridViewComboBoxExColumn4.Items
```

---

## Important Source Data

All three Source option lists are ultimately populated in `Column3Prepare()` from:

```sql
select *
from tag
where category='Student'
order by prefix,name
```

The display value is built from:

```text
prefix + ":" + name
```

When prefix is empty, the leading `:` is removed before adding it to the Source ComboBox items.

Therefore, same-name comparison should use the actual displayed Source text, not raw database `prefix:name` values.

---

## Recommended Implementation

Modify `BindConfigMappingsToGrid()` so it also receives the Source item collection.

For example:

```csharp
private void BindConfigMappingsToGrid(
    DataGridView grid,
    IEnumerable<string> predefinedTargets,
    IEnumerable<string> sourceItems,
    XmlElement config,
    string sectionName)
```

Then build a lookup for Source display values.

Conceptual implementation:

```csharp
HashSet<string> sourceSet = new HashSet<string>(
    sourceItems,
    StringComparer.Ordinal
);
```

Then load each target using this priority:

```csharp
string source = "";

if (mappings.ContainsKey(target))
{
    // Highest priority: existing Config mapping
    source = mappings[target];
}
else if (sourceSet.Contains(target))
{
    // No Config mapping, but Source has exact same text
    source = target;
}

row.Cells[0].Value = target;
row.Cells[1].Value = source;
```

Use exact string comparison.

Do NOT use:

- `Contains`
- partial match
- fuzzy match
- starts-with
- ends-with

The requirement is exact same text only.

---

## Update Calls to `BindConfigMappingsToGrid()`

Current calls are similar to:

```csharp
BindConfigMappingsToGrid(
    dataGridViewX1,
    Column2.Items,
    config,
    "入學方式");
```

Update them to also pass Source items.

Recommended mapping:

```csharp
BindConfigMappingsToGrid(
    dataGridViewX1,
    Column2.Items.Cast<string>(),
    Column3.Items.Cast<string>(),
    config,
    "入學方式");

BindConfigMappingsToGrid(
    dataGridViewX2,
    dataGridViewComboBoxExColumn1.Items.Cast<string>(),
    dataGridViewComboBoxExColumn2.Items.Cast<string>(),
    config,
    "入學身分");

BindConfigMappingsToGrid(
    dataGridViewX3,
    dataGridViewComboBoxExColumn3.Items.Cast<string>(),
    dataGridViewComboBoxExColumn4.Items.Cast<string>(),
    config,
    "新生中具原住民身分者");
```

If the existing collection type does not support the exact syntax above, use the safest compatible implementation for the current .NET Framework project.

Do not introduce unnecessary architectural changes.

---

## Preserve Current Default Target Behavior

Keep the current behavior already implemented:

- `dataGridViewX1` always displays all 11 predefined targets.
- `dataGridViewX2` always displays all 4 predefined targets.
- `dataGridViewX3` always displays its predefined target.
- Config does not decide row count.
- Missing Config target does not remove the row.
- Reloading does not create duplicate rows.

Do not undo the previous fix.

---

## Preserve Existing Config Priority

This requirement is critical.

If Config contains:

```xml
<item
    target="A"
    source="B" />
```

and Source items contain:

```text
A
B
```

Expected result:

```text
Target = A
Source = B
```

NOT:

```text
Target = A
Source = A
```

Same-name auto-fill is only a fallback when Config has no target mapping.

---

## Empty Config Source Case

Treat the existence of the target in Config as a mapping decision.

If Config contains:

```xml
<item
    target="A"
    source="" />
```

do not automatically replace it with `A` unless the existing business rule explicitly considers an empty source equivalent to no mapping.

For this task, prefer the safer rule:

```text
Config target exists
    -> respect Config source, even when blank
```

This prevents the new fallback logic from silently changing an intentionally blank Config value.

---

## Do Not Change Saving Logic

Do not modify `SaveMappingXmlRecord()` unless compilation requires a minimal related fix.

Existing behavior should remain:

```text
User selects Source
    -> SaveMappingXmlRecord()
    -> Config target/source saved
```

The new automatic same-name Source value should naturally be saved later if the existing save flow persists the current DataGridView values.

Do not add automatic saving during form load.

---

## Do Not Change Mapping / Report Logic

Do not modify unrelated logic, including:

- `Column2Prepare()`
- `dataGridViewComboBoxExColumn2Prepare()`
- `dataGridViewComboBoxExColumn3Prepare()`
- `Column3Prepare()`
- `ReadXMLMappingData()`
- `SetXMLMappingDataKey()`
- `SaveMappingXmlRecord()`
- Tag ID lookup
- student filtering
- Excel export
- report statistics
- admission method statistics
- admission identity statistics
- Aboriginal identity statistics

This task only changes the default Source value shown when Config has no target mapping.

---

## Validation

### Test 1 - Config Mapping Exists

Target:

```text
入學方式:免試入學--校內直升
```

Config source:

```text
自訂來源:校內直升
```

Available Source list also contains the same text as Target.

Expected:

```text
Source = 自訂來源:校內直升
```

Config wins.

---

### Test 2 - No Config Mapping, Same-name Source Exists

Config does not contain:

```text
入學方式:免試入學--優先免試
```

Source list contains:

```text
入學方式:免試入學--優先免試
```

Expected:

```text
Source = 入學方式:免試入學--優先免試
```

---

### Test 3 - No Config Mapping, No Same-name Source

Expected:

```text
Source = blank
```

No other Source should be guessed.

---

### Test 4 - X2 Same-name Match

For one of the four `dataGridViewX2` targets:

- remove its Config mapping;
- ensure Source items contain the exact same displayed text.

Expected:

- Source auto-fills with the same text.

---

### Test 5 - X3 Same-name Match

Target:

```text
新生中具原住民身分者
```

If Config has no mapping and Source items contain the exact same text:

Expected:

```text
Source = 新生中具原住民身分者
```

---

### Test 6 - Partial Text Must Not Match

Target:

```text
入學方式:免試入學--優先免試
```

Source only contains:

```text
優先免試
```

Expected:

```text
Source = blank
```

Do not use partial matching.

---

### Test 7 - Reload

Call the existing reload/reset flow.

Expected:

- X1 remains 11 predefined rows.
- X2 remains 4 predefined rows.
- X3 remains 1 predefined row.
- Config mappings still have priority.
- Same-name fallback is applied only to targets without Config mappings.
- No duplicate rows appear.

---

### Test 8 - Save and Reopen

1. Start with a target that has no Config mapping.
2. Let same-name fallback auto-fill its Source.
3. Execute the existing save/print flow.
4. Reopen the form.

Expected:

- the mapping is now read normally from Config;
- displayed result remains the same;
- no report logic regression occurs.

---

## Completion Record

After implementation, update:

`新生入學統計調整0915.md`

Record:

1. Modified file(s).
2. Original issue.
3. New Source loading priority.
4. Config mapping priority behavior.
5. Exact same-name Source fallback behavior.
6. Exact string comparison rule.
7. Behavior when no same-name Source exists.
8. Confirmation that X1/X2/X3 all use the same logic.
9. Validation results.
10. Confirmation that report/statistics logic was not intentionally changed.

---

## Acceptance Criteria

The task is complete only when:

- [ ] Existing Config target/source mapping remains highest priority.
- [ ] When Config has no target mapping, exact same-name Source is automatically selected.
- [ ] No partial/fuzzy matching is used.
- [ ] If no exact same-name Source exists, Source remains blank.
- [ ] `dataGridViewX1` applies the fallback logic.
- [ ] `dataGridViewX2` applies the fallback logic.
- [ ] `dataGridViewX3` applies the fallback logic.
- [ ] Previous default-target loading behavior remains intact.
- [ ] No duplicate rows are introduced.
- [ ] Existing save/mapping/report/statistics logic remains unchanged.
- [ ] Changes are documented in `新生入學統計調整0915.md`.
