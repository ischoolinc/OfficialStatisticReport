# TODO - 新生入學統計「重設」功能調整

## Goal

Modify the `重設` function in `Form2.cs`.

Target event:

```csharp
linkLabel1_LinkClicked
```

The reset behavior must be redefined as:

1. Clear all current rows in:
   - `dataGridViewX1`
   - `dataGridViewX2`
   - `dataGridViewX3`
2. Reload the predefined Target items for all three DataGridViews.
3. Reload the Source item lists from the current Student Tag data.
4. Compare Target and Source by exact text.
5. If Source contains an item exactly equal to the Target, automatically fill that Source.
6. If no exact same-name Source exists, leave Source blank.
7. Do NOT load old Config mappings during reset.
8. Do NOT immediately save/reset Config during reset.
9. Existing Config should only be overwritten later by the existing save flow, such as `SaveMappingXmlRecord()` when printing.

After modification, document the changes in:

`新生入學統計調整0916.md`

---

## Current Source Data

The Source lists of all three DataGridViews come from:

```sql
select *
from tag
where category='Student'
order by prefix,name
```

This logic is currently implemented in:

```csharp
Column3Prepare()
```

Display text is built from:

```text
prefix + ":" + name
```

If `prefix` is empty, the leading `:` is removed before adding the item to the UI.

The same Student Tag source list is added to:

```csharp
Column3.Items
dataGridViewComboBoxExColumn2.Items
dataGridViewComboBoxExColumn4.Items
```

Mapping:

```text
dataGridViewX1 Source -> Column3.Items
dataGridViewX2 Source -> dataGridViewComboBoxExColumn2.Items
dataGridViewX3 Source -> dataGridViewComboBoxExColumn4.Items
```

---

## Current Problem

Current reset flow:

```csharp
private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
{
    try
    {
        AccessHelper _A = new AccessHelper();
        List<myTableUDT> UDTlist = _A.Select<myTableUDT>();
        _A.DeletedValues(UDTlist);

        dataGridViewX1.Rows.Clear();

        LoadConfigXml();
    }
    catch
    {
        MessageBox.Show("網路或資料庫異常,請稍後再試...");
    }
}
```

Problems:

1. It only explicitly clears `dataGridViewX1`.
2. It deletes old UDT data even though the current active mapping flow uses Config XML.
3. It calls `LoadConfigXml()`.
4. `LoadConfigXml()` loads existing Config sources first.
5. Therefore reset does not really reset to the current default/source same-name state.
6. Old Config source mappings can immediately come back after pressing reset.

The reset function should no longer depend on old UDT or old Config mappings.

---

## New Reset Flow

Implement the reset flow in this order:

```text
linkLabel1_LinkClicked
        ↓
Clear X1 / X2 / X3 rows
        ↓
Clear all three Source ComboBox item collections
        ↓
Reload Student Tag Source data
        ↓
Rebuild X1 default Targets
        ↓
Exact-match Target against X1 Source list
        ↓
Rebuild X2 default Targets
        ↓
Exact-match Target against X2 Source list
        ↓
Rebuild X3 default Targets
        ↓
Exact-match Target against X3 Source list
        ↓
Same-name Source exists -> auto-fill
No same-name Source -> blank
```

---

## 1. Clear All Three DataGridViews

Reset must clear:

```csharp
dataGridViewX1.Rows.Clear();
dataGridViewX2.Rows.Clear();
dataGridViewX3.Rows.Clear();
```

Do not clear only `dataGridViewX1`.

---

## 2. Reload Source Item Lists

Before calling `Column3Prepare()` again, clear the existing Source item collections to avoid duplicates:

```csharp
Column3.Items.Clear();
dataGridViewComboBoxExColumn2.Items.Clear();
dataGridViewComboBoxExColumn4.Items.Clear();
```

Then call:

```csharp
Column3Prepare();
```

This reloads the latest Student Tag data from:

```sql
select *
from tag
where category='Student'
order by prefix,name
```

This also recreates:

```csharp
_column3Items
```

Do not append duplicate Source options.

---

## 3. Do NOT Re-run Target Prepare Methods

Do not call these again during reset:

```csharp
Column2Prepare();
dataGridViewComboBoxExColumn2Prepare();
dataGridViewComboBoxExColumn3Prepare();
```

Reason:

These methods use `Items.Add(...)`.

They were already called once in the constructor.

Calling them again during reset may duplicate the predefined Target items.

During reset, reuse the already prepared Target collections:

```csharp
Column2.Items
dataGridViewComboBoxExColumn1.Items
dataGridViewComboBoxExColumn3.Items
```

---

## 4. Rebuild `dataGridViewX1`

Target list:

```csharp
Column2.Items
```

Source list:

```csharp
Column3.Items
```

Create all predefined X1 target rows.

For each Target:

```text
if Source list contains exactly the same displayed text
    -> Source = Target
else
    -> Source = blank
```

Do not use Config during reset.

---

## 5. Rebuild `dataGridViewX2`

Target list:

```csharp
dataGridViewComboBoxExColumn1.Items
```

Source list:

```csharp
dataGridViewComboBoxExColumn2.Items
```

Use the exact same reset behavior:

```text
Target == Source exactly
    -> auto-fill
otherwise
    -> blank
```

Do not load existing Config source.

---

## 6. Rebuild `dataGridViewX3`

Target list:

```csharp
dataGridViewComboBoxExColumn3.Items
```

Source list:

```csharp
dataGridViewComboBoxExColumn4.Items
```

Use the same exact-match reset behavior.

---

## 7. Reuse Existing Shared Binding Logic Safely

Current helper:

```csharp
BindConfigMappingsToGrid(
    DataGridView grid,
    IEnumerable<string> predefinedTargets,
    IEnumerable<string> sourceItems,
    XmlElement config,
    string sectionName)
```

Current behavior already supports exact same-name fallback when Config has no valid mapping.

A safe minimal-change implementation is to reuse it during reset with:

```csharp
config = null
```

Example:

```csharp
BindConfigMappingsToGrid(
    dataGridViewX1,
    Column2.Items.Cast<string>(),
    Column3.Items.Cast<string>(),
    null,
    "入學方式");

BindConfigMappingsToGrid(
    dataGridViewX2,
    dataGridViewComboBoxExColumn1.Items.Cast<string>(),
    dataGridViewComboBoxExColumn2.Items.Cast<string>(),
    null,
    "入學身分");

BindConfigMappingsToGrid(
    dataGridViewX3,
    dataGridViewComboBoxExColumn3.Items.Cast<string>(),
    dataGridViewComboBoxExColumn4.Items.Cast<string>(),
    null,
    "新生中具原住民身分者");
```

Because `config == null`:

```text
No saved mappings are loaded
        ↓
sourceSet.Contains(target)
        ↓
Exact same-name Source is auto-filled
```

This is preferred over duplicating the same matching logic three times.

---

## 8. Remove Old UDT Reset Dependency

Current reset code uses:

```csharp
AccessHelper _A = new AccessHelper();
List<myTableUDT> UDTlist = _A.Select<myTableUDT>();
_A.DeletedValues(UDTlist);
```

The current active flow uses:

```csharp
SaveMappingXmlRecord()
LoadConfigXml()
```

with:

```text
新生入學統計報表_來源目標設定Config
```

Therefore, do not use UDT deletion as part of the new reset behavior unless it is still required by another active feature.

For this task, reset should focus on the current UI and Student Tag source state.

Do not modify unrelated legacy methods unless necessary.

---

## 9. Do NOT Call `LoadConfigXml()` During Reset

This is critical.

Do not call:

```csharp
LoadConfigXml();
```

inside the new reset event.

Reason:

`LoadConfigXml()` intentionally restores existing Config source mappings.

The new reset requirement is:

```text
Ignore old Config mapping
        ↓
Load default Targets
        ↓
Reload current Student Tag Sources
        ↓
Exact same-name matching only
```

---

## 10. Do NOT Save Config During Reset

Pressing reset should only reset the current screen state.

Do not call:

```csharp
SaveMappingXmlRecord();
```

inside reset.

Do not directly clear or overwrite:

```text
新生入學統計報表_來源目標設定Config
```

during reset.

Expected behavior:

```text
Press Reset
    -> UI becomes default + exact same-name Source mapping

User can review/change Sources

Press Print
    -> existing SaveMappingXmlRecord()
    -> current UI mappings are saved to Config
```

This avoids destroying saved settings immediately when the user clicks reset.

---

## Recommended Reset Implementation

```csharp
private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
{
    try
    {
        dataGridViewX1.Rows.Clear();
        dataGridViewX2.Rows.Clear();
        dataGridViewX3.Rows.Clear();

        Column3.Items.Clear();
        dataGridViewComboBoxExColumn2.Items.Clear();
        dataGridViewComboBoxExColumn4.Items.Clear();

        Column3Prepare();

        BindConfigMappingsToGrid(
            dataGridViewX1,
            Column2.Items.Cast<string>(),
            Column3.Items.Cast<string>(),
            null,
            "入學方式");

        BindConfigMappingsToGrid(
            dataGridViewX2,
            dataGridViewComboBoxExColumn1.Items.Cast<string>(),
            dataGridViewComboBoxExColumn2.Items.Cast<string>(),
            null,
            "入學身分");

        BindConfigMappingsToGrid(
            dataGridViewX3,
            dataGridViewComboBoxExColumn3.Items.Cast<string>(),
            dataGridViewComboBoxExColumn4.Items.Cast<string>(),
            null,
            "新生中具原住民身分者");
    }
    catch
    {
        MessageBox.Show("網路或資料庫異常,請稍後再試...");
    }
}
```

Use this as implementation guidance.

Adjust only if required for compatibility with the current project.

---

## Exact Match Rule

Keep exact comparison:

```csharp
HashSet<string> sourceSet =
    new HashSet<string>(sourceItems, StringComparer.Ordinal);
```

and:

```csharp
sourceSet.Contains(target)
```

Do NOT use partial, fuzzy, `Contains`, `StartsWith`, `EndsWith`, or case-insensitive matching.

---

## Validation

### Test 1 - X1 Reset

Before reset:

```text
Target = 入學方式:免試入學--校內直升
Source = custom Config value
```

Current Student Tag list contains:

```text
入學方式:免試入學--校內直升
```

After reset:

```text
Source = 入學方式:免試入學--校內直升
```

Old Config value must NOT be restored.

### Test 2 - X2 Reset

Current Student Tag list contains an exact match for:

```text
入學身分:外加錄取--原住民生
```

After reset:

```text
Source = 入學身分:外加錄取--原住民生
```

### Test 3 - X3 Reset

Current Student Tag list contains:

```text
新生中具原住民身分者
```

After reset:

```text
Source = 新生中具原住民身分者
```

### Test 4 - No Same-name Source

If a predefined Target has no exact matching Student Tag:

Expected:

```text
Source = blank
```

### Test 5 - Source Refresh

1. Open the form.
2. Add or modify a Student Tag through the system.
3. Press reset.

Expected:

- Source dropdown lists are refreshed from the latest `tag` table data.
- No duplicate Source items appear.

### Test 6 - No Duplicate Targets

Press reset multiple times.

Expected:

- X1 row count remains the predefined count.
- X2 row count remains the predefined count.
- X3 row count remains the predefined count.
- Target item collections do not keep growing.

### Test 7 - Config Not Immediately Modified

1. Existing Config has a custom mapping.
2. Press reset.
3. Close the form without printing/saving.

Expected:

- reset only changes the current UI;
- reset itself does not directly write new mappings into Config.

### Test 8 - Save After Reset

1. Press reset.
2. Verify exact same-name Sources are filled.
3. Press Print.

Expected:

```text
buttonX1_Click
    -> SaveMappingXmlRecord()
```

saves the reset/current screen mappings into Config normally.

---

## Do Not Change Other Logic

Do not modify unrelated logic, including:

- form initialization flow
- `LoadConfigXml()` normal startup behavior
- `SaveMappingXmlRecord()`
- `ReadXMLMappingData()`
- `DataSetting()`
- BackgroundWorker
- report generation
- Excel export
- student filtering
- Tag ID conversion
- report statistics

The scope is only the reset behavior.

---

## Completion Record

After implementation and verification, update:

`新生入學統計調整0916.md`

Document:

1. Modified file(s).
2. Original reset behavior.
3. New reset behavior.
4. Removal of old Config reload from reset.
5. Removal/avoidance of legacy UDT reset dependency if applicable.
6. Three DataGridViews cleared together.
7. Student Tag Sources refreshed during reset.
8. Exact same-name matching behavior.
9. Confirmation that Config is not immediately overwritten during reset.
10. Validation results.
11. Confirmation that print/report logic was not intentionally changed.

---

## Acceptance Criteria

- [ ] Reset clears X1/X2/X3 rows.
- [ ] Reset reloads the latest Student Tag Source options.
- [ ] Source lists do not duplicate after repeated resets.
- [ ] X1 predefined Targets are restored.
- [ ] X2 predefined Targets are restored.
- [ ] X3 predefined Targets are restored.
- [ ] Exact same-name Source is automatically filled.
- [ ] No same-name Source leaves Source blank.
- [ ] Reset does not restore old Config mappings.
- [ ] Reset does not call `LoadConfigXml()`.
- [ ] Reset does not immediately save/overwrite Config.
- [ ] Existing print/save/report logic remains unchanged.
- [ ] Changes are documented in `新生入學統計調整0916.md`.
