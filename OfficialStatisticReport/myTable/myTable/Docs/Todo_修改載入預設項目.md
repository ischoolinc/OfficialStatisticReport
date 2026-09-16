# TODO - 新生入學統計載入預設選項調整

## Goal

Modify the default option loading behavior in `Form2.cs`.

When the form is loaded, all predefined target options must always be displayed in:

- `dataGridViewX1` - 入學方式
- `dataGridViewX2` - 入學身分
- `dataGridViewX3` - 新生中具原住民身分者

If an existing Config mapping contains a matching `target`, load its `source`.
If no matching mapping exists, keep the `source` cell blank.

After completing the modification, document the changes in:

`新生入學統計調整0915.md`

---

## Current Problem

The current `LoadConfigXml()` behavior depends on the number of `<item>` elements stored in Config.

If Config already exists but contains only part of the predefined target options, the DataGridView will only display those saved items.

This causes predefined options to disappear from the UI.

Example:

If `dataGridViewX1` has 11 predefined admission method targets, but Config only contains 3 items, only those 3 rows are currently displayed.

Expected behavior:

- Always display all 11 predefined targets.
- Existing Config only supplies the corresponding `source`.
- Missing mappings must remain blank.

The same behavior must apply to `dataGridViewX2` and `dataGridViewX3`.

---

## Required Changes

### 1. Modify `LoadConfigXml()`

Change the loading flow to:

1. Clear existing rows from all three DataGridViews.
2. Build all rows from the predefined target lists first.
3. Read the existing Config.
4. Match Config items by exact `target`.
5. If a matching Config item exists:
   - assign its `source` to column 1.
6. If no matching Config item exists:
   - leave column 1 blank.

Config must NOT determine how many target rows are displayed.

---

## 2. `dataGridViewX1` - 入學方式

The target list comes from:

`Column2.Items`

Always display all 11 predefined targets:

1. `入學方式:免試入學--校內直升`
2. `入學方式:免試入學--優先免試`
3. `入學方式:免試入學--完全免試`
4. `入學方式:免試入學--就學區免試(含共同就學區)`
5. `入學方式:免試入學--技優甄審`
6. `入學方式:免試入學--免試獨招`
7. `入學方式:免試入學--其他`
8. `入學方式:特色招生--考試分發`
9. `入學方式:特色招生--甄選入學`
10. `入學方式:適性輔導安置(十二年安置)`
11. `入學方式:其他`

Required behavior:

```text
Target                                      Source
-------------------------------------------------------------
入學方式:免試入學--校內直升                existing source or blank
入學方式:免試入學--優先免試                existing source or blank
入學方式:免試入學--完全免試                existing source or blank
...
入學方式:其他                               existing source or blank
```

Do not remove a target just because Config does not contain that target.

---

## 3. `dataGridViewX2` - 入學身分

The target list comes from:

`dataGridViewComboBoxExColumn1.Items`

Always display all 4 predefined targets:

1. `入學身份:一般生(非外加錄取)`
2. `入學身份:外加錄取--原住民生`
3. `入學身份:外加錄取--身心障礙生`
4. `入學身份:外加錄取--其他`

Use the same loading logic as `dataGridViewX1`.

Existing Config:

```xml
<入學身分>
    <item target="..." source="..." />
</入學身分>
```

Match by `target`.

- Match found -> load `source`.
- No match -> blank source.

---

## 4. `dataGridViewX3` - 新生中具原住民身分者

The target list comes from:

`dataGridViewComboBoxExColumn3.Items`

Always display:

`新生中具原住民身分者`

Use the same loading logic as `dataGridViewX1` and `dataGridViewX2`.

Existing Config:

```xml
<新生中具原住民身分者>
    <item target="新生中具原住民身分者" source="..." />
</新生中具原住民身分者>
```

- Match found -> load `source`.
- No match -> blank source.

---

## 5. Clear DataGridView Rows Before Reload

Before rebuilding the rows in `LoadConfigXml()`, clear:

```csharp
dataGridViewX1.Rows.Clear();
dataGridViewX2.Rows.Clear();
dataGridViewX3.Rows.Clear();
```

This is required because `LoadConfigXml()` can be called again by the reset/reload flow.

Avoid duplicate rows, especially in `dataGridViewX2` and `dataGridViewX3`.

---

## 6. Fix Initial Config Creation

Review the `config == null` branch in `LoadConfigXml()`.

The initial Config creates the `入學方式` element and its 11 items.

Ensure it is actually appended to the Config root:

```csharp
config.AppendChild(EnterSchool_Way);
```

The resulting Config should contain all three sections:

```xml
<新生入學統計報表_來源目標設定Config>
    <入學方式>
        ...
    </入學方式>

    <入學身分>
        ...
    </入學身分>

    <新生中具原住民身分者>
        ...
    </新生中具原住民身分者>
</新生入學統計報表_來源目標設定Config>
```

Do not change the existing Config key:

`新生入學統計報表_來源目標設定Config`

---

## 7. Preserve Existing Source Data

Do not replace or clear valid existing Config mappings.

Example Config:

```xml
<item
    target="入學方式:免試入學--校內直升"
    source="入學方式:校內直升" />
```

The UI must display:

```text
標記：入學方式:免試入學--校內直升
來源：入學方式:校內直升
```

For a predefined target not found in Config:

```text
標記：入學方式:免試入學--優先免試
來源：
```

---

## 8. Do Not Change Other Logic

Do NOT modify unrelated business logic.

In particular, preserve the existing behavior of:

- `Column2Prepare()`
- `dataGridViewComboBoxExColumn2Prepare()`
- `dataGridViewComboBoxExColumn3Prepare()`
- `Column3Prepare()`
- `SaveMappingXmlRecord()`
- `ReadXMLMappingData()`
- `SetXMLMappingDataKey()`
- student Tag ID mapping
- report statistics
- Excel export
- admission method statistics
- admission identity statistics
- Aboriginal identity statistics
- student filtering logic

The purpose of this change is only to fix the UI loading/default option behavior.

---

## Suggested Implementation

A safe implementation is:

1. Read Config sections.
2. Convert each section into a `Dictionary<string, string>`:
   - key = `target`
   - value = `source`
3. Loop through each predefined target collection.
4. Create one DataGridView row per predefined target.
5. Set:
   - `Cells[0]` = predefined target
   - `Cells[1]` = mapped source if found, otherwise `""`

Conceptual example:

```csharp
Dictionary<string, string> mappings = new Dictionary<string, string>();

foreach (XmlElement item in EnterSchool_Way.SelectNodes("item"))
{
    string target = item.HasAttribute("target")
        ? item.GetAttribute("target")
        : "";

    string source = item.HasAttribute("source")
        ? item.GetAttribute("source")
        : "";

    if (!string.IsNullOrEmpty(target))
    {
        mappings[target] = source;
    }
}

foreach (string target in Column2.Items)
{
    DataGridViewRow row = new DataGridViewRow();
    row.CreateCells(dataGridViewX1);

    row.Cells[0].Value = target;

    if (mappings.ContainsKey(target))
        row.Cells[1].Value = mappings[target];
    else
        row.Cells[1].Value = "";

    dataGridViewX1.Rows.Add(row);
}
```

Apply the same pattern to all three DataGridViews.

Avoid introducing a new architecture unless necessary.

---

## Validation

After modification, verify all of the following.

### Test 1 - No existing Config

Open the form with no previous Config.

Expected:

- `dataGridViewX1` = 11 target rows
- `dataGridViewX2` = 4 target rows
- `dataGridViewX3` = 1 target row
- Source values may be blank.

### Test 2 - Complete Config

Config contains all targets and sources.

Expected:

- all predefined targets are displayed;
- all matching sources are restored correctly.

### Test 3 - Partial Config

Example:

`dataGridViewX1` Config only contains 3 of the 11 targets.

Expected:

- UI still displays all 11 rows;
- the 3 matched rows contain their existing sources;
- the other 8 rows have blank sources.

### Test 4 - Reload / Reset

Trigger the existing reload/reset function that calls `LoadConfigXml()` again.

Expected:

- row counts remain:
  - X1 = 11
  - X2 = 4
  - X3 = 1
- no duplicate rows are added.

### Test 5 - Save and Reopen

1. Select several source mappings.
2. Save/print using the existing flow.
3. Close and reopen the form.

Expected:

- all predefined targets are still displayed;
- saved sources are restored to the correct targets;
- unsaved/unmapped targets remain blank.

### Test 6 - Report Regression

Generate the 新生入學方式統計表.

Expected:

- existing report generation remains functional;
- existing Tag ID mapping remains functional;
- no changes to statistics caused by the UI loading modification.

---

## Completion Record

After implementation and verification, create/update:

`新生入學統計調整0915.md`

Document:

1. Files modified.
2. Original problem.
3. Root cause.
4. Changes made to `LoadConfigXml()`.
5. Default target row counts:
   - `dataGridViewX1`: 11
   - `dataGridViewX2`: 4
   - `dataGridViewX3`: 1
6. Config `target` -> `source` matching behavior.
7. Missing source behavior.
8. DataGridView row clearing behavior.
9. Initial Config `入學方式` node fix.
10. Regression tests performed.
11. Confirmation that report/statistics logic was not intentionally changed.

---

## Acceptance Criteria

The task is complete only when:

- [ ] `dataGridViewX1` always displays all 11 predefined targets.
- [ ] `dataGridViewX2` always displays all 4 predefined targets.
- [ ] `dataGridViewX3` always displays its predefined target.
- [ ] Existing source mappings are restored by exact target matching.
- [ ] Targets without mappings show a blank source.
- [ ] Reloading does not create duplicate rows.
- [ ] Initial Config correctly contains the `入學方式` section.
- [ ] Existing save/mapping/report logic continues to work.
- [ ] No unrelated logic is modified.
- [ ] Changes are documented in `新生入學統計調整0915.md`.
