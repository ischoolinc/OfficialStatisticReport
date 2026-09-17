# TODO - 新生入學統計新版樣板列印位置調整

## Goal

Adjust `Form2.cs` report output positions so the generated report matches the new workbook template:

`新生入學方式統計表_樣板.xlsx`

The original filling logic was based on:

`template(112.7 ver).xls`

The new template removes some sections and moves later sections upward.

After implementation, document all changes in:

`新生入學統計調整0917.md`

---

# Completion Status

**Status: DONE**（2025/09/17）

Implemented in `Form2.WriteDepartmentSheet`:

- Disabled removed sections: 按國中畢/修業年度分、新生中具原住民身分者
- Moved 按戶籍地分 Aspose 42/43 → 39/40
- Moved 按畢業國中所在地分 Aspose 45/46 → 42/43
- Fixed 澎湖縣 column 46 → 36 (AK); 其他 remains 46 (AU)
- Added detail-row overflow guard (Aspose 12~31 / Excel 13~32)
- Documented in `新生入學統計調整0917.md` §12

---

# Source of Truth

Use the new workbook layout as the final output target:

`新生入學方式統計表_樣板.xlsx`

Preserve existing business/statistical logic where the corresponding field still exists in the new template.

If a field/section existed in the old template but no longer exists in the new template, disable the related output code instead of writing into another location.

---

# Template Difference Summary

## Old Template Layout

Old `template(112.7 ver).xls`:

```text
Excel row 13~32   科別/班別明細
Excel row 36~39   按入學身分分
Excel row 40~42   按國中畢/修業年度分
Excel row 43~44   按戶籍地分
Excel row 45~47   按畢業國中所在地分
Excel row 48~51   新生中具原住民身分者
Excel row 52+     附註
```

## New Template Layout

New `新生入學方式統計表_樣板.xlsx`:

```text
Excel row 13~32   科別/班別明細
Excel row 36~39   按入學身分分
Excel row 40~41   按戶籍地分
Excel row 42~44   按畢業國中所在地分
Excel row 45+     附註
```

The new template has removed:

```text
按國中畢/修業年度分
新生中具原住民身分者
```

---

# Important Aspose Row Rule

Aspose.Cells uses zero-based indexes.

Examples:

```text
Aspose row 12 = Excel row 13
Aspose row 35 = Excel row 36
Aspose row 39 = Excel row 40
```

Do not confuse Excel row numbers with Aspose row indexes.

---

# 1. Keep Existing Detail Section

Current detail output starts with:

```csharp
int index = 12;
```

This corresponds to:

```text
Excel row 13
```

Keep the current detail-row start.

Current important fields remain valid:

```csharp
cs[index, 1]  // B 科別代碼
cs[index, 2]  // C 科別名稱
cs[index, 3]  // D 班別
cs[index, 6]  // G 實際招生班數
cs[index, 7]  // H 新生總計
cs[index, 8]  // I 男
cs[index, 9]  // J 女
```

Do not move these columns.

Do not change the existing admission-method / admission-identity column mapping in this task.

---

# 2. Keep `按入學身分分` Position

Current code uses:

```csharp
cs[35, ...]
cs[36, ...]
cs[37, ...]
cs[38, ...]
```

These correspond to:

```text
Excel row 36
Excel row 37
Excel row 38
Excel row 39
```

This section still exists in the new template at the same location.

Keep these row indexes unchanged.

---

# 3. Disable `按國中畢/修業年度分`

The old template had:

```text
Excel row 40 當年畢業
Excel row 41 當年修業
Excel row 42 其他(含領結業證書)
```

The new template no longer has this section.

Current code still writes:

```csharp
cs[39, ...]
cs[40, ...]
cs[41, ...]
```

These now overlap the new template's other sections.

Disable this entire old section.

Target region:

```csharp
#region 按國中畢/修業年度分
...
#endregion
```

Prefer disabling/commenting the entire obsolete block, including its data preparation, not only the final `PutValue()` calls.

This includes logic related to:

```text
collect__LastGrade
collect__LastComplete
collect__LastOther

UpdateRecord.SelectByStudentIDs(...)
SHBeforeEnrollment.SelectByStudentIDs(...)

studentBeforeStatusMap

cs[39, ...]
cs[40, ...]
cs[41, ...]
```

Do not write these old statistics anywhere else.

---

# 4. Disable `新生中具原住民身分者` Old Output Section

The old template had a dedicated section around:

```text
Excel row 48~51
```

The new template no longer contains this section.

Disable the old output code including:

```csharp
cs[48, 6]...
cs[49, 6]...
cs[50, 6]...

cs[48, 7]...
cs[48, 8]...

cs[49, 7]...
cs[49, 8]...

cs[50, 7]...
cs[50, 8]...
```

These are part of the obsolete graduation/original-identity statistics block.

Do not relocate these values into another section.

---

# 5. Move `按戶籍地分` Up

## Old Position

Current code uses:

```csharp
cs[42, ...]
cs[43, ...]
```

These correspond to:

```text
Excel row 43
Excel row 44
```

That was correct for the old template.

## New Position

The new template expects:

```text
Excel row 40 本縣市
Excel row 41 其他縣市
```

Therefore change:

```text
Aspose row 42 -> 39
Aspose row 43 -> 40
```

Apply this change to the entire `按戶籍地分` section.

Examples:

```csharp
// Old
cs[42, 6].PutValue(collect__LocalCounty.Count);
cs[43, 6].PutValue(collect__OtherCounty.Count);

// New
cs[39, 6].PutValue(collect__LocalCounty.Count);
cs[40, 6].PutValue(collect__OtherCounty.Count);
```

Gender fields:

```csharp
// Old
cs[42, 7]
cs[42, 8]
cs[43, 7]
cs[43, 8]

// New
cs[39, 7]
cs[39, 8]
cs[40, 7]
cs[40, 8]
```

Admission-method cross-statistics must also move:

```csharp
// Old
cs[42, col]
cs[42, col + flexInsex]
cs[43, col]
cs[43, col + flexInsex]

// New
cs[39, col]
cs[39, col + flexInsex]
cs[40, col]
cs[40, col + flexInsex]
```

Do not change the column calculation logic.

---

# 6. Move `按畢業國中所在地分` Up

## Old Position

Current code uses:

```csharp
cs[45, ...] // male
cs[46, ...] // female
```

These correspond to:

```text
Excel row 46
Excel row 47
```

This was correct for the old template.

## New Position

The new template has:

```text
Excel row 42 = section/header
Excel row 43 = male
Excel row 44 = female
```

Therefore change:

```text
Aspose row 45 -> 42
Aspose row 46 -> 43
```

Apply to all city/county values.

Example:

```csharp
// Old
cs[45, 6].PutValue(...);
cs[46, 6].PutValue(...);

// New
cs[42, 6].PutValue(...);
cs[43, 6].PutValue(...);
```

Repeat for every county/city column.

---

# 7. Fix Existing `澎湖縣` Column Bug

Current code incorrectly writes `澎湖縣` to column index `46`.

Aspose column index `46` is:

```text
AU
```

which is the `其他` column.

Later, `其他` writes to the same cell and overwrites the `澎湖縣` value.

The correct `澎湖縣` column is:

```text
AK
```

Aspose zero-based column index:

```text
36
```

Therefore change the new-template output to:

```csharp
cs[42, 36].PutValue(
    filter.getGenderCount(
        Collect_BeforeSchoolLocationList["澎湖縣"], "1"));

cs[43, 36].PutValue(
    filter.getGenderCount(
        Collect_BeforeSchoolLocationList["澎湖縣"], "0"));
```

Keep `其他` at:

```csharp
cs[42, 46]
cs[43, 46]
```

Do not allow both categories to use column `46`.

---

# 8. Keep School Year Position

Current:

```csharp
cs["U5"].PutValue(_SchoolYear);
```

The new template still uses `U5` for the school year value.

Keep unchanged.

---

# 9. Do Not Change Workbook Structure

Keep the current complete-template loading:

```csharp
_wk = new Workbook();
_wk.Open(
    new MemoryStream(
        Properties.Resources.新生入學方式統計表_樣板
    )
);
```

Keep the existing worksheets:

```text
普通科
專業群科(職業科)
綜合高中
實用技能學程
進修部(學校)
異常資料表
```

Do not rename worksheets.

Do not recreate worksheets.

---

# 10. Detail Row Overflow Safety Check

The new template has the main detail area approximately:

```text
Excel row 13~32
```

which allows 20 detail rows.

Current code increments:

```csharp
index++;
```

for each:

```text
科別代碼 + 科別名稱 + 班別
```

group.

Add a safety check or at minimum diagnostic warning if detail rows exceed the template's intended detail range.

Do not silently overwrite:

```text
Excel row 33+
```

summary/header areas.

Preferred behavior:

- detect overflow before writing into the next report section
- log/report the condition
- do not corrupt summary sections

Do not redesign the template in this task.

---

# 11. Preserve Existing Business Logic

Do not intentionally change:

- department-group mapping
- department code
- department name
- ClassType logic
- actual class count
- gender count
- admission-method Tag mapping
- admission-identity Tag mapping
- Filter grouping
- performance optimizations
- abnormal student logic
- save/open workflow

This task is only:

```text
new template output position adjustment
+ disable removed old-template sections
+ fix 澎湖縣 column bug
```

---

# Validation

## Test 1 - Detail Section

Verify output still starts at:

```text
Excel row 13
```

and fills the correct columns:

```text
B 科別代碼
C 科別名稱
D 班別
G 實際招生班數
H 新生總數
I 男
J 女
K~AZ admission-method / identity data
```

---

## Test 2 - 入學身分

Verify:

```text
Excel row 36~39
```

still receives:

```text
一般生
原住民生
身心障礙生
其他
```

No row change.

---

## Test 3 - Removed Graduation Section

Verify no code writes the old:

```text
當年畢業
當年修業
其他(含領結業證書)
```

section.

No writes should remain at old Aspose rows:

```text
39
40
41
```

for this obsolete purpose.

---

## Test 4 - 戶籍地

Verify new output:

```text
Excel row 40 = 本縣市
Excel row 41 = 其他縣市
```

with correct total, male/female, and admission-method statistics.

---

## Test 5 - 畢業國中所在地

Verify:

```text
Excel row 43 = male
Excel row 44 = female
```

and all city/county columns align with the new template.

---

## Test 6 - 澎湖縣

Verify:

```text
澎湖縣 -> AK
其他   -> AU
```

and that neither overwrites the other.

---

## Test 7 - Removed Aboriginal Section

Verify the old dedicated:

```text
新生中具原住民身分者
```

graduation-status section is not written.

Do not relocate it.

---

## Test 8 - School Year

Verify:

```text
U5
```

contains `_SchoolYear`.

---

## Test 9 - All 5 Report Worksheets

Generate report and verify the same position adjustment is applied to:

```text
普通科
專業群科(職業科)
綜合高中
實用技能學程
進修部(學校)
```

Each worksheet must use the new template row layout.

---

## Test 10 - No Summary Overwrite

Verify detail groups do not overwrite the summary section.

If detail groups exceed the intended detail range, the program must detect/log the condition instead of silently corrupting later sections.

---

# Completion Record

After implementation, update:

`新生入學統計調整0917.md`

Record:

1. Modified files.
2. Compared old/new template layout.
3. Removed old-template output sections:
   - 按國中畢/修業年度分
   - 新生中具原住民身分者
4. `按戶籍地分` row change:
   - Aspose 42/43 -> 39/40
5. `按畢業國中所在地分` row change:
   - Aspose 45/46 -> 42/43
6. `按入學身分分` remains unchanged.
7. Detail row start remains Aspose row 12.
8. `U5` school year remains unchanged.
9. 澎湖縣 column corrected to AK / Aspose column 36.
10. 其他 remains AU / Aspose column 46.
11. Detail-row overflow handling/check.
12. Validation results for all 5 normal worksheets.
13. Confirmation that business/statistical logic was not intentionally changed.

---

# Acceptance Criteria

- [x] New template is used as the final layout.
- [x] Detail output position remains correct.
- [x] 入學身分 section remains at Excel rows 36~39.
- [x] Old graduation-status section is disabled.
- [x] Old Aboriginal graduation-status section is disabled.
- [x] 戶籍地 output moves to Excel rows 40~41.
- [x] 畢業國中所在地 output moves to Excel rows 43~44.
- [x] 澎湖縣 writes to AK / Aspose column 36.
- [x] 其他 writes to AU / Aspose column 46.
- [x] U5 school year remains correct.
- [x] Same position logic applies to all 5 report worksheets.
- [x] No obsolete old-template values overwrite new-template fields.
- [x] Detail-row overflow is detected instead of silently overwriting summary sections.
- [x] Existing statistical/business rules remain unchanged.
- [x] Changes are documented in `新生入學統計調整0917.md`.
