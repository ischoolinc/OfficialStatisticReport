# TODO - 新生入學統計輸出資料表結構調整

## Goal

Modify the report export structure in `Form2.cs`.

The report must use the complete workbook template:

`新生入學方式統計表_樣板.xlsx`

The output workbook must keep and fill all 6 worksheets:

1. `普通科`
2. `專業群科(職業科)`
3. `綜合高中`
4. `實用技能學程`
5. `進修部(學校)`
6. `異常資料表`

Do not collapse the report into a single worksheet.

After implementation, document the changes in:

`新生入學統計調整0916.md`

---

## Current Problem

Current `Export()` does not use the complete template workbook directly.

Current pattern:

```csharp
_wk = new Workbook();
_wk.Worksheets.Add();

ws = _wk.Worksheets[1];
ws.Name = "異常資料表";
```

Then the template is opened separately:

```csharp
Workbook wk2 = new Workbook();

wk2.Open(
    new MemoryStream(
        Properties.Resources.新生入學方式統計表_樣板
    )
);
```

But only the first worksheet is copied:

```csharp
_wk.Worksheets[0].Copy(wk2.Worksheets[0]);
```

and then renamed:

```csharp
ws = _wk.Worksheets[0];
ws.Name = "新生入學方式統計表";
```

This means the current output does NOT preserve the 5 department-specific worksheets from the template.

---

# Required New Export Structure

## 1. Load the Complete Template Workbook Directly

Do not create a new blank workbook and copy only one worksheet.

Preferred structure:

```csharp
_wk = new Workbook();

_wk.Open(
    new MemoryStream(
        Properties.Resources.新生入學方式統計表_樣板
    )
);
```

After loading, use the existing worksheet names directly:

```csharp
_wk.Worksheets["普通科"];
_wk.Worksheets["專業群科(職業科)"];
_wk.Worksheets["綜合高中"];
_wk.Worksheets["實用技能學程"];
_wk.Worksheets["進修部(學校)"];
_wk.Worksheets["異常資料表"];
```

Do not add replacement worksheets, copy only worksheet 0, rename template worksheets, or remove existing template worksheets.

---

# Student Data Requirements

The existing student query must be extended so each student has:

```text
部別
科別代碼
科別名稱
班別
ref_class_id
Student Tag IDs
```

## 2. Department / Subject Data Source

Keep the current subject resolution priority:

```text
student.ref_dept_id has value
    -> use student.ref_dept_id

student.ref_dept_id is null
    -> use class.ref_dept_id
```

Then join `dept` and read:

```sql
dept.code AS dept_code,
dept.name AS dept_name
```

## 3. Department Group Data Source

Use `dept_group` as the department-group source.

Join:

```sql
LEFT JOIN dept_group
    ON dept.ref_dept_group_id = dept_group.id
```

Read:

```sql
dept_group.name AS dept_group_name
```

Reference query:

```sql
select
    dept.name AS 科別名稱,
    dept_group.name AS 部別名稱
from dept
left join dept_group
    ON dept.ref_dept_group_id = dept_group.id;
```

---

# Worksheet Classification

## 4. Route Students by `dept_group.name`

Students must be classified into one of these 5 report worksheets:

```text
普通科
專業群科(職業科)
綜合高中
實用技能學程
進修部(學校)
```

Expected routing:

```text
dept_group.name
    ↓
matching worksheet name
    ↓
write statistics to that worksheet
```

Use exact worksheet-category matching unless the existing database requires an explicit mapping table.

If a student cannot be mapped to one of the 5 supported report worksheets, treat the student as abnormal and write it to `異常資料表`.

Do not silently place unknown department groups into another worksheet.

---

# Grouping Rules

## 5. Group by Department + Subject + Class Type

Statistics must first be separated by:

```text
dept_group_name
    ↓
dept_code
    ↓
dept_name
    ↓
ClassType
```

`ClassType` continues to come from the freshman update record:

```xml
ContextInfo/ClassType
```

Current SQL already extracts it through:

```sql
TRIM(update_record_info.Class_Type) AS Class_Type
```

Keep this data source.

---

# Sorting Rule

## 6. Sort by Subject Code

Within each worksheet, groups must be sorted by:

```text
dept.code
```

Use:

```sql
dept.code AS dept_code
```

or sort the grouped result in C# before writing.

Do NOT rely on `ORDER BY dept_name` or Dictionary enumeration order.

The final report order must explicitly follow the subject code.

If multiple groups have the same subject code but different ClassType, keep a stable secondary order such as `ClassType`.

---

# Admission Type / Identity Statistics

## 7. Keep Existing Student Tag Mapping Logic

Do not replace the existing Tag mapping architecture.

Current flow must remain conceptually:

```text
DataGridView Target / Source
        ↓
ReadXMLMappingData()
        ↓
Source name -> tag.id
        ↓
XML_mappingData
        ↓
student tag_student.ref_tag_id
        ↓
filter.getListByTagId(...)
```

Use student Tags to calculate the existing 11 admission methods and 4 admission identities.

Do not change the existing mapping definitions in this task.

---

# Actual Class Count

## 8. `實際招生班數` Uses `ref_class_id`

For this version, calculate actual class count from students in the same report group.

Rule:

```text
same dept_group
+ same dept
+ same ClassType
        ↓
collect ref_class_id
        ↓
exclude null / empty
        ↓
Distinct
        ↓
Count
```

Example:

```text
Student A -> ref_class_id = 1001
Student B -> ref_class_id = 1001
Student C -> ref_class_id = 1002
Student D -> ref_class_id = 1002
```

Expected:

```text
實際招生班數 = 2
```

If the existing `Filter.getClassCount()` already follows this rule, it may be reused.

If not, modify only the necessary logic so the result is based on distinct non-empty `ref_class_id`.

Do not introduce a new class-table based count in this task.

---

# Normal Statistics Per Group

## 9. Each Department/Subject/ClassType Group Must Calculate

For every output row/group, calculate at minimum:

```text
科別代碼
科別名稱
班別
實際招生班數
新生總計
男生
女生
入學方式統計
入學身分統計
```

Use the existing report calculation rules for the detailed admission method / identity cells unless a worksheet layout requires a different destination cell.

Do not change business rules unrelated to the worksheet split.

---

# Workbook Writing Strategy

## 10. Each Worksheet Needs Its Own Row Index

Do not use one global row index for all department types.

Each normal worksheet should independently start from the template's detail row.

Current code uses:

```csharp
index = 12;
```

Aspose.Cells is zero-based, so this corresponds to Excel row 13.

Recommended concept:

```text
普通科              -> own index
專業群科(職業科)    -> own index
綜合高中            -> own index
實用技能學程        -> own index
進修部(學校)        -> own index
```

Each worksheet's data rows must progress independently.

---

# Error Worksheet

## 11. Use the Template `異常資料表`

Do not create a new `異常資料表`.

Use:

```csharp
Worksheet errorSheet =
    _wk.Worksheets["異常資料表"];
```

Write abnormal students into the existing template worksheet.

Keep the existing abnormal-data concept:

```csharp
filter.error_list
```

and extend abnormal handling as needed for the new worksheet classification.

At minimum, treat data as abnormal when a student cannot safely participate in normal worksheet statistics because required classification data is missing/invalid, for example:

```text
cannot determine dept
dept code missing
dept group missing
dept group cannot map to one of the 5 supported worksheets
required freshman update/class type data invalid
```

Do not include abnormal students in normal worksheet statistics if the data needed for classification is invalid.

---

# Recommended Export Flow

Refactor `Export()` conceptually into:

```text
Export()
    ↓
Load complete template workbook
    ↓
Get worksheet references
    ↓
Prepare Tag ID mappings
    ↓
Prepare normal students
    ↓
Validate department / department group data
    ↓
Invalid students -> 異常資料表
    ↓
Valid students
    ↓
Group by dept_group + dept_code + dept_name + ClassType
    ↓
Sort by dept_code
    ↓
For each dept_group:
        select corresponding worksheet
        calculate statistics
        write rows
    ↓
Write other existing summary sections as required
    ↓
Set SchoolYear on required worksheets
    ↓
Save the entire workbook once
```

---

# Suggested Helper Methods

Prefer small helper methods where practical, for example:

```csharp
private Worksheet GetReportWorksheet(string deptGroupName)
```

```csharp
private int GetActualClassCount(List<myStudent> students)
```

```csharp
private void WriteErrorSheet(...)
```

```csharp
private void WriteDepartmentSheet(...)
```

Helper names may be adjusted to fit the existing project style.

Do not over-refactor unrelated code.

---

# Preserve Existing Behavior

Do not intentionally change:

- `buttonX1_Click()` event flow
- `SaveMappingXmlRecord()`
- `ReadXMLMappingData()`
- reset behavior
- Config mapping behavior
- Student Tag source logic
- `CheckStudentStatus()`
- freshman update code `< 100` rule
- gender statistics
- admission method definitions
- admission identity definitions
- save-file dialog flow
- output/open-file behavior

The main scope is:

```text
report workbook structure
+ department group classification
+ subject code sorting
+ per-worksheet output
+ error worksheet routing
```

---

# Validation

## Test 1 - Workbook Structure

Generate the report.

Expected worksheets:

```text
普通科
專業群科(職業科)
綜合高中
實用技能學程
進修部(學校)
異常資料表
```

No worksheet should be renamed to `新生入學方式統計表`.

No extra replacement `異常資料表` should be created.

## Test 2 - 普通科

Students with:

```text
dept_group.name = 普通科
```

must only be written to `普通科`.

## Test 3 - 專業群科(職業科)

Students with:

```text
dept_group.name = 專業群科(職業科)
```

must be written to `專業群科(職業科)`.

## Test 4 - Other Department Groups

Verify `綜合高中`, `實用技能學程`, and `進修部(學校)` each writes only to its matching worksheet.

## Test 5 - Subject Code Sorting

Prepare subject codes:

```text
305
101
220
```

Expected output order:

```text
101
220
305
```

## Test 6 - ClassType

Students in the same subject but with different freshman update `ClassType` values must be grouped separately if the report structure requires separate rows.

The ClassType value must continue to come from freshman update data.

## Test 7 - Actual Class Count

Students:

```text
ref_class_id = 1001
ref_class_id = 1001
ref_class_id = 1002
ref_class_id = 1002
```

Expected:

```text
實際招生班數 = 2
```

## Test 8 - Tag Statistics

Verify admission method and admission identity counts remain consistent with current Tag mapping rules.

## Test 9 - Abnormal Department Group

If:

```text
dept_group.name = null
```

or a value is not mapped to the 5 supported worksheets:

Expected:

- student appears in `異常資料表`
- student is not included in normal worksheet statistics

## Test 10 - Complete Workbook Save

Press Print once.

Expected:

```text
one Excel file
```

containing all 6 worksheets with their respective data.

Do not create separate files for each department group.

---

# Completion Record

After implementation, update:

`新生入學統計調整0916.md`

Record:

1. Modified file(s).
2. Original workbook export behavior.
3. New complete-template loading behavior.
4. Six preserved worksheet names.
5. `dept_group` join and department classification.
6. `dept.code` source and sorting rule.
7. Subject resolution priority: student department first, class department fallback.
8. Freshman update `ClassType` source.
9. Actual class count using distinct `ref_class_id`.
10. Student Tag based admission method / identity calculation.
11. Error worksheet routing rules.
12. Validation results.
13. Confirmation that Config/reset/print entry logic was not intentionally changed.

---

# Acceptance Criteria

- [ ] Complete template workbook is loaded directly.
- [ ] All 6 template worksheets are preserved.
- [ ] No single generic report worksheet replaces the 5 category worksheets.
- [ ] Department group comes from `dept_group.name`.
- [ ] Subject code comes from `dept.code`.
- [ ] Subject name comes from `dept.name`.
- [ ] Student department overrides class department when available.
- [ ] Students are routed to the correct worksheet by department group.
- [ ] Unknown/unusable department-group data goes to `異常資料表`.
- [ ] Data inside each worksheet is sorted by subject code.
- [ ] ClassType continues to come from freshman update data.
- [ ] Admission method and admission identity continue to use Student Tags.
- [ ] Actual class count uses distinct non-empty `ref_class_id`.
- [ ] Each worksheet has an independent output row index.
- [ ] One Print action outputs one workbook containing all worksheets.
- [ ] Existing Config/reset logic is not intentionally changed.
- [ ] Changes are documented in `新生入學統計調整0916.md`.
