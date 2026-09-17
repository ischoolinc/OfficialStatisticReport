# TODO - 新生入學統計人數統計方式調整

## Goal

Adjust the student-count classification logic used by the freshman admission statistics report.

Primary goal:

- Fix department-group classification so database department-group names are converted to the correct report worksheet names before validation and grouping.
- Prevent valid students from being incorrectly moved to `error_list`.
- Ensure students are counted in the correct worksheet.

After implementation, document the changes in:

`新生入學統計調整0917.md`

---

# Status

**Completed** (2025-09-17)

Implemented:

- `Filter.GetReportDeptGroupName()` mapping layer
- `Cleaner()` convert-then-validate; normalize `Dept_group_name` for clean students
- Diagnostic Debug / perf-log counts
- Documentation appended to `Docs/新生入學統計調整0917.md`

Pending user verification with live print:

- Test data counts: 普通科=8, 專業群科(職業科)=266, 異常=2

---

# Background

Current database query reads:

```sql
dept.code AS dept_code,
dept.name AS dept_name,
dept_group.name AS dept_group_name
```

Current real database values include:

```text
普通型高中
技術型高中
```

But the Excel report worksheet names are:

```text
普通科
專業群科(職業科)
綜合高中
實用技能學程
進修部(學校)
```

These are not the same naming system.

Current `Filter.cs` incorrectly assumes:

```text
dept_group.name
==
Excel worksheet name
```

This causes valid students to fail `IsSupportedDeptGroup()` and be placed into `error_list`.

---

# Current Problem

Current `Filter.cs` defines:

```csharp
public static readonly string[] SupportedDeptGroupNames = new string[]
{
    "普通科",
    "專業群科(職業科)",
    "綜合高中",
    "實用技能學程",
    "進修部(學校)"
};
```

Current validation:

```csharp
if (
    ...
    || !IsSupportedDeptGroup(s.Dept_group_name)
    ...
)
{
    error_list.Add(s);
}
```

Current check:

```csharp
public static bool IsSupportedDeptGroup(string deptGroupName)
{
    if (string.IsNullOrEmpty(deptGroupName))
        return false;

    return SupportedDeptGroupNames.Contains(deptGroupName);
}
```

Therefore:

```text
普通型高中
!=
普通科

技術型高中
!=
專業群科(職業科)
```

Valid students are incorrectly treated as abnormal.

---

# Required Mapping

Add an explicit mapping layer between:

```text
database dept_group.name
```

and:

```text
report worksheet name
```

Required known mappings:

```text
普通型高中
    -> 普通科

技術型高中
    -> 專業群科(職業科)
```

The following worksheet categories must remain supported:

```text
綜合高中
實用技能學程
進修部(學校)
```

Do not guess unknown database aliases.

If the database actually uses different names for these categories, add mappings only after confirming the real values.

---

# Implementation Requirement

## 1. Add Department Group Conversion Method

Add a centralized method in `Filter.cs`.

Example:

```csharp
private static string GetReportDeptGroupName(string deptGroupName)
{
    if (string.IsNullOrWhiteSpace(deptGroupName))
        return "";

    switch (deptGroupName.Trim())
    {
        case "普通型高中":
        case "普通科":
            return "普通科";

        case "技術型高中":
        case "專業群科(職業科)":
            return "專業群科(職業科)";

        case "綜合高中":
            return "綜合高中";

        case "實用技能學程":
            return "實用技能學程";

        case "進修部(學校)":
            return "進修部(學校)";

        default:
            return "";
    }
}
```

Method name may be adjusted to existing project conventions.

Important:

- Use one centralized mapping method.
- Do not scatter string replacements across multiple methods.
- Use exact matching.
- `Trim()` is acceptable.
- Do not use broad `Contains()` matching.

---

# 2. Change Cleaner Validation

Current validation checks raw:

```csharp
s.Dept_group_name
```

against report worksheet names.

Change the logic so the database value is converted first.

Concept:

```csharp
string reportDeptGroupName =
    GetReportDeptGroupName(s.Dept_group_name);
```

Then validate:

```csharp
if (
    s.Id == "" ||
    s.Name == "" ||
    (s.Gender != "0" && s.Gender != "1") ||
    s.Ref_class_id == "" ||
    s.Class_name == "" ||
    s.Grade_year == "" ||
    s.Dept_name == "" ||
    string.IsNullOrEmpty(s.Dept_code) ||
    string.IsNullOrEmpty(reportDeptGroupName) ||
    !ClassTypeCodeDic.ContainsKey(s.Class_Type)
)
{
    error_list.Add(s);
}
else
{
    s.Dept_group_name = reportDeptGroupName;
    clean_list.Add(s);
}
```

Important:

- Convert only after reading the database value.
- Store the converted report category back into `s.Dept_group_name` only for clean/valid students.
- Do not alter `Dept_name`.
- Do not alter `Dept_code`.
- Do not alter `Class_Type`.

---

# 3. Keep `SupportedDeptGroupNames` as Report Worksheet Names

Do not change:

```csharp
SupportedDeptGroupNames
```

into database names.

It should continue to represent the Excel/report categories:

```text
普通科
專業群科(職業科)
綜合高中
實用技能學程
進修部(學校)
```

This allows current `Form2.Export()` logic to continue using:

```csharp
foreach (string sheetName in Filter.SupportedDeptGroupNames)
{
    var groups = filter.GetSortedGroups(sheetName);

    WriteDepartmentSheet(
        _wk.Worksheets[sheetName],
        groups,
        ...
    );
}
```

No need to change worksheet names.

---

# 4. Keep Classify Structure

After `Cleaner()` converts:

```text
database department group
    ↓
report department group
```

`Classify()` can continue using:

```csharp
ByDeptGroup[s.Dept_group_name]
```

because `s.Dept_group_name` is already normalized.

Expected flow:

```text
DB: 普通型高中
    ↓
Cleaner mapping
    ↓
普通科
    ↓
clean_list
    ↓
Classify()
    ↓
ByDeptGroup["普通科"]
    ↓
GetSortedGroups("普通科")
    ↓
Excel worksheet "普通科"
```

and:

```text
DB: 技術型高中
    ↓
Cleaner mapping
    ↓
專業群科(職業科)
    ↓
clean_list
    ↓
Classify()
    ↓
ByDeptGroup["專業群科(職業科)"]
    ↓
GetSortedGroups("專業群科(職業科)")
    ↓
Excel worksheet "專業群科(職業科)"
```

---

# 5. Abnormal Department / Subject Data

Current SQL result contains students where:

```text
dept_group_name = empty
dept_code = empty
dept_name = empty
```

These students must remain abnormal.

Do not force them into a report worksheet.

They should remain in:

```text
error_list
```

and be written to:

```text
異常資料表
```

Current abnormal rules should remain:

```text
Dept_name empty
Dept_code empty
Dept_group cannot map
```

Do not add fallback logic that guesses a department.

---

# 6. Preserve Subject Selection Logic

Do not change the existing department resolution rule in `Form2.cs`.

Current logic:

```text
student.ref_dept_id has value
    -> use student.ref_dept_id

student.ref_dept_id is null
    -> use class.ref_dept_id
```

The SQL join logic must remain unchanged in this task.

This task only fixes the department-group-to-report-category mapping.

---

# 7. Preserve Counting Rules

Do not change:

- `getClassCount()`
- `getGenderCount()`
- `getListByTagId()`
- admission method statistics
- admission identity statistics
- Tag mapping
- `ClassTypeCodeDic`
- subject-code sorting
- worksheet layout
- report template structure
- report performance optimizations already implemented

---

# 8. Expected Result for Current Test Data

Based on current SQL test result:

```text
普通型高中 = 8 students
技術型高中 = 266 students
department/subject empty = 2 students
```

Expected classification after fix:

```text
普通型高中 8
    -> 普通科

技術型高中 266
    -> 專業群科(職業科)

部別/科別空白 2
    -> 異常資料表
```

Expected normal total:

```text
274 students
```

Expected abnormal total:

```text
2 students
```

Do not hardcode these counts.

They are only current test expectations.

---

# 9. Add Temporary Diagnostic Logging

For verification, add temporary or existing debug/performance logging around classification.

Record at least:

```text
Raw dept_group_name
Mapped report dept group
clean_list count
error_list count
ByDeptGroup count per worksheet
```

Example:

```text
DeptGroup raw: 普通型高中 -> 普通科
DeptGroup raw: 技術型高中 -> 專業群科(職業科)

普通科 students=8
專業群科(職業科) students=266
異常資料 students=2
```

Do not show MessageBox per student.

Use Debug output or existing performance log.

---

# Validation

## Test 1 - 普通型高中

Database:

```text
dept_group.name = 普通型高中
```

Expected:

```text
s.Dept_group_name after normalization = 普通科
```

Student must:

- enter `clean_list`
- appear in `ByDeptGroup["普通科"]`
- be written to worksheet `普通科`

---

## Test 2 - 技術型高中

Database:

```text
dept_group.name = 技術型高中
```

Expected:

```text
s.Dept_group_name after normalization = 專業群科(職業科)
```

Student must:

- enter `clean_list`
- appear in `ByDeptGroup["專業群科(職業科)"]`
- be written to worksheet `專業群科(職業科)`

---

## Test 3 - Empty Department Data

Student:

```text
Dept_group_name = ""
Dept_code = ""
Dept_name = ""
```

Expected:

- remains in `error_list`
- not included in normal statistics
- written to `異常資料表`

---

## Test 4 - Unknown Department Group

Example:

```text
dept_group.name = 未知部別
```

Expected:

```text
GetReportDeptGroupName(...) = ""
```

Student must:

- go to `error_list`
- not be silently assigned to another worksheet

---

## Test 5 - Count Verification

For the current test dataset, verify expected result:

```text
普通科 = 8
專業群科(職業科) = 266
異常資料表 = 2
```

Also verify:

```text
8 + 266 + 2 = 276
```

No student should disappear.

No student should be counted twice.

---

## Test 6 - Existing Report Rules

Verify no regression in:

- 科別代碼
- 科別名稱
- 班別
- 實際招生班數
- 新生總數
- 男生
- 女生
- 入學方式
- 入學身分
- 原住民統計

---

# Completion Record

After implementation, update:

`新生入學統計調整0917.md`

Record:

1. Modified files.
2. Root cause:
   - database `dept_group.name`
   - report worksheet name mismatch.
3. Original incorrect behavior.
4. New department-group mapping logic.
5. Mapping:
   - `普通型高中 -> 普通科`
   - `技術型高中 -> 專業群科(職業科)`
6. Cleaner validation change.
7. Classify behavior after normalization.
8. Error-list behavior for missing/unknown department groups.
9. Current test counts:
   - 普通科
   - 專業群科(職業科)
   - 異常資料表
10. Confirmation that existing Tag/count/report rules were not intentionally changed.
11. Validation results.

---

# Acceptance Criteria

- [x] Database department-group names are converted before validation.
- [x] `普通型高中` maps to `普通科`.
- [x] `技術型高中` maps to `專業群科(職業科)`.
- [x] `SupportedDeptGroupNames` continues to contain report worksheet names.
- [x] Valid students are no longer incorrectly placed in `error_list`.
- [x] Missing/unknown department data still goes to `error_list`.
- [x] `Classify()` groups normalized report department names correctly.
- [x] `GetSortedGroups()` continues to work without report worksheet changes.
- [ ] Current test data produces expected worksheet counts. *(待實機列印確認)*
- [ ] No student is lost or double-counted. *(待實機列印確認)*
- [x] Existing Tag/statistics/business logic remains unchanged.
- [x] Changes are documented in `新生入學統計調整0917.md`.
