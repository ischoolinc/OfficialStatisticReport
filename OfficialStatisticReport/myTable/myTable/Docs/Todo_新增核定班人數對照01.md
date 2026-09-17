# TODO - 新生入學統計新增核班人數對照

## Goal

Add approved class-count / approved student-count lookup to the freshman admission statistics report.

Data source:

`$campus.updaterecord.govapprovednumofclass`

Use the selected school year from the screen to load approved admission data once, then match report rows by:

```text
部別ID
科別代碼
科別名稱
```

When all 3 values match:

```text
classnum   -> 報表「總班數」
studentnum -> 報表「總學生數」
```

If no matching record exists, or the numeric value cannot be converted:

```text
總班數 = 0
總學生數 = 0
```

After implementation, document all changes in:

`新生入學統計調整0917.md`

---

# Data Source

Use:

```sql
SELECT *
FROM $campus.updaterecord.govapprovednumofclass;
```

Required columns:

```text
deptgroup
dept_code
dept_name
classnum
studentnum
schoolyear
```

Column meanings:

```text
deptgroup  = 部別ID
dept_code  = 科別代碼
dept_name  = 科別名稱
classnum   = 核定班數
studentnum = 核定人數
schoolyear = 學年度
```

---

# 1. Add `dept_group.id` to Main Student SQL

Current student SQL already reads:

```sql
dept.code AS dept_code,
dept.name AS dept_name,
dept_group.name AS dept_group_name
```

Add:

```sql
dept_group.id AS dept_group_id
```

Example:

```sql
SELECT
    ...
    dept.code AS dept_code,
    dept.name AS dept_name,
    dept_group.id AS dept_group_id,
    dept_group.name AS dept_group_name,
    ...
```

Do not remove the existing `dept_group.name`.

The report still needs `dept_group.name` for worksheet classification.

---

# 2. Add `Dept_group_id` to `myStudent`

Add a property:

```csharp
public string Dept_group_id { get; set; }
```

When creating `myStudent`, read:

```csharp
string dept_group_id =
    row["dept_group_id"].ToString();
```

Then assign:

```csharp
studentObj.Dept_group_id = dept_group_id;
```

Important:

`Dept_group_id` must keep the original database ID.

Do NOT replace it with the normalized worksheet/report category name.

---

# 3. Add Approved Admission Data Model

Add a simple class such as:

```csharp
class GovApprovedAdmissionInfo
{
    public string DeptGroupId { get; set; }
    public string DeptCode { get; set; }
    public string DeptName { get; set; }

    public int ClassNum { get; set; }
    public int StudentNum { get; set; }
}
```

---

# 4. Add Key Builder

Use one centralized function for matching:

```csharp
private string BuildGovApprovedKey(
    string deptGroupId,
    string deptCode,
    string deptName)
{
    return
        (deptGroupId ?? "").Trim() + "⊕" +
        (deptCode ?? "").Trim() + "⊕" +
        (deptName ?? "").Trim();
}
```

Matching rule must be exact:

```text
部別ID
+ 科別代碼
+ 科別名稱
```

All 3 must match.

Do not use partial matching.

---

# 5. Add School-Year Query Function

Add a new function that receives the selected school year.

Example:

```csharp
private Dictionary<string, GovApprovedAdmissionInfo>
    LoadGovApprovedAdmissionInfo(string schoolYear)
```

Query only the selected school year:

```sql
SELECT
    deptgroup,
    dept_code,
    dept_name,
    classnum,
    studentnum,
    schoolyear
FROM $campus.updaterecord.govapprovednumofclass
WHERE schoolyear::text = '{0}'
ORDER BY
    deptgroup,
    dept_code,
    dept_name
```

Use:

```csharp
_SchoolYear
```

Do not hardcode a school year.

---

# 6. Convert `classnum` / `studentnum` to Integer

Use safe conversion:

```csharp
int classNum = 0;
int studentNum = 0;

int.TryParse(
    ("" + row["classnum"]).Trim(),
    out classNum);

int.TryParse(
    ("" + row["studentnum"]).Trim(),
    out studentNum);
```

Required behavior:

```text
NULL
empty
non-numeric
conversion failure
```

all result in:

```text
0
```

---

# 7. Return a Dictionary

Build:

```csharp
Dictionary<string, GovApprovedAdmissionInfo>
```

Key:

```text
deptgroup ⊕ dept_code ⊕ dept_name
```

If duplicate keys exist:

- do not use `Add()` in a way that throws
- use safe overwrite or explicit duplicate handling
- log duplicate keys for diagnostics

Preferred minimum-safe behavior:

```csharp
result[key] = info;
```

---

# 8. Load Approved Data Only Once Per Report

Do NOT query the approved table inside every report row.

Load once before looping through worksheets / groups:

```csharp
Dictionary<string, GovApprovedAdmissionInfo>
    govApprovedData =
        LoadGovApprovedAdmissionInfo(_SchoolYear);
```

Then pass the dictionary into:

```csharp
WriteDepartmentSheet(...)
```

---

# 9. Pass Approved Data into `WriteDepartmentSheet()`

Extend the method parameter list:

```csharp
private void WriteDepartmentSheet(
    Worksheet ws,
    List<KeyValuePair<string, List<myStudent>>> groups,
    ...
    Dictionary<string, GovApprovedAdmissionInfo> govApprovedData)
```

Do not introduce database access inside this method.

---

# 10. Match Approved Data Per Report Group

Current report group key is:

```text
科別代碼 ⊕ 科別名稱 ⊕ 班別
```

Extract:

```csharp
string deptCode =
    keyArray.Length >= 1 ? keyArray[0] : "";

string deptName =
    keyArray.Length >= 2 ? keyArray[1] : "";
```

Get `deptGroupId` from the group student data:

```csharp
string deptGroupId = "";

if (k.Value.Count > 0)
{
    deptGroupId =
        k.Value[0].Dept_group_id;
}
```

Build lookup key:

```csharp
string approvedKey =
    BuildGovApprovedKey(
        deptGroupId,
        deptCode,
        deptName);
```

---

# 11. Fill Report Columns E / F

Use default values:

```csharp
int approvedClassNum = 0;
int approvedStudentNum = 0;
```

Lookup:

```csharp
GovApprovedAdmissionInfo approvedInfo;

if (govApprovedData.TryGetValue(
        approvedKey,
        out approvedInfo))
{
    approvedClassNum =
        approvedInfo.ClassNum;

    approvedStudentNum =
        approvedInfo.StudentNum;
}
```

Write:

```csharp
// E：總班數 = 核定班數
cs[index, 4].PutValue(approvedClassNum);

// F：總學生數 = 核定人數
cs[index, 5].PutValue(approvedStudentNum);

// G：實際招生班數
cs[index, 6].PutValue(
    filter.getClassCount(k.Value));

// H：新生總計
cs[index, 7].PutValue(
    k.Value.Count);
```

Final meaning:

```text
E 總班數
= govapprovednumofclass.classnum

F 總學生數
= govapprovednumofclass.studentnum

G 實際招生班數
= current distinct ref_class_id count

H 新生總計
= current actual student count
```

---

# 12. No Match = 0

If no approved record matches:

```text
DeptGroupId
DeptCode
DeptName
```

then report must show:

```text
E = 0
F = 0
```

Do not leave blank cells.

Do not fallback to actual class/student counts.

---

# 13. Add Diagnostic Logging

For testing, log unmatched approved-data rows:

```csharp
if (!govApprovedData.TryGetValue(
        approvedKey,
        out approvedInfo))
{
    Debug.WriteLine(
        "核定招生資料未比對到：" +
        "DeptGroupId=" + deptGroupId +
        ", DeptCode=" + deptCode +
        ", DeptName=" + deptName);
}
```

Also log:

```text
selected school year
approved-data record count
matched count
unmatched count
```

Do not show one MessageBox per unmatched row.

---

# 14. Preserve Existing Department Group Mapping

Current report has department-group normalization such as:

```text
普通型高中
-> 普通科

技術型高中
-> 專業群科(職業科)
```

Do not use this normalized worksheet name for approved-data matching.

Approved-data matching must use:

```text
original dept_group.id
```

from the database.

That is why `Dept_group_id` must be preserved separately.

---

# 15. Preserve Existing Report Rules

Do not intentionally change:

- worksheet routing
- department-group name mapping
- subject-code sorting
- ClassType
- Tag mapping
- admission-method counts
- admission-identity counts
- actual class count
- new-student total
- gender counts
- abnormal-data handling
- new-template row positions
- performance optimizations
- workbook save logic

This task only adds approved class count and approved student count into E/F.

---

# Recommended Data Flow

```text
畫面選擇學年度
        ↓
_SchoolYear
        ↓
LoadGovApprovedAdmissionInfo(_SchoolYear)
        ↓
SELECT
deptgroup
dept_code
dept_name
classnum
studentnum
FROM $campus.updaterecord.govapprovednumofclass
WHERE schoolyear = selected year
        ↓
Dictionary
Key =
部別ID ⊕ 科別代碼 ⊕ 科別名稱
        ↓
WriteDepartmentSheet()
        ↓
取得：
Dept_group_id
Dept_code
Dept_name
        ↓
精確比對
        ↓
有資料：
E = classnum
F = studentnum

無資料：
E = 0
F = 0

G/H 維持既有實際統計
```

---

# Validation

## Test 1 - Matching Record

Approved-data row:

```text
deptgroup = 10
dept_code = 301
dept_name = 資訊科
classnum = 3
studentnum = 105
schoolyear = 115
```

Report group:

```text
Dept_group_id = 10
Dept_code = 301
Dept_name = 資訊科
```

Expected:

```text
總班數 = 3
總學生數 = 105
```

---

## Test 2 - School Year Filter

If screen selects:

```text
115
```

Only:

```text
schoolyear = 115
```

approved data may be used.

---

## Test 3 - No Matching Record

If no exact key exists:

```text
Dept_group_id
Dept_code
Dept_name
```

Expected:

```text
總班數 = 0
總學生數 = 0
```

---

## Test 4 - Null / Invalid Numeric Data

Examples:

```text
classnum = NULL
studentnum = ""
```

or:

```text
classnum = abc
```

Expected:

```text
0
```

No exception.

---

## Test 5 - Exact Three-Field Match

Verify that same `dept_code` + `dept_name` but different `deptgroup` does NOT match incorrectly.

All three values must match.

---

## Test 6 - Existing Actual Statistics

Verify unchanged:

```text
G 實際招生班數
H 新生總計
I 男
J 女
```

---

## Test 7 - Performance

Verify:

```text
$campus.updaterecord.govapprovednumofclass
```

is queried once per report generation.

---

# Completion Record

After implementation, update:

`新生入學統計調整0917.md`

Record:

1. Modified files.
2. Added `dept_group.id` to student SQL.
3. Added `Dept_group_id` to `myStudent`.
4. Added approved-admission data class.
5. Added approved-data key builder.
6. Added school-year based query function.
7. Data source:
   - `$campus.updaterecord.govapprovednumofclass`
8. Matching key:
   - `deptgroup`
   - `dept_code`
   - `dept_name`
9. Numeric conversion behavior.
10. No-match behavior:
    - E = 0
    - F = 0
11. E/F report mapping:
    - E = 核定班數 / 總班數
    - F = 核定人數 / 總學生數
12. Confirmation G/H logic remains unchanged.
13. Match/unmatched test results.
14. Confirmation approved data is queried only once.
15. Any duplicate-key handling.
16. Validation results.

---

# Acceptance Criteria

- [x] Main SQL includes `dept_group.id AS dept_group_id`.
- [x] `myStudent` keeps original `Dept_group_id`.
- [x] New function accepts selected school year.
- [x] Approved-data query filters by selected school year.
- [x] Approved data is loaded only once per report.
- [x] Matching uses exact `部別ID + 科別代碼 + 科別名稱`.
- [x] `classnum` converts to integer safely.
- [x] `studentnum` converts to integer safely.
- [x] Invalid/missing numeric values become 0.
- [x] No matching approved record results in E=0 and F=0.
- [x] Matching record writes E=核定班數.
- [x] Matching record writes F=核定人數.
- [x] G=實際招生班數 remains unchanged.
- [x] H=新生總計 remains unchanged.
- [x] Worksheet classification remains unchanged.
- [x] Existing Tag/statistics rules remain unchanged.
- [x] Changes are documented in `新生入學統計調整0917.md`.
