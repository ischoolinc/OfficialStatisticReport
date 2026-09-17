# TODO - 新生入學統計列印效能優化

## Goal

Optimize the very slow report generation flow in `Form2.cs`.

Target:

- Improve the performance of the `列印` / report generation process.
- Preserve the current report results and business rules.
- First add timing diagnostics so the actual bottleneck can be measured.
- Then apply low-risk performance optimizations to remove repeated database/API calls and repeated full-list scans.

After implementation, document all changes and test results in:

`新生入學統計調整0917.md`

---

# Scope

Primary target file:

`Form2.cs`

Main execution flow:

```text
buttonX1_Click()
    ↓
SaveMappingXmlRecord()
    ↓
ReadXMLMappingData()
    ↓
DataSetting()
    ↓
_BGWClassStudentAbsenceDetail_DoWork()
    ↓
main SQL
    ↓
build myDic / mylist
    ↓
new Filter(...)
    ↓
Export()
    ↓
WriteDepartmentSheet(...)
    ↓
_BGWClassStudentAbsenceDetail_Completed()
    ↓
Save XLSX
```

Do not change the report output rules unless required for performance only.

---

# Current Performance Risks

## 1. N+1 Query: `BeforeEnrollment.SelectByStudentID(id)`

Current code calls:

```csharp
K12.Data.BeforeEnrollmentRecord ber =
    K12.Data.BeforeEnrollment.SelectByStudentID(id);
```

inside:

```csharp
foreach (DataRow row in dt.Rows)
```

This is a major performance risk.

The main SQL joins:

```sql
tag_student
update_record_info
```

which can return multiple rows for the same student.

Therefore the same student may call:

```csharp
BeforeEnrollment.SelectByStudentID(id)
```

multiple times.

Example:

```text
800 students
x multiple Student Tags
= several thousand DataRows
```

This can cause thousands of repeated API/database calls.

---

## 2. Main SQL Can Multiply Rows

Current join structure includes:

```text
student
  ↓
tag_student (one-to-many)
  ↓
update_record_info (potentially one-to-many)
```

This can cause one student to appear in multiple SQL rows.

The current code correctly collapses students into:

```csharp
myDic[id]
```

but repeated per-row work is still executed before/around that collapse.

Avoid repeated expensive operations for duplicate student rows.

---

## 3. `CheckStudentStatus()` Repeated Full Scan

Current flow:

```csharp
foreach (KeyValuePair<string, myStudent> kvp in myDic)
{
    if (CheckStudentStatus(records, kvp.Key))
    {
        mylist.Add(kvp.Value);
    }
}
```

`CheckStudentStatus()` loops through the full UpdateRecord list for each student.

This can approach:

```text
students x update records
```

repeated comparisons.

Replace with a precomputed lookup such as:

```csharp
HashSet<string> validNewStudentIds
```

or:

```csharp
Dictionary<string, ...>
```

---

## 4. `recl × summary` Nested Loop

Current `WriteDepartmentSheet()` contains:

```csharp
foreach (SHSchool.Data.SHBeforeEnrollmentRecord rec in recl)
{
    foreach (myStudent student in summary)
    {
        if (rec.RefStudentID == student.Id)
        {
            ...
        }
    }
}
```

This is O(n²).

Replace the inner scan with a dictionary:

```csharp
Dictionary<string, myStudent> studentMap
```

and use:

```csharp
studentMap.TryGetValue(rec.RefStudentID, out student)
```

---

## 5. `CheckStudentBeforeStatus()` Called Repeatedly

Current code can call the same method up to 3 times for the same student:

```csharp
if (CheckStudentBeforeStatus(...) == "當年畢業")
...
if (CheckStudentBeforeStatus(...) == "當年修業")
...
if (CheckStudentBeforeStatus(...) == "其他(含領結業證書)")
...
```

The method itself scans all update records.

At minimum:

```csharp
string status = CheckStudentBeforeStatus(...);
```

then compare `status` once.

Preferred:

Precompute:

```csharp
Dictionary<string, string> studentBeforeStatusMap
```

and read directly.

---

## 6. Aboriginal Student List Recomputed Inside Loop

Current code computes:

```csharp
filter.getListByTagId(aboIDList, summary)
```

inside the `foreach (rec in recl)` loop.

This list does not depend on `rec`.

Move this calculation outside the loop.

Prefer also avoiding a second nested loop by using lookup dictionaries.

---

## 7. Repeated `getListByTagId()` / `getGenderCount()`

`WriteDepartmentSheet()` repeatedly scans the same student lists for:

- 11 admission methods
- 4 admission identities
- gender counts
- graduation status
- county statistics
- aboriginal statistics

Do not redesign all report logic in the first optimization pass unless needed.

First measure these calls.

If they are still a major bottleneck after the high-risk items above are fixed, optimize them in a second pass.

---

# Phase 1 - Add Timing Diagnostics

Add `System.Diagnostics.Stopwatch`.

The goal is to measure each major stage separately.

Do not guess.

Log elapsed milliseconds for at least the following:

```text
[1] SaveMappingXmlRecord
[2] ReadXMLMappingData
[3] Main SQL _Q.Select
[4] Build myDic from DataTable
[5] BeforeEnrollment loading
[6] UpdateRecord.SelectByStudentIDs
[7] Build mylist / student status filtering
[8] new Filter(...)
[9] Export - load workbook template
[10] WriteErrorSheet
[11] WriteDepartmentSheet - 普通科
[12] WriteDepartmentSheet - 專業群科(職業科)
[13] WriteDepartmentSheet - 綜合高中
[14] WriteDepartmentSheet - 實用技能學程
[15] WriteDepartmentSheet - 進修部(學校)
[16] Workbook Save Xlsx
[17] TOTAL
```

Also log useful counts:

```text
DataTable row count
unique student count
UpdateRecord count
BeforeEnrollment count
error_list count
group count per worksheet
student count per worksheet
```

Recommended output example:

```text
============================================================
新生入學統計列印開始
SchoolYear=115
SQL rows=6420
Unique students=812

[1] SaveMappingXmlRecord: 12 ms
[2] ReadXMLMappingData: 8 ms
[3] Main SQL: 950 ms
[4] Build myDic: 18300 ms
    BeforeEnrollment calls=6420
[5] UpdateRecord load: 120 ms
[6] Student status filter: 3100 ms
[7] Filter build: 80 ms
[8] Export template load: 180 ms
[9] 普通科: 400 ms
[10] 專業群科(職業科): 2600 ms
...
[17] TOTAL: 26000 ms
============================================================
```

Use an appropriate existing logging method for the project.

If no dedicated logging framework exists, use a simple trace/debug output or a collected StringBuilder shown/written for testing.

Do not interrupt the user with multiple MessageBoxes for each stage.

---

# Phase 2 - Optimize `BeforeEnrollment` Loading

## Preferred Approach

Do not call:

```csharp
BeforeEnrollment.SelectByStudentID(id)
```

for every `DataRow`.

First build the unique student IDs.

Then load BeforeEnrollment data in batch if a supported batch API already exists.

If a batch API is available, prefer:

```text
unique student ids
    ↓
SelectByStudentIDs(...)
    ↓
Dictionary<StudentID, BeforeEnrollmentRecord>
```

Do not invent an unsupported API.

---

## Safe Fallback

If only `SelectByStudentID(id)` is available, add an in-memory cache:

```csharp
Dictionary<string, K12.Data.BeforeEnrollmentRecord>
    beforeEnrollmentCache
```

Only query once per unique student ID.

Pseudo logic:

```csharp
if (!beforeEnrollmentCache.ContainsKey(id))
{
    beforeEnrollmentCache[id] =
        K12.Data.BeforeEnrollment.SelectByStudentID(id);
}

K12.Data.BeforeEnrollmentRecord ber =
    beforeEnrollmentCache[id];
```

Expected effect:

```text
SQL rows: 6000+
unique students: 800
```

BeforeEnrollment calls should fall from thousands to at most ~800.

---

# Phase 3 - Avoid Repeated Work for Duplicate SQL Rows

Inside:

```csharp
foreach (DataRow row in dt.Rows)
```

expensive student-level work should only be performed when:

```csharp
!myDic.ContainsKey(id)
```

Examples:

- BeforeEnrollment lookup
- permanent_address XML parsing
- student-level object creation
- any data that does not change between duplicate tag rows

Only the Tag accumulation should run for every SQL row:

```csharp
myDic[id].Tag.Add(ref_tag_id);
```

Do not repeatedly parse the same address XML for the same student.

---

# Phase 4 - Optimize Student Status Filtering

Current:

```csharp
CheckStudentStatus(records, studentId)
```

scans the full `records` collection repeatedly.

Replace with a precomputed set.

Build once:

```csharp
HashSet<string> validNewStudentIds
```

Include student IDs whose record satisfies:

```text
record.SchoolYear == _SchoolYear
AND Convert.ToInt16(record.UpdateCode) < 100
```

Then:

```csharp
foreach (KeyValuePair<string, myStudent> kvp in myDic)
{
    if (validNewStudentIds.Contains(kvp.Key))
        mylist.Add(kvp.Value);
}
```

Do not change the business rule.

---

# Phase 5 - Optimize `WriteDepartmentSheet()` Student Lookup

Current:

```text
recl
  x
summary
```

nested scanning must be replaced.

Build once:

```csharp
Dictionary<string, myStudent> summaryById =
    summary.ToDictionary(x => x.Id);
```

Then:

```csharp
foreach (SHBeforeEnrollmentRecord rec in recl)
{
    myStudent student;

    if (!summaryById.TryGetValue(rec.RefStudentID, out student))
        continue;

    // existing logic
}
```

Do not change the graduation/previous-school business rule.

---

# Phase 6 - Cache `CheckStudentBeforeStatus()`

Avoid calling:

```csharp
CheckStudentBeforeStatus(...)
```

multiple times for the same student.

Preferred:

Build once:

```csharp
Dictionary<string, string> studentBeforeStatusMap
```

Possible values remain exactly:

```text
當年畢業
當年修業
其他(含領結業證書)
```

Use the existing `CheckStudentBeforeStatus()` business rule to populate the lookup.

Do not alter the existing UpdateCode meanings.

If a full lookup refactor is considered too risky, at minimum:

```csharp
string status =
    CheckStudentBeforeStatus(UpdateRecord_records, student.Id);
```

must be evaluated once per student, not 3 times.

---

# Phase 7 - Move Invariant Aboriginal Calculation Outside Loop

Move:

```csharp
List<myStudent> aboStudentlist =
    filter.getListByTagId(aboIDList, summary);
```

outside:

```csharp
foreach (SHBeforeEnrollmentRecord rec in recl)
```

Do not recalculate an unchanged list for every record.

Prefer dictionary lookup for aboriginal students as well if practical.

---

# Phase 8 - Measure Tag Filtering Cost

After Phases 2-7, run timing again.

If `WriteDepartmentSheet()` is still slow, instrument:

```csharp
filter.getListByTagId(...)
filter.getGenderCount(...)
filter.getClassCount(...)
```

Do not aggressively rewrite `Filter` without measuring.

If optimization is needed, prefer:

- `HashSet<string>` for Tag ID membership checks
- precomputed student-tag lookup
- one-pass counters
- avoid re-filtering the same list repeatedly

Preserve exact report results.

---

# Phase 9 - Workbook / Aspose Optimization

Do not optimize Aspose first.

Only investigate workbook writing if timing proves it is significant.

Potential safe considerations:

- load the template only once
- do not copy worksheets unnecessarily
- avoid repeated formatting operations
- write only required cells
- save only once after all worksheets are complete

Do not alter layout, merged cells, formulas, styles, or worksheet names.

---

# Phase 10 - Workbook Save Timing

`_wk.Save(...)` is executed in:

```csharp
_BGWClassStudentAbsenceDetail_Completed
```

Add separate timing around:

```csharp
_wk.Save(sd.FileName, SaveFormat.Xlsx);
```

This allows us to distinguish:

```text
report calculation time
vs
xlsx serialization/save time
```

---

# Important Business Rules to Preserve

Do not change:

- SchoolYear validation
- `SaveMappingXmlRecord()`
- `ReadXMLMappingData()` behavior
- Student Tag mapping
- department / department group classification
- subject code sorting
- worksheet routing
- freshman UpdateCode `< 100` rule
- ClassType source
- admission method definitions
- admission identity definitions
- actual class count rule currently based on `ref_class_id`
- error worksheet logic
- workbook layout
- reset behavior

This task is performance optimization only.

---

# Recommended Optimization Order

Implement and test in this order:

```text
1. Add Stopwatch diagnostics
2. Cache / batch BeforeEnrollment
3. Move student-level work inside `!myDic.ContainsKey(id)`
4. Replace CheckStudentStatus full scans with HashSet
5. Replace recl × summary with Dictionary lookup
6. Cache CheckStudentBeforeStatus result
7. Move aboStudentlist calculation outside loop
8. Re-run timing
9. Optimize getListByTagId / getGenderCount only if still necessary
10. Measure Xlsx Save separately
```

Do not combine too many logic changes before obtaining a baseline timing.

---

# Validation

## Test 1 - Report Content Must Match

Generate the same school year before and after optimization.

Verify:

- same worksheet names
- same subject rows
- same department group routing
- same class count
- same student totals
- same gender totals
- same admission method counts
- same admission identity counts
- same error-list records

Performance optimization must not change report results.

---

## Test 2 - BeforeEnrollment Call Count

Log:

```text
SQL row count
unique student count
BeforeEnrollment query count
```

Expected after optimization:

```text
BeforeEnrollment query count <= unique student count
```

It must no longer equal the SQL DataTable row count when duplicate tag rows exist.

---

## Test 3 - Duplicate Student Rows

A student with many Student Tags must still:

- appear once in `myDic`
- keep all Tag IDs
- have BeforeEnrollment/address/student-level data loaded once

---

## Test 4 - Student Status

Verify all students previously included/excluded by:

```text
SchoolYear
UpdateCode < 100
```

remain identical after optimization.

---

## Test 5 - Previous Enrollment Status

Verify:

```text
當年畢業
當年修業
其他(含領結業證書)
```

statistics are identical before/after optimization.

---

## Test 6 - 5 Worksheets

Verify timing and output for:

```text
普通科
專業群科(職業科)
綜合高中
實用技能學程
進修部(學校)
```

Each worksheet must still receive the correct data.

---

## Test 7 - Performance Result

Record before/after timing in:

`新生入學統計調整0917.md`

Include at least:

```text
Main SQL
Build myDic
BeforeEnrollment
Student status filter
Filter construction
Each WriteDepartmentSheet
Workbook Save
TOTAL
```

Document:

```text
Before
After
Improvement
```

---

# Completion Record

After implementation, create/update:

`新生入學統計調整0917.md`

Record:

1. Modified files.
2. Original slow-flow analysis.
3. Baseline timing.
4. BeforeEnrollment N+1 issue.
5. BeforeEnrollment cache/batch solution.
6. Duplicate DataRow optimization.
7. Student status HashSet/Dictionary optimization.
8. `recl × summary` nested-loop optimization.
9. `CheckStudentBeforeStatus` cache optimization.
10. Aboriginal list loop optimization.
11. Any Filter optimization performed.
12. XLSX Save timing.
13. Before/after total elapsed time.
14. Verification that report output is unchanged.
15. Any remaining performance bottleneck.

---

# Acceptance Criteria

- [ ] Stopwatch timing is added for major report stages.
- [ ] Baseline timing is recorded.
- [ ] `BeforeEnrollment.SelectByStudentID()` is not called once per SQL row.
- [ ] Expensive student-level work is executed once per unique student.
- [ ] `CheckStudentStatus()` no longer repeatedly scans all update records for every student.
- [ ] `recl × summary` nested lookup is replaced with dictionary lookup.
- [ ] `CheckStudentBeforeStatus()` is not repeatedly calculated multiple times for the same student.
- [ ] Aboriginal student list is not recomputed inside the BeforeEnrollment loop.
- [ ] Tag/filter optimization is only added if measurements show it is still needed.
- [ ] Workbook layout and report rules remain unchanged.
- [ ] XLSX save time is measured separately.
- [ ] Before/after timing results are documented.
- [ ] Report content is verified identical before and after optimization.
- [ ] Changes are documented in `新生入學統計調整0917.md`.
