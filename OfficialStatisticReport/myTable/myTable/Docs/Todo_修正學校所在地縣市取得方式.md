# 目標

修正 `Form2.cs` 新生入學方式統計表「按戶籍地分」的學校所在地縣市取得方式。

目前：

```csharp
string LocalCounty = "";
```

因 `LocalCounty` 固定為空字串，造成「戶籍位於本縣市 / 戶籍非位於本縣市」判斷錯誤。

修改為從資料庫 `list` 資料表取得「學校資訊」XML，解析 `<County>` 作為 `LocalCounty`。

---

# 修改檔案

- `Form2.cs`

---

# 需求 1：新增取得學校所在地縣市的方法

在 `Form2.cs` 新增方法：

```csharp
private string GetSchoolCounty()
```

使用以下 SQL：

```sql
SELECT content
FROM list
WHERE name = '學校資訊'
LIMIT 1
```

`content` 欄位內容為 XML，例如：

```xml
<SchoolInformation>
    <ChineseName>國立臺灣海洋大學附屬基隆海事高級中等學校</ChineseName>
    <Address>202006 基隆市中正區祥豐街246號</Address>
    <County>基隆市</County>
    <Code>170403</Code>
</SchoolInformation>
```

使用 `QueryHelper` 執行 SQL，取得 `content`。

使用 `XmlDocument` 解析 XML：

```csharp
XmlDocument doc = new XmlDocument();
doc.LoadXml(content);

XmlNode countyNode =
    doc.SelectSingleNode("/SchoolInformation/County");
```

取得：

```text
基隆市
```

回傳時使用：

```csharp
countyNode.InnerText.Trim()
```

---

# 需求 2：錯誤與空值處理

`GetSchoolCounty()` 預設回傳空字串：

```csharp
string county = "";
```

需處理：

1. SQL 查不到資料。
2. `content` 為 null 或空白。
3. XML 沒有 `<County>`。
4. XML 格式錯誤。
5. SQL 或 XML 解析發生 Exception。

發生 Exception 時不要中斷整份報表產生流程。

使用：

```csharp
Debug.WriteLine("取得學校所在地縣市失敗：" + ex.Message);
```

最後回傳空字串。

---

# 需求 3：修改按戶籍地分的 LocalCounty

找到目前：

```csharp
#region 按戶籍地分

string LocalCounty = "";
```

修改為：

```csharp
#region 按戶籍地分

string LocalCounty = GetSchoolCounty();
```

後面的既有判斷邏輯保留：

```csharp
foreach (myStudent student in summary)
{
    if (student.County == LocalCounty)
    {
        collect__LocalCounty.Add(student);
    }
    else
    {
        collect__OtherCounty.Add(student);
    }
}
```

不要修改後續 Excel 欄位位置與入學方式統計邏輯。

---

# 需求 4：確認學生戶籍縣市來源維持原邏輯

學生的戶籍縣市目前來自：

```text
student.permanent_address
    ↓
AddressList
    ↓
Address
    ↓
County
```

現有解析邏輯不要修改。

例如：

```xml
<AddressList>
    <Address>
        <County>基隆市</County>
    </Address>
</AddressList>
```

則：

```csharp
student.County == "基隆市"
```

---

# 預期結果

假設學校資訊：

```xml
<County>基隆市</County>
```

則：

```csharp
LocalCounty == "基隆市"
```

學生資料：

```text
學生A County = 基隆市
→ 戶籍位於本縣市

學生B County = 新北市
→ 戶籍非位於本縣市

學生C County = 臺北市
→ 戶籍非位於本縣市
```

Excel 原本：

```csharp
cs[39, 6] // 戶籍位於本縣市
cs[40, 6] // 戶籍非位於本縣市
```

以及後續男女統計、各入學方式男女統計位置全部維持原本邏輯。

---

# 注意事項

- 不要修改學生 `County` 原本的取得方式。
- 不要修改 `summary` 資料來源。
- 不要修改本縣市/其他縣市後續 Excel `PutValue` 位置。
- 不要修改入學方式 Tag Mapping。
- 不要修改其他報表統計邏輯。
- 本次修改範圍只處理「學校所在地縣市 LocalCounty 的取得方式」。
- 優先採取最小修改原則，避免影響目前已完成的新生入學統計功能。

---

# 測試

至少測試：

1. `list` 有「學校資訊」，且 `<County>基隆市</County>`。
2. 確認 `GetSchoolCounty()` 回傳 `基隆市`。
3. 戶籍為 `基隆市` 的學生計入「戶籍位於本縣市」。
4. 戶籍為其他縣市的學生計入「戶籍非位於本縣市」。
5. `content` 為空時程式不應發生未處理例外。
6. XML 找不到 `<County>` 時程式不應發生未處理例外。
7. 確認其他新生入學統計資料及 Excel 輸出結果不受影響。

---

# 修改紀錄

修改完成後，將本次調整內容記錄至：

`新生入學統計調整0929.md`

紀錄至少包含：

- 問題原因：`LocalCounty` 原本固定為空字串。
- 新增 `GetSchoolCounty()`。
- SQL 來源：`list` 資料表 `name = '學校資訊'`。
- XML 來源：`content`。
- XML 節點：`/SchoolInformation/County`。
- `LocalCounty` 改為呼叫 `GetSchoolCounty()`。
- 本縣市/其他縣市判斷方式。
- 測試結果。