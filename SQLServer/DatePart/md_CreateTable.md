### **✅ `computed column` (計算列) を使う方法**
---
## **🔹 `CREATE TABLE` の SQL (SQL Server 側)**
```sql
CREATE TABLE day_work (
    day_work_id INT IDENTITY(1,1) PRIMARY KEY, 
    date DATE NOT NULL,
    weekly_number AS DATEPART(WEEK, date) -- `computed column`
);
```
✅ **ポイント**
- `weekly_number` は **計算列 (`computed column`) のため、データベースには保存されない**。
- `SELECT` 時に **`date` の値に基づいて動的に計算される**。
- `INSERT` のとき **`weekly_number` を指定する必要はない**。

---

## **🔹 `DayWork` クラス (C# 側)**
```csharp
using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

[Table("day_work")]
public class DayWork
{
    [Key]
    [Column("day_work_id")]
    public int ID { get; set; }

    [Column("date")]
    public DateOnly WorkDate { get; set; } // `date` 列（作業日）

    [Column("weekly_number")]
    [DatabaseGenerated(DatabaseGeneratedOption.Computed)] // SQL Server 側で計算される
    public int WeeklyNumber { get; private set; } // `computed column` のため `set` は `private`
}
```

✅ **ポイント**
- `DatabaseGeneratedOption.Computed` を指定し、**SQL Server 側で `weekly_number` を計算することを示す**。
- `WeeklyNumber` の **`set` を `private` にすることで、C# 側からの変更を防ぐ**。

---

## **🔹 `DbContext` (`Fluent API` の適用)**
```csharp
using Microsoft.EntityFrameworkCore;

public class DayWorkContext : DbContext
{
    public DbSet<DayWork> DayWorks { get; set; }

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        optionsBuilder.UseSqlServer("Server=DESKTOP-7861UKF\\MSSQLLOCAL;Database=DataPartSample;Integrated Security=True;");
    }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<DayWork>()
            .Property(d => d.WeeklyNumber)
            .HasComputedColumnSql("DATEPART(WEEK, date)"); // `computed column` を設定
    }
}
```

✅ **ポイント**
- **`HasComputedColumnSql("DATEPART(WEEK, date)")` を `Fluent API` で設定** することで、EF Core に「このカラムは SQL Server 側で自動計算される」と伝える。

---

## **🔹 C# でデータの追加 (`INSERT`)**
```csharp
using System;

class Program
{
    static void Main()
    {
        using (var context = new DayWorkContext())
        {
            // 2025/1/9 のデータを追加
            var newWork = new DayWork { WorkDate = new DateOnly(2025, 1, 9) };

            // データベースに追加 (`weekly_number` は自動計算される)
            context.DayWorks.Add(newWork);
            context.SaveChanges();

            Console.WriteLine($"データを追加しました: ID = {newWork.ID}, Date = {newWork.WorkDate}");
        }
    }
}
```

✅ **ポイント**
- `INSERT` の際に **`weekly_number` は指定せずに `date` だけを指定する**。
- **SQL Server 側で `weekly_number` が `DATEPART(WEEK, date)` に基づいて計算される**。

---

## **🔹 C# でデータの取得 (`SELECT`)**
```csharp
using System;
using System.Linq;

class Program
{
    static void Main()
    {
        using (var context = new DayWorkContext())
        {
            var records = context.DayWorks.ToList();
            foreach (var work in records)
            {
                Console.WriteLine($"ID: {work.ID}, Date: {work.WorkDate}, WeeklyNumber: {work.WeeklyNumber}");
            }
        }
    }
}
```

✅ **出力結果**
```
ID: 1, Date: 2025-01-09, WeeklyNumber: 2
```
✅ **動作の流れ**
1. **SQL Server 側で `weekly_number = DATEPART(WEEK, date)` を適用**。
2. `SELECT` すると **`weekly_number` が自動計算される**。

---

## **✅ 既存のテーブルに `computed column` を追加する場合**
既に `day_work` テーブルが存在する場合、  
**`ALTER TABLE` を使って `computed column` に変更する** 必要があります。

```sql
ALTER TABLE day_work DROP COLUMN weekly_number;
ALTER TABLE day_work ADD weekly_number AS DATEPART(WEEK, date);
```

✅ **ポイント**
- `DROP COLUMN` してから `ADD COLUMN` しないとエラーになることがある。
- `weekly_number` を `computed column` に変更。

---

## **✅ まとめ**
| **項目** | **設定内容** |
|---------|------------|
| `weekly_number` のデータ型 | ✅ `INT` でOK（`DATEPART(WEEK, date)` の戻り値は `INT`） |
| `Fluent API` の設定 | ✅ `HasComputedColumnSql("DATEPART(WEEK, date)")` を使用 |
| `CREATE TABLE` 時の型指定 | ❌ **不要**（`computed column` の場合、`ALTER TABLE` で定義する） |
| `既存テーブルの変更方法` | ✅ `ALTER TABLE` を適用 (`DROP COLUMN` → `ADD COLUMN`) |

---

## **🚀 最終的な実装 (C# + SQL Server)**
### **📌 `CREATE TABLE` (SQL Server)**
```sql
CREATE TABLE day_work (
    day_work_id INT IDENTITY(1,1) PRIMARY KEY,
    date DATE NOT NULL,
    weekly_number AS DATEPART(WEEK, date) -- `computed column`
);
```

### **📌 `DayWork` クラス**
```csharp
public class DayWork
{
    [Key]
    [Column("day_work_id")]
    public int ID { get; set; }

    [Column("date")]
    public DateOnly WorkDate { get; set; }

    [Column("weekly_number")]
    [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
    public int WeeklyNumber { get; private set; }
}
```

### **📌 `DbContext`**
```csharp
protected override void OnModelCreating(ModelBuilder modelBuilder)
{
    modelBuilder.Entity<DayWork>()
        .Property(d => d.WeeklyNumber)
        .HasComputedColumnSql("DATEPART(WEEK, date)");
}
```

### **📌 `INSERT` と `SELECT`**
```csharp
var newWork = new DayWork { WorkDate = new DateOnly(2025, 1, 9) };
context.DayWorks.Add(newWork);
context.SaveChanges();

var records = context.DayWorks.ToList();
foreach (var work in records)
{
    Console.WriteLine($"ID: {work.ID}, Date: {work.WorkDate}, WeeklyNumber: {work.WeeklyNumber}");
}
```

✅ **期待される出力**
```
ID: 1, Date: 2025-01-09, WeeklyNumber: 2
```

---

## **✅ まとめ**
- `computed column` を使うことで、`weekly_number` を **SQL Server 側で自動計算** できる。
- **C# 側で `weekly_number` を指定せずに `INSERT` 可能**。
- **最もシンプルで管理しやすい方法！** 🚀

この方法で試してみてください！ 💡
