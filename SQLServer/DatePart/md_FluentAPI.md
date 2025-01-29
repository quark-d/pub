### **✅ C# から **``** を使わずに **``** を設定する方法**

C# から **Entity Framework Core (EF Core) を使って **``** を **``** で自動計算** させる方法はいくつかあります。\
``** を直接使わずに、C# 側からデータベースを操作する方法を紹介します。**

---

## **🔹 方法①: **``** を使って **``** を設定 (推奨)**

EF Core の `Fluent API` を使って、**C# 側で **``** を自動適用する**。

### **📌 **``** クラス**

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
    public DateOnly WorkDate { get; set; } // 作業日

    [Column("weekly_number")]
    [DatabaseGenerated(DatabaseGeneratedOption.Computed)] // SQL Server 側で計算
    public int WeeklyNumber { get; private set; } // `computed column` のため `set` は `private`
}
```

---

### **📌 **``** (**``** で **``** を設定)**

```csharp
using Microsoft.EntityFrameworkCore;

public class DayWorkContext : DbContext
{
    public DbSet<DayWork> DayWorks { get; set; }

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        optionsBuilder.UseSqlServer("Server=DESKTOP-7861UKF\\MSS_LOCAL;Database=DataPartSample;Integrated Security=True;");
    }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<DayWork>()
            .Property(d => d.WeeklyNumber)
            .HasComputedColumnSql("DATEPART(WEEK, date)"); // `computed column` を設定
    }
}
```

✅ **特徴**

- `OnModelCreating` の `HasComputedColumnSql("DATEPART(WEEK, date)")` で `weekly_number` を計算列として設定。
- `CREATE TABLE` を **C# 側から実行せずに **``** を適用可能**。

✅ **使用例**

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

            // データベースに追加
            context.DayWorks.Add(newWork);
            context.SaveChanges();

            Console.WriteLine($"データを追加しました: ID = {newWork.ID}, Date = {newWork.WorkDate}");
        }
    }
}
```

✅ **出力結果**

```
データを追加しました: ID = 1, Date = 2025-01-09
```

（`weekly_number` は SQL Server 側で自動計算）

✅ ``** で取得**

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

---

## **🔹 方法②: **``** で **``** を実行**

C# から `ALTER TABLE` を実行して、`weekly_number` を `DATEPART(WEEK, date)` で自動計算するように設定できます。

```csharp
using System;
using Microsoft.EntityFrameworkCore;

class Program
{
    static void Main()
    {
        using (var context = new DayWorkContext())
        {
            // SQL を実行して `computed column` を追加
            context.Database.ExecuteSqlRaw("ALTER TABLE day_work ADD weekly_number AS DATEPART(WEEK, date)");

            Console.WriteLine("weekly_number を computed column に設定しました。");
        }
    }
}
```

✅ **特徴**

- `CREATE TABLE` を使わずに、C# から `ALTER TABLE` を実行して `computed column` を追加。
- **データベースを変更できる権限が必要**。

---

## **🔹 方法③: **``** をオーバーライド**

EF Core で `SaveChanges()` をオーバーライドし、C# 側で `weekly_number` を設定する方法もあります。

```csharp
using System;
using System.Linq;
using Microsoft.EntityFrameworkCore;

public class DayWorkContext : DbContext
{
    public DbSet<DayWork> DayWorks { get; set; }

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        optionsBuilder.UseSqlServer("Server=DESKTOP-7861UKF\\MSS_LOCAL;Database=DataPartSample;Integrated Security=True;");
    }

    public override int SaveChanges()
    {
        foreach (var entry in ChangeTracker.Entries<DayWork>())
        {
            if (entry.State == EntityState.Added || entry.State == EntityState.Modified)
            {
                entry.Entity.WeeklyNumber = GetWeekNumber(entry.Entity.WorkDate);
            }
        }
        return base.SaveChanges();
    }

    private int GetWeekNumber(DateOnly date)
    {
        return System.Globalization.CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(
            new DateTime(date.Year, date.Month, date.Day),
            System.Globalization.CalendarWeekRule.FirstFourDayWeek,
            DayOfWeek.Monday
        );
    }
}
```

✅ **特徴**

- `weekly_number` を **C# 側で計算して保存** する。
- **SQL Server 側での **``** の使用は不要**。

---

## **🔹 どの方法を選ぶべきか？**

| **方法**                       | **特徴**                      | **メリット**                                | **デメリット**              |
| ---------------------------- | --------------------------- | --------------------------------------- | ---------------------- |
| **方法①: Fluent API (**``**)** | `computed column` を C# 側で適用 | SQL Server で `DATEPART(WEEK, date)` を使用 | `UPDATE` 不可            |
| **方法②: **``                  | `computed column` を後付け      | 既存のDBに適用可能                              | `ALTER TABLE` 実行の権限が必要 |
| **方法③: **``** をオーバーライド**     | C# 側で `weekly_number` を計算   | SQL Server に依存しない                       | `weekly_number` を手動で設定 |

✅ **最適な選択**

- **DBのスキーマ変更をC#側で管理したい** → **方法① (**``**)**
- **既存のDBに**``** を後付けしたい** → **方法② (**``**)**
- **SQL Server ではなく、C# 側で **``** を設定したい** → **方法③ (**``** をオーバーライド)**

---

## **🚀 まとめ**

- ``** を使わずに **``** を適用する方法はある**。
- **最もシンプルで推奨なのは **``** (**``**) を使う方法**。
- **既存のDBに適用するなら **``** を使う方法** もあり。

この方法で `weekly_number` を `DATEPART(WEEK, date)` で設定できます！ 🚀
