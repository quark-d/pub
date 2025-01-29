### **✅ `Fluent API` (`HasComputedColumnSql()`) を使う場合、`weekly_number` の型は `int` で良いか？**
**結論:**  
**はい、`weekly_number` のデータ型は `INT` でOK です。**  
ただし、テーブル作成時には **`weekly_number` の型を明示的に指定する必要はありません**。  
なぜなら、**`HasComputedColumnSql("DATEPART(WEEK, date)")` を適用すると、EF Core が自動的に `INT` 型の `computed column` として処理する** ためです。

---

## **🔹 `Fluent API` (`HasComputedColumnSql()`) を適用する方法**
### **📌 `DayWork` クラス**
```csharp
using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

[Table("day_work2")]  // 既存のテーブル名
public class DayWork
{
    [Key]
    [Column("day_work_id")]
    public int ID { get; set; }

    [Column("date")]
    public DateOnly WorkDate { get; set; }

    [Column("weekly_number")]
    [DatabaseGenerated(DatabaseGeneratedOption.Computed)] // SQL Server 側で計算
    public int? WeeklyNumber { get; private set; }  // `computed column` は `set` を `private` にする
}
```

✅ **ポイント**
- `weekly_number` は `DatabaseGeneratedOption.Computed` にすることで、**SQL Server 側で自動計算** される。
- `set;` を `private` にすることで、C# 側から `weekly_number` を変更できないようにする。

---

### **📌 `DbContext` (`Fluent API` の適用)**
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
            .HasComputedColumnSql("DATEPART(WEEK, date)"); // SQL Server 側で `computed column` を設定
    }
}
```

✅ **ポイント**
- `HasComputedColumnSql("DATEPART(WEEK, date)")` を使うことで、**SQL Server 側で `weekly_number` を自動計算する `computed column` に設定**。
- `HasComputedColumnSql()` を適用すると、**`weekly_number` のデータ型は `INT` になる**。

---

## **🔹 `weekly_number` の型は `int` でOK ？**
**答え:** ✅ **はい、SQL Server では `DATEPART(WEEK, date)` の戻り値は `INT` なので `weekly_number` の型は `int` で問題ありません。**

**📌 `DATEPART(WEEK, date)` の戻り値を確認**
```sql
SELECT DATEPART(WEEK, '2025-01-09') AS week_number;
```
✅ **結果**
```
week_number
------------
2
```
このように、**`DATEPART(WEEK, date)` の戻り値は `INT`** なので、`weekly_number` の型は `int` でOK。

---

## **🔹 既存の `day_work2` テーブルに `computed column` を適用**
**`CREATE TABLE` 時に `weekly_number` を `INT` で作成している場合、`ALTER TABLE` で `computed column` に変更する必要がある。**

### **📌 SQL Server 側で `ALTER TABLE` を適用**
```sql
ALTER TABLE day_work2 DROP COLUMN weekly_number; -- 既存の `weekly_number` を削除
ALTER TABLE day_work2 ADD weekly_number AS DATEPART(WEEK, date); -- `computed column` を追加
```

### **📌 C# (`ExecuteSqlRaw()`) から `ALTER TABLE` を適用**
```csharp
using System;
using Microsoft.EntityFrameworkCore;

class Program
{
    static void Main()
    {
        using (var context = new DayWorkContext())
        {
            // `weekly_number` を `computed column` に変更
            context.Database.ExecuteSqlRaw("ALTER TABLE day_work2 DROP COLUMN weekly_number;");
            context.Database.ExecuteSqlRaw("ALTER TABLE day_work2 ADD weekly_number AS DATEPART(WEEK, date);");

            Console.WriteLine("weekly_number を computed column に設定しました。");
        }
    }
}
```

✅ **ポイント**
- `weekly_number` を `computed column` に変更するには、**一旦削除 (`DROP COLUMN`) してから再作成 (`ADD COLUMN`) する必要がある**。
- **`ExecuteSqlRaw()` を使えば、C# から `ALTER TABLE` を適用できる**。

---

## **✅ まとめ**
| **項目** | **設定内容** |
|---------|------------|
| `weekly_number` のデータ型 | ✅ `INT` でOK（`DATEPART(WEEK, date)` の戻り値は `INT`） |
| `Fluent API` の設定 | ✅ `HasComputedColumnSql("DATEPART(WEEK, date)")` を使用 |
| `CREATE TABLE` 時の型指定 | ❌ **不要**（`computed column` の場合、`ALTER TABLE` で定義する） |
| `既存テーブルの変更方法` | ✅ `ALTER TABLE` を適用 (`DROP COLUMN` → `ADD COLUMN`) |

✅ **おすすめの方法**
1. **新規にテーブルを作る場合 → `Fluent API` (`HasComputedColumnSql`) を使う**
2. **既存のテーブルに適用する場合 → `ALTER TABLE` で `computed column` を追加**
3. **C# から適用する場合 → `ExecuteSqlRaw("ALTER TABLE ...")` を使う**

**💡 `weekly_number` の型は `INT` で問題なし！**
**SQL Server 側で `DATEPART(WEEK, date)` を適用して、自動計算を実現しましょう！ 🚀**
