### schema
#### 一覧
```sql
SELECT schema_name FROM information_schema.schemata;
```


## DbContext
```csharp
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

public class CBbb
{
    [Key]
    public int id { get; init; }

    [Required]
    [StringLength(100)]
    public string name { get; init; }
}

public class CAaa
{
    [Key]
    public int id { get; init; }

    [ForeignKey("CaaaId")]
    public List<CBbb> listCBbb { get; set; }
}

public class MyDbContext : DbContext
{
    public DbSet<CAaa> CAaas { get; set; }
    public DbSet<CBbb> CBbbs { get; set; }

    public MyDbContext(DbContextOptions<MyDbContext> options)
        : base(options)
    {
    }
}
```

## CBbb を 追加する
```csharp
using Microsoft.EntityFrameworkCore;

public class ExampleService
{
    private readonly MyDbContext _context;

    public ExampleService(MyDbContext context){
        _context = context;
    }

    public async Task AddNewCBbbToExistingCAaa(int caaaId, string newCBbbName){
        // 既存のCAaaオブジェクトをデータベースから取得
        var caaa = await _context.CAaas
            // Include メソッドを使用して、関連する CBbb オブジェクトも同時に取得
            .Include(a => a.listCBbb)
            .FirstOrDefaultAsync(a => a.id == caaaId);

        if (caaa == null)
            throw new Exception("指定されたCAaaが見つかりません。");

        // 新しいCBbbオブジェクトを作成
        var newCBbb = new CBbb{
            name = newCBbbName
        };

        // 新しいCBbbをCAaaのリストに追加
        // この時点では、まだデータベースには反映されていない
        caaa.listCBbb.Add(newCBbb);

        
        // Entity Framework Coreが自動的に
        // 新しい CBbb オブジェクトをデータベースに挿入し
        // CAaa と CBbb の関連付けも行う

        // 変更をデータベースに保存
        await _context.SaveChangesAsync();
        // CBbb の追加と CAaa との関連付けも行われる
    }
}
```
## CAaa, CBbb 両方追加する
```csharp
public class ExampleService
{
    private readonly MyDbContext _context;

    public ExampleService(MyDbContext context){
        _context = context;
    }

    public async Task<int> CreateNewCAaaWithCBbb(string cBbbName)
    {
        // 新しいCAaaオブジェクトを作成
        var newCAaa = new CAaa { listCBbb = new List<CBbb>() };

        // 新しいCBbbオブジェクトを作成
        var newCBbb = new CBbb { name = cBbbName };

        // 新しいCBbbをCAaaのリストに追加
        newCAaa.listCBbb.Add(newCBbb);

        // 新しいCAaaをコンテキストに追加
        // これにより、Entity Framework Core は新しい CAaa とそれに関連する CBbb の両方を追跡し始めます。
        _context.CAaas.Add(newCAaa);

        // 変更をデータベースに保存
        await _context.SaveChangesAsync();

        // 新しく作成されたCAaaのIDを返す
        // データベースに保存された後、Entity Framework Core は自動的に生成された ID を newCAaa.id に設定する。
        // この ID を返すことで、呼び出し元が新しく作成された CAaa を参照できるようになる
        return newCAaa.id;
    }
}
```

## 双方向の場合
```csharp
using Microsoft.EntityFrameworkCore;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

public class CAaa
{
    [Key]
    public int id { get; init; }

    public List<CBbb> listCBbb { get; set; }
}

public class CBbb
{
    [Key]
    public int id { get; init; }

    [ForeignKey("CAaa")]
    public int caaaId { get; init; }

    [Required]
    [StringLength(100)]
    public string name { get; init; }

    public CAaa CAaa { get; set; }
}

public class MyDbContext : DbContext
{
    public DbSet<CAaa> CAaas { get; set; }
    public DbSet<CBbb> CBbbs { get; set; }

    public MyDbContext(DbContextOptions<MyDbContext> options) : base(options) { }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<CAaa>()
            .HasMany(a => a.listCBbb)
            .WithOne(b => b.CAaa)
            .HasForeignKey(b => b.caaaId);
    }
}
```
```csharp
public class ExampleService
{
    private readonly MyDbContext _context;

    public ExampleService(MyDbContext context){
        _context = context;
    }

    public async Task<int> AddNewCBbbToExistingCAaa(int caaaId, string newCBbbName)
    {
        var caaa = await _context.CAaas.FindAsync(caaaId);
        if (caaa == null)
            throw new Exception("指定されたCAaaが見つかりません。");

        var newCBbb = new CBbb
        {
            caaaId = caaaId,
            name = newCBbbName
        };

        _context.CBbbs.Add(newCBbb);
        await _context.SaveChangesAsync();

        return newCBbb.id;
    }
}
```
```csharp
public class ExampleService
{
    private readonly MyDbContext _context;

    public ExampleService(MyDbContext context){
        _context = context;
    }

    public async Task<(int caaaId, int cbbbId)> CreateNewCAaaWithCBbb(string cBbbName)
    {
        var newCAaa = new CAaa { listCBbb = new List<CBbb>() };

        var newCBbb = new CBbb{ name = cBbbName };

        newCAaa.listCBbb.Add(newCBbb);

        _context.CAaas.Add(newCAaa);
        await _context.SaveChangesAsync();

        return (newCAaa.id, newCBbb.id);
    }
}
```

## InversePropertyの例
InverseProperty 属性は、双方向のナビゲーションプロパティを明示的に指定する場合に特に有用。
これを使用することで、Entity Framework Core に対してリレーションシップをより明確に定義できる。
```csharp
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

public class CAaa
{
    [Key]
    public int id { get; init; }

    // CBbb クラスの CAaa プロパティがこのリレーションシップの逆側であることを示す
    [InverseProperty("CAaa")]
    public List<CBbb> listCBbb { get; set; }
}

public class CBbb
{
    [Key]
    public int id { get; init; }

    [ForeignKey("CAaa")]
    public int caaaId { get; init; }

    [Required]
    [StringLength(100)]
    public string name { get; init; }

    // CAaa クラスの listCBbb プロパティがこのリレーションシップの逆側であることを示す
    [InverseProperty("listCBbb")]
    public CAaa CAaa { get; set; }
}

public class MyDbContext : DbContext
{
    public DbSet<CAaa> CAaas { get; set; }
    public DbSet<CBbb> CBbbs { get; set; }

    public MyDbContext(DbContextOptions<MyDbContext> options) : base(options) { }
}
```
## InverseProperty nameof
```csharp
public class CAaa
{
    [Key]
    public int id { get; init; }

    [InverseProperty(nameof(CBbb.CAaa))]
    public List<CBbb> listCBbb { get; set; }
}

public class CBbb
{
    [Key]
    public int id { get; init; }

    [ForeignKey(nameof(CAaa))]
    public int caaaId { get; init; }

    [Required]
    [StringLength(100)]
    public string name { get; init; }

    public CAaa CAaa { get; set; }
}
```

## CBbb に対する変更
```csharp
public class CBbbService
{
    private readonly MyDbContext _context;

    public CBbbService(MyDbContext context){
        _context = context;
    }

    public async Task<bool> UpdateCBbb(int id, string newName)
    {
        // 指定されたIDのCBbbエンティティを取得
        var cBbb = await _context.CBbbs.FindAsync(id);

        // エンティティが見つからない場合はfalseを返す
        if (cBbb == null)
            return false;

        // エンティティのプロパティを更新
        cBbb.name = newName;

        // 変更を保存
        await _context.SaveChangesAsync();

        return true;
    }
}

```
```csharp
var dbContext = new MyDbContext(/* options */);
var cBbbService = new CBbbService(dbContext);

bool updateResult = await cBbbService.UpdateCBbb(1, "新しい名前");
if (updateResult)
    Console.WriteLine("更新成功");
else
    Console.WriteLine("更新失敗：指定されたIDのCBbbが見つかりません");
```

## Updateの必要性
DbContext.Update を明示的に呼び出す必要があるかどうかは、状況によって異なる。
updateの利用は、Entity Framework Core の変更追跡機能に関する重要な点でもある。
基本的には、以下のようなケースでは DbContext.Update を呼び出す必要はない。

**DbContext.Update が必要でないケース:**
1. DbContext を通じてエンティティを取得した場合（例：FindAsync や FirstOrDefaultAsync など）
1. 取得したエンティティのプロパティを直接変更する場合

これは、Entity Framework Core の変更追跡機能が自動的に働くため。
DbContext を通じて取得されたエンティティは「追跡状態」にあり、そのプロパティの変更は自動的に検出される。

**DbContext.Update が必要になる主なケース:**
1. 切り離されたエンティティの更新：
DbContext を通じて取得されていないエンティティ
(例：クライアントから送られてきたデータで再構築されたエンティティ）を更新する場合。
1. 一括更新：
エンティティの全プロパティを一度に更新する場合。

**バージョン1: Update を使用しない(通常のケース)**
```csharp
public async Task<bool> UpdateCBbb(int id, string newName)
{
    var cBbb = await _context.CBbbs.FindAsync(id);
    if (cBbb == null)
        return false;

    cBbb.name = newName;
    await _context.SaveChangesAsync();
    return true;
}
```
**バージョン2: Update を使用する(切り離されたエンティティの場合)**
```csharp
public async Task<bool> UpdateCBbb(CBbb updatedCBbb)
{
    _context.Update(updatedCBbb);
    try
    {
        await _context.SaveChangesAsync();
        return true;
    }
    catch (DbUpdateConcurrencyException)
    {
        if (!await _context.CBbbs.AnyAsync(c => c.id == updatedCBbb.id))
            return false;
        throw;
    }
}
```
---
Add、Removeなどの関数も同様の原則に従う。

Entity Framework Coreの変更追跡システムは、
コンテキストを通じて取得または作成されたエンティティに対して自動的に機能する。

<u>Add（追加）:</u>
新しいエンティティを作成する場合は通常 Add を使用する。
```csharp
var newCBbb = new CBbb { name = "New CBbb" };
_context.CBbbs.Add(newCBbb);
await _context.SaveChangesAsync();
```
<u>Remove（削除）:</u>
コンテキストで追跡されているエンティティを削除する場合は Remove を使用する。
```csharp
var cBbbToRemove = await _context.CBbbs.FindAsync(id);
if (cBbbToRemove != null)
{
    _context.CBbbs.Remove(cBbbToRemove);
    await _context.SaveChangesAsync();
}
```

<u>Update（更新）:</u>
通常は不要。切り離されたエンティティの場合に使用。


<u>Attach（アタッチ）:</u>
切り離されたエンティティをコンテキストに追加する。
変更を追跡しない場合に使用。
```csharp
_context.Attach(someCBbb);
```

**重要なポイント：**

<u>自動追跡:</u>
_context.CBbbs.Find(), _context.CBbbs.FirstOrDefault() などで取得したエンティティは自動的に追跡される。
新規エンティティ:
new キーワードで作成した新しいエンティティは、Add メソッドを使用するまで追跡されない。

<u>一括操作:</u>
大量のデータを扱う場合、AddRange, RemoveRange などのメソッドを使用できる。
```csharp
_context.CBbbs.AddRange(newCBbbs);
await _context.SaveChangesAsync();
```
<u>状態の明示的な設定:</u>
必要に応じて、エンティティの状態を明示的に設定することもできる。
```csharp
_context.Entry(someCBbb).State = EntityState.Modified;
```
<u>非追跡クエリ:</u>
パフォーマンス向上のため、変更を追跡する必要がない場合は AsNoTracking() を使う
```csharp
var cBbbs = await _context.CBbbs.AsNoTracking().ToListAsync();
```
基本的に、コンテキストを通じて操作を行う限り、EF Coreは適切に変更を追跡する。


Add, Remove などのメソッドは、
主に新しいエンティティの追加や、コンテキストに存在しないエンティティの操作に使用する。

