using Microsoft.EntityFrameworkCore;

namespace EnumForeignKey;
public class ProductDbContext : DbContext
{
    public DbSet<Product> Products { get; set; }

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        // SQL Server の接続文字列を指定
        optionsBuilder.UseSqlServer("Your_Connection_String");
    }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        base.OnModelCreating(modelBuilder);
        // modelBuilder.Entity<Product>()
        //     .HasKey(d => d.ID);
        // .hasforeignKey(d => d.TypeId)
        //     // .HasOne(d => d.Type)
        //     // .WithMany()
        //     .HasForeignKey(d => d.TypeId)
        //     .HasPrincipalKey(t => t.Id);    

    }
}
