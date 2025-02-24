using System.ComponentModel.DataAnnotations.Schema;

namespace EnumForeignKey;

[Table("product")]
public class Product
{
    [Column("product_id")]
    public int ID { get; set; }

    [Column("product_name")]
    public string? Name { get; set; }

    private int typeId;

    [Column("product_type_id")]
    public int TypeId
    {
        get => typeId;
        set
        {
            if (!Enum.IsDefined(typeof(TypeEnum), value))
            {
                throw new ArgumentOutOfRangeException(nameof(value), "Invalid value for TypeEnum");
            }
            typeId = value;
        }
    }

    [NotMapped]
    public TypeEnum Type
    {
        get => (TypeEnum)TypeId;
        set => TypeId = (int)value;
    }
}
