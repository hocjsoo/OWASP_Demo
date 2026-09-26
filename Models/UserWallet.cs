using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace OwaspDemo.Models;

[Table("UserWallets")]
public class UserWallet
{
    [Key]
    public int Id { get; set; }

    [Required]
    [StringLength(50)]
    public string OwnerName { get; set; } = string.Empty;

    [Column(TypeName = "decimal(18,2)")]
    public decimal Balance { get; set; } = 50000000m; // 50 triệu ban đầu
}
