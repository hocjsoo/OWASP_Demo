using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace OwaspDemo.Models;

[Table("Accounts")]
public class Account
{
    [Key]
    public int Id { get; set; }

    [Required]
    [StringLength(50)]
    public string Username { get; set; } = string.Empty;

    [Required]
    [StringLength(100)]
    public string Password { get; set; } = string.Empty;

    [Required]
    [StringLength(100)]
    public string FullName { get; set; } = string.Empty;

    [StringLength(20)]
    public string Role { get; set; } = "Customer";

    [Column(TypeName = "decimal(18,2)")]
    public decimal Balance { get; set; } = 0;

    public string SecretNote { get; set; } = string.Empty;
}
