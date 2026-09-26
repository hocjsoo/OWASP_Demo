using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace OwaspDemo.Models;

[Table("Comments")]
public class Comment
{
    [Key]
    public int Id { get; set; }

    [Required]
    [StringLength(50)]
    public string Author { get; set; } = string.Empty;

    [Required]
    public string Content { get; set; } = string.Empty;

    public DateTime CreatedAt { get; set; } = DateTime.Now;

    public bool IsSecureStored { get; set; } = false;
}
