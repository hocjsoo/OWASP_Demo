using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OwaspDemo.Data;
using OwaspDemo.Models;
using System.Text.Encodings.Web;

namespace OwaspDemo.Controllers;

public class XssController : Controller
{
    private readonly AppDbContext _context;

    public XssController(AppDbContext context)
    {
        _context = context;
    }

    [HttpGet]
    public IActionResult Index()
    {
        // Thiết lập một cookie nhạy cảm giả lập để hacker đánh cắp qua document.cookie
        Response.Cookies.Append("AuthSessionToken", "SEC_COOKIE_VIP_TOKEN_998877_SECRET_USER_HOC", new CookieOptions
        {
            HttpOnly = false, // Cố tình để false để minh họa hậu quả XSS đọc được cookie
            Expires = DateTimeOffset.Now.AddDays(1)
        });

        var comments = _context.Comments.OrderByDescending(c => c.CreatedAt).ToList();
        return View(comments);
    }

    // 1. NHÁNH LỖI (Vulnerable): Lưu mã độc raw HTML không lọc
    [HttpPost]
    public IActionResult AddVulnerable(string author, string content)
    {
        if (string.IsNullOrWhiteSpace(author) || string.IsNullOrWhiteSpace(content))
        {
            return RedirectToAction(nameof(Index));
        }

        var comment = new Comment
        {
            Author = author.Trim(),
            Content = content, // Lưu nguyên văn script độc hại
            CreatedAt = DateTime.Now,
            IsSecureStored = false
        };

        _context.Comments.Add(comment);
        _context.SaveChanges();

        return RedirectToAction(nameof(Index));
    }

    // 2. NHÁNH AN TOÀN (Secure): Mã hóa HTML / Lọc sạch thẻ script trước khi lưu hoặc hiển thị
    [HttpPost]
    public IActionResult AddSecure(string author, string content)
    {
        if (string.IsNullOrWhiteSpace(author) || string.IsNullOrWhiteSpace(content))
        {
            return RedirectToAction(nameof(Index));
        }

        // Phòng thủ: Mã hóa các ký tự <, >, ", ' thành các thực thể HTML an toàn
        string encodedContent = HtmlEncoder.Default.Encode(content);

        var comment = new Comment
        {
            Author = HtmlEncoder.Default.Encode(author.Trim()),
            Content = encodedContent, // Lưu nội dung đã được Sanitized / Encoded
            CreatedAt = DateTime.Now,
            IsSecureStored = true
        };

        _context.Comments.Add(comment);
        _context.SaveChanges();

        return RedirectToAction(nameof(Index));
    }

    [HttpPost]
    public IActionResult ClearComments()
    {
        _context.Comments.RemoveRange(_context.Comments);
        _context.Comments.Add(new Comment
        {
            Author = "Lê Thanh Bình",
            Content = "Dịch vụ tư vấn tâm lý rất tận tình và chuyên nghiệp!",
            CreatedAt = DateTime.Now.AddDays(-2),
            IsSecureStored = true
        });
        _context.SaveChanges();
        return RedirectToAction(nameof(Index));
    }
}
