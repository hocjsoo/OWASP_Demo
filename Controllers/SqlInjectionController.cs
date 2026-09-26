using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OwaspDemo.Data;
using OwaspDemo.Models;

namespace OwaspDemo.Controllers;

public class SqlInjectionController : Controller
{
    private readonly AppDbContext _context;

    public SqlInjectionController(AppDbContext context)
    {
        _context = context;
    }

    [HttpGet]
    public IActionResult Index()
    {
        ViewBag.TotalAccounts = _context.Accounts.Count();
        return View();
    }

    [HttpGet]
    public IActionResult DevToolsNetwork()
    {
        return View();
    }

    // 1. VULNERABLE LOGIN (Ghép chuỗi SQL - Dính lỗi nghiêm trọng)
    [HttpPost]
    public IActionResult LoginVulnerable(string username, string password)
    {
        username ??= "";
        password ??= "";

        // LỖI: Ghép trực tiếp input từ người dùng vào câu truy vấn
        string rawSql = $"SELECT * FROM Accounts WHERE Username = '{username}' AND Password = '{password}'";

        try
        {
            var accounts = _context.Accounts.FromSqlRaw(rawSql).ToList();

            bool isBypassed = accounts.Count > 0 && 
                              (username.Contains("'") || username.Contains("--") || username.Contains("OR") || username.Contains("or"));

            string analysis = "";
            if (accounts.Count > 0)
            {
                if (isBypassed)
                {
                    analysis = "🚨 TẤN CÔNG THÀNH CÔNG! Kẻ tấn công đã đưa payload phá vỡ cấu trúc SQL. Mệnh đề điều kiện luôn đúng (OR '1'='1') và dấu chú thích '--' đã vô hiệu hóa hoàn toàn bước kiểm tra mật khẩu phía sau.";
                }
                else
                {
                    analysis = "Đăng nhập thành công với tài khoản hợp lệ trong cơ sở dữ liệu.";
                }
            }
            else
            {
                analysis = "Không tìm thấy tài khoản phù hợp.";
            }

            return Json(new
            {
                success = accounts.Count > 0,
                isExploited = isBypassed,
                executedSql = rawSql,
                analysis = analysis,
                count = accounts.Count,
                accounts = accounts.Select(a => new
                {
                    a.Id,
                    a.Username,
                    a.Password,
                    a.FullName,
                    a.Role,
                    Balance = a.Balance.ToString("N0") + " VNĐ",
                    a.SecretNote
                })
            });
        }
        catch (Exception ex)
        {
            return Json(new
            {
                success = false,
                isExploited = false,
                executedSql = rawSql,
                analysis = $"⚠️ Lỗi cú pháp SQL từ phía CSDL do ký tự đặc biệt phá vỡ câu lệnh: {ex.Message}",
                accounts = Array.Empty<object>()
            });
        }
    }

    // 2. SECURE LOGIN (Phòng thủ chuẩn với EF Core Parameterized LINQ)
    [HttpPost]
    public IActionResult LoginSecure(string username, string password)
    {
        username ??= "";
        password ??= "";

        // PHÒNG THỦ: EF Core LINQ tự động chuyển input thành tham số @p0, @p1
        var matched = _context.Accounts
            .AsNoTracking()
            .Where(a => a.Username == username && a.Password == password)
            .ToList();

        string simulatedSql = "SELECT * FROM Accounts WHERE Username = @p0 AND Password = @p1";
        string parametersLog = $"@p0 = N'{username}' (được coi là văn bản thuần)\n@p1 = N'{password}' (được coi là văn bản thuần)";

        string analysis = "";
        if (matched.Count > 0)
        {
            analysis = "Đăng nhập thành công hợp lệ. Dữ liệu input được truyền qua tham số an toàn.";
        }
        else
        {
            analysis = "🛡️ PHÒNG THỦ AN TOÀN! Hệ thống đã coi toàn bộ chuỗi payload (kể cả dấu ngoặc kép hay từ khóa OR/--) là một chuỗi văn bản thuần (String Literal). Kẻ tấn công không thể can thiệp vào logic truy vấn.";
        }

        return Json(new
        {
            success = matched.Count > 0,
            isExploited = false,
            executedSql = simulatedSql,
            parameters = parametersLog,
            analysis = analysis,
            count = matched.Count,
            accounts = matched.Select(a => new
            {
                a.Id,
                a.Username,
                Password = "••••••••", // Che mật khẩu ở nhánh an toàn
                a.FullName,
                a.Role,
                Balance = a.Balance.ToString("N0") + " VNĐ"
            })
        });
    }

    // 3. RESET CSDL về trạng thái ban đầu
    [HttpPost]
    public IActionResult ResetDatabase()
    {
        _context.Database.EnsureDeleted();
        _context.Database.EnsureCreated();
        return Json(new { message = "Đã khôi phục dữ liệu ban đầu thành công!" });
    }
}
