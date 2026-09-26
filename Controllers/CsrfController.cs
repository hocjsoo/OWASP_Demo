using Microsoft.AspNetCore.Mvc;
using OwaspDemo.Data;

namespace OwaspDemo.Controllers;

public class CsrfController : Controller
{
    private readonly AppDbContext _context;

    public CsrfController(AppDbContext context)
    {
        _context = context;
    }

    [HttpGet]
    public IActionResult Index()
    {
        var victim = _context.Wallets.Find(1);
        var hacker = _context.Wallets.Find(2);

        ViewBag.Victim = victim;
        ViewBag.Hacker = hacker;
        return View();
    }

    // 1. NHÁNH BỊ LỖI (Vulnerable): Không có [ValidateAntiForgeryToken]
    // Bất kỳ trang web của hacker nào cũng có thể gửi POST ngầm tới đây mà không bị chặn
    [HttpPost]
    public IActionResult TransferVulnerable(decimal amount)
    {
        var victim = _context.Wallets.Find(1);
        var hacker = _context.Wallets.Find(2);

        if (victim != null && hacker != null && victim.Balance >= amount && amount > 0)
        {
            victim.Balance -= amount;
            hacker.Balance += amount;
            _context.SaveChanges();
            TempData["CsrfVulnMsg"] = $"🚨 TẤN CÔNG THÀNH CÔNG! Kẻ tấn công vừa kích hoạt giao dịch ngầm, rút sạch {amount:N0} VNĐ từ ví nạn nhân sang ví hacker!";
        }
        else
        {
            TempData["CsrfVulnMsg"] = "Số dư không đủ hoặc số tiền không hợp lệ.";
        }

        return RedirectToAction(nameof(Index));
    }

    // 2. NHÁNH AN TOÀN (Secure): Bắt buộc kiểm tra Token phòng thủ [ValidateAntiForgeryToken]
    [HttpPost]
    [ValidateAntiForgeryToken]
    public IActionResult TransferSecure(decimal amount)
    {
        var victim = _context.Wallets.Find(1);
        var hacker = _context.Wallets.Find(2);

        if (victim != null && hacker != null && victim.Balance >= amount && amount > 0)
        {
            victim.Balance -= amount;
            hacker.Balance += amount;
            _context.SaveChanges();
            TempData["CsrfSecMsg"] = $"Giao dịch hợp lệ chính chủ đã chuyển: {amount:N0} VNĐ.";
        }
        else
        {
            TempData["CsrfSecMsg"] = "Giao dịch không thành công.";
        }

        return RedirectToAction(nameof(Index));
    }

    // Trang web giả mạo do hacker dựng lên để lừa người dùng
    [HttpGet]
    public IActionResult AttackerSite()
    {
        return View();
    }

    [HttpPost]
    public IActionResult ResetWallets()
    {
        var victim = _context.Wallets.Find(1);
        var hacker = _context.Wallets.Find(2);
        if (victim != null) victim.Balance = 50000000m;
        if (hacker != null) hacker.Balance = 0m;
        _context.SaveChanges();
        return RedirectToAction(nameof(Index));
    }
}
