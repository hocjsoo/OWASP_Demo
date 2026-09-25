using Microsoft.EntityFrameworkCore;
using OwaspDemo.Models;

namespace OwaspDemo.Data;

public class AppDbContext(DbContextOptions<AppDbContext> options) : DbContext(options)
{
    public DbSet<Account> Accounts => Set<Account>();

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        base.OnModelCreating(modelBuilder);

        modelBuilder.Entity<Account>().HasData(
            new Account
            {
                Id = 1,
                Username = "admin",
                Password = "AdminPassword@2026",
                FullName = "Quản Trị Viên Hệ Thống",
                Role = "Admin",
                Balance = 100000000m,
                SecretNote = "Mật mã máy chủ & kho dữ liệu: SEC#2026!TOP"
            },
            new Account
            {
                Id = 2,
                Username = "chuyengia_an",
                Password = "AnExpert#Pass1",
                FullName = "ThS. Nguyễn Văn An (Chuyên Gia)",
                Role = "Specialist",
                Balance = 25000000m,
                SecretNote = "Thông tin cá nhân & chứng chỉ hành nghề VIP"
            },
            new Account
            {
                Id = 3,
                Username = "khachhang_binh",
                Password = "BinhUser#Pass2",
                FullName = "Lê Thanh Bình (Khách Hàng)",
                Role = "Customer",
                Balance = 5200000m,
                SecretNote = "Lịch sử tư vấn bệnh án tâm lý cá nhân"
            },
            new Account
            {
                Id = 4,
                Username = "khachhang_cuong",
                Password = "CuongUser#Pass3",
                FullName = "Trần Minh Cường (Khách Hàng)",
                Role = "Customer",
                Balance = 2800000m,
                SecretNote = "Số thẻ tín dụng liên kết: 4111-XXXX-XXXX-9921"
            }
        );
    }
}
