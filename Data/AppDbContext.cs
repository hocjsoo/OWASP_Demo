using Microsoft.EntityFrameworkCore;
using OwaspDemo.Models;

namespace OwaspDemo.Data;

public class AppDbContext(DbContextOptions<AppDbContext> options) : DbContext(options)
{
    public DbSet<Account> Accounts => Set<Account>();
    public DbSet<Comment> Comments => Set<Comment>();
    public DbSet<UserWallet> Wallets => Set<UserWallet>();

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

        modelBuilder.Entity<Comment>().HasData(
            new Comment
            {
                Id = 1,
                Author = "Lê Thanh Bình",
                Content = "Dịch vụ tư vấn tâm lý rất tận tình và chuyên nghiệp!",
                CreatedAt = DateTime.Now.AddDays(-2),
                IsSecureStored = false
            },
            new Comment
            {
                Id = 2,
                Author = "Trần Minh Cường",
                Content = "Chuyên gia luật tư vấn hợp đồng kinh tế rất chi tiết, 5 sao!",
                CreatedAt = DateTime.Now.AddDays(-1),
                IsSecureStored = false
            }
        );

        modelBuilder.Entity<UserWallet>().HasData(
            new UserWallet
            {
                Id = 1,
                OwnerName = "Nạn nhân (Nguyễn Danh Học)",
                Balance = 50000000m
            },
            new UserWallet
            {
                Id = 2,
                OwnerName = "Kẻ tấn công (Hacker BlackHat)",
                Balance = 0m
            }
        );
    }
}
