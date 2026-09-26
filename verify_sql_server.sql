-- ==============================================================================
-- BỘ TRUY VẤN THỰC NGHIỆM ĐỐI CHỨNG TRÊN MICROSOFT SQL SERVER 2025
-- Database: OwaspDemoDB | Nhóm: WNC.G01 (Học - Bình - Cường)
-- Hướng dẫn: Bôi đen từng phần và bấm Ctrl+Shift+E để thực thi trong VS Code
-- ==============================================================================

USE OwaspDemoDB;
GO

-- ==============================================================================
-- PHẦN 1: THỰC NGHIỆM SQL INJECTION (NGUYỄN DANH HỌC)
-- ==============================================================================

-- 1.1 Xem bảng tài khoản gốc
SELECT Id, Username, Password, FullName, Role, Balance, SecretNote 
FROM dbo.Accounts;
GO

-- 1.2 Câu lệnh dính lỗi nối chuỗi: Bẻ gãy logic kiểm tra mật khẩu
-- Dấu '--' vô hiệu hóa toàn bộ vế AND Password = ...
SELECT * 
FROM dbo.Accounts 
WHERE Username = '' OR '1'='1' --' AND Password = 'mat_khau_bat_ky';
GO
-- Kết quả: Trả về TOÀN BỘ 4 bản ghi vì điều kiện (1=1) luôn đúng!

-- 1.3 Câu lệnh tham số hóa chuẩn (EF Core LINQ sinh ra):
EXEC sp_executesql 
    N'SELECT * FROM dbo.Accounts WHERE Username = @p0 AND Password = @p1',
    N'@p0 nvarchar(50), @p1 nvarchar(50)',
    @p0 = N''' OR ''1''=''1'' --',
    @p1 = N'mat_khau_bat_ky';
GO
-- Kết quả: Trả về 0 bản ghi vì chuỗi payload chỉ được coi là text thuần.


-- ==============================================================================
-- PHẦN 2: THỰC NGHIỆM STORED XSS (NGUYỄN THANH BÌNH)
-- ==============================================================================

-- 2.1 Xem bảng bình luận đã lưu trong CSDL
SELECT Id, Author, Content, CreatedAt, IsSecureStored 
FROM dbo.Comments 
ORDER BY CreatedAt DESC;
GO

-- 2.2 Phân tích: Nếu hacker lưu trực tiếp chuỗi script vào cột Content:
-- Ví dụ: <script>alert(document.cookie);</script>
-- Khi ứng dụng đọc ra và render bằng @Html.Raw(), script sẽ thực thi trên trình duyệt nạn nhân.


-- ==============================================================================
-- PHẦN 3: THỰC NGHIỆM CSRF ATTACK (NGUYỄN MINH CƯỜNG)
-- ==============================================================================

-- 3.1 Xem số dư ví của Nạn nhân và Hacker trong CSDL
SELECT Id, OwnerName, Balance 
FROM dbo.UserWallets;
GO

-- 3.2 Khôi phục số dư ví ban đầu nếu cần reset:
-- UPDATE dbo.UserWallets SET Balance = 50000000.00 WHERE Id = 1;
-- UPDATE dbo.UserWallets SET Balance = 0.00 WHERE Id = 2;
-- GO
