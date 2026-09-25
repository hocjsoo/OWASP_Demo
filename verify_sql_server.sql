-- ==============================================================================
-- MINH CHỨNG TRUY VẤN TRỰC TIẾP TRÊN MICROSOFT SQL SERVER 2025
-- Database: OwaspDemoDB | Bảng: Accounts
-- Mở file này trên VS Code (nhấn Ctrl+Shift+E hoặc dùng extension MSSQL)
-- ==============================================================================

USE OwaspDemoDB;
GO

-- ------------------------------------------------------------------------------
-- 1. XEM DỮ LIỆU GỐC TRONG BẢNG ACCOUNTS (4 TÀI KHOẢN MẪU)
-- ------------------------------------------------------------------------------
SELECT Id, Username, Password, FullName, Role, Balance, SecretNote
FROM Accounts;
GO

-- ------------------------------------------------------------------------------
-- 2. MINH CHỨNG TẤN CÔNG: CÂU LỆNH SQL THỰC TẾ BỊ NỐI CHUỖI (INJECTION)
-- Payload đưa vào Username: ' OR '1'='1' --
-- Dấu '--' vô hiệu hóa toàn bộ phần kiểm tra mật khẩu phía sau!
-- ------------------------------------------------------------------------------
SELECT * 
FROM Accounts 
WHERE Username = '' OR '1'='1' --' AND Password = 'mat_khau_bat_ky';
GO
-- KẾT QUẢ: SQL Server trả về toàn bộ các dòng vì điều kiện ('1'='1') luôn đúng!


-- ------------------------------------------------------------------------------
-- 3. MINH CHỨNG PHÒNG THỦ: TRUY VẤN THAM SỐ HÓA (PARAMETERIZED QUERY)
-- Đây là câu lệnh SQL mà EF Core LINQ tự động sinh ra và thực thi trên SQL Server
-- ------------------------------------------------------------------------------
EXEC sp_executesql 
    N'SELECT * FROM Accounts WHERE Username = @p0 AND Password = @p1',
    N'@p0 nvarchar(50), @p1 nvarchar(50)',
    @p0 = N''' OR ''1''=''1'' --',
    @p1 = N'mat_khau_bat_ky';
GO
-- KẾT QUẢ: SQL Server trả về 0 dòng vì chuỗi payload chỉ được coi là text thuần!
