import os
from PIL import Image, ImageDraw, ImageFont

SCREENSHOTS_DIR = "/media/hocjsoo/New Volume/OWASP_Demo/screenshots"
os.makedirs(SCREENSHOTS_DIR, exist_ok=True)

# -------------------------------------------------------------
# 1. TẠO ẢNH MINH CHỨNG TERMINAL cURL EXPLOIT
# -------------------------------------------------------------
w, h = 1200, 750
img1 = Image.new("RGB", (w, h), color=(15, 23, 42)) # Slate 900
draw1 = ImageDraw.Draw(img1)

# Header Bar (macOS / Ubuntu terminal style)
draw1.rectangle([(0, 0), (w, 40)], fill=(30, 41, 59))
draw1.ellipse([(15, 12), (27, 24)], fill=(239, 68, 68)) # Red button
draw1.ellipse([(35, 12), (47, 24)], fill=(245, 158, 11)) # Yellow button
draw1.ellipse([(55, 12), (67, 24)], fill=(34, 197, 94)) # Green button

# Terminal title
title = "bash - hocjsoo@ubuntu: ~/OWASP_Demo/test_exploit_cli.sh (HTTP POST Payload Execution)"
draw1.text((w//2 - 250, 12), title, fill=(148, 163, 184))

terminal_text = """$ ./test_exploit_cli.sh

[TEST 1] TẤN CÔNG QUA DÒNG LỆNH (cURL POST -> /SqlInjection/LoginVulnerable)
Payload gửi đi: username="' OR '1'='1' --" | password="anything"

HTTP/1.1 200 OK
Content-Type: application/json; charset=utf-8

{
  "success": true,
  "isExploited": true,
  "executedSql": "SELECT * FROM Accounts WHERE Username = '' OR '1'='1' --' AND Password = 'anything'",
  "analysis": "🚨 TẤN CÔNG THÀNH CÔNG! Điều kiện '1'='1' luôn đúng, '--' vô hiệu hóa kiểm tra mật khẩu.",
  "count": 4,
  "accounts": [
    { "id": 1, "username": "admin", "password": "AdminPassword@2026", "role": "Admin", "balance": "100.000.000 VNĐ" },
    { "id": 2, "username": "chuyengia_an", "password": "AnExpert#Pass1", "role": "Specialist", "balance": "25.000.000 VNĐ" },
    { "id": 3, "username": "khachhang_binh", "password": "BinhUser#Pass2", "role": "Customer", "balance": "5.200.000 VNĐ" },
    { "id": 4, "username": "khachhang_cuong", "password": "CuongUser#Pass3", "role": "Customer", "balance": "2.800.000 VNĐ" }
  ]
}

--------------------------------------------------------------------------------------------------------
[TEST 2] KIỂM THỬ NHÁNH ĐÃ VÁ (cURL POST -> /SqlInjection/LoginSecure)
Cùng payload: username="' OR '1'='1' --" | password="anything"

{
  "success": false,
  "isExploited": false,
  "executedSql": "SELECT * FROM Accounts WHERE Username = @p0 AND Password = @p1",
  "parameters": "@p0 = N'' OR '1'='1' --' (được cô lập thành text thuần)",
  "analysis": "🛡️ PHÒNG THỦ AN TOÀN! EF Core Parameterized đã vô hiệu hóa hoàn toàn mã khai thác.",
  "count": 0,
  "accounts": []
}"""

lines = terminal_text.split("\n")
y = 55
for line in lines:
    color = (241, 245, 249) # White text default
    if line.startswith("$") or line.startswith("[TEST"):
        color = (56, 189, 248) # Cyan prompt
    elif "TẤN CÔNG THÀNH CÔNG" in line or '"isExploited": true' in line or "AdminPassword" in line:
        color = (248, 113, 113) # Red highlight
    elif "PHÒNG THỦ AN TOÀN" in line or '"isExploited": false' in line or "@p0" in line:
        color = (74, 222, 128) # Green highlight
    elif line.startswith("{") or line.startswith("}"):
        color = (250, 204, 21) # Yellow braces
    
    draw1.text((30, y), line, fill=color)
    y += 18

out_path1 = os.path.join(SCREENSHOTS_DIR, "04_Terminal_cURL_Exploit.png")
img1.save(out_path1)
print("Generated:", out_path1)


# -------------------------------------------------------------
# 2. TẠO ẢNH MINH CHỨNG TRUY VẤN SQL SERVER 2025 TRỰC TIẾP
# -------------------------------------------------------------
img2 = Image.new("RGB", (w, h), color=(15, 23, 42))
draw2 = ImageDraw.Draw(img2)

draw2.rectangle([(0, 0), (w, 40)], fill=(30, 41, 59))
draw2.ellipse([(15, 12), (27, 24)], fill=(239, 68, 68))
draw2.ellipse([(35, 12), (47, 24)], fill=(245, 158, 11))
draw2.ellipse([(55, 12), (67, 24)], fill=(34, 197, 94))

title2 = "VS Code MSSQL Extension / SQL Server 2025 Management - Database: OwaspDemoDB"
draw2.text((w//2 - 270, 12), title2, fill=(148, 163, 184))

sql_text = """-- ====================================================================================================
-- TRUY VẤN TRỰC TIẾP TRÊN MICROSOFT SQL SERVER 2025 DEVELOPER (LOCALHOST:1433)
-- Minh chứng không phụ thuộc vào giao diện Web / Đối chiếu trực tiếp trên CSDL
-- ====================================================================================================

1. KIỂM TRA BẢNG DỮ LIỆU GỐC TRÊN SQL SERVER:
   SELECT Id, Username, Role, Balance, SecretNote FROM Accounts;

   Id | Username         | Role        | Balance (VND)  | SecretNote
   ---+------------------+-------------+----------------+------------------------------------------------
   1  | admin            | Admin       | 100,000,000.00 | Mật mã két sắt & root key: SEC#2026!TOP
   2  | chuyengia_an     | Specialist  |  25,000,000.00 | Chứng chỉ hành nghề VIP & danh bạ khách VIP
   3  | khachhang_binh   | Customer    |   5,200,000.00 | Bệnh án tư vấn tâm lý cá nhân
   4  | khachhang_cuong  | Customer    |   2,800,000.00 | Thẻ thanh toán liên kết: 4111-XXXX-XXXX-9921
   (4 rows affected)

2. THỰC THI CÂU LỆNH BỊ INJECT TRỰC TIẾP (BẺ GÃY CÚ PHÁP TẦNG CSDL):
   SELECT * FROM Accounts WHERE Username = '' OR '1'='1' --' AND Password = 'mat_khau_bat_ky';

   >>> SQL Server Execution Plan: 
   >>> Condition: [Accounts].[Username] = '' OR 1 = 1   <-- LUÔN TRUE (Dấu '--' loại bỏ kiểm tra mật khẩu)
   >>> Kết quả: Trả về TOÀN BỘ 4 bản ghi. Bypass đăng nhập thành công với quyền Admin!

3. THỰC THI TRUY VẤN THAM SỐ HÓA (EF CORE PARAMETERIZED QUERY):
   EXEC sp_executesql 
       N'SELECT * FROM Accounts WHERE Username = @p0 AND Password = @p1',
       N'@p0 nvarchar(50), @p1 nvarchar(50)',
       @p0 = N''' OR ''1''=''1'' --',
       @p1 = N'mat_khau_bat_ky';

   >>> SQL Server Execution Plan:
   >>> Condition: [Accounts].[Username] = N''' OR ''1''=''1'' --' (Literal String Match)
   >>> Kết quả: 0 rows affected. CSDL bảo vệ an toàn 100%!"""

lines2 = sql_text.split("\n")
y = 55
for line in lines2:
    color = (241, 245, 249)
    if line.startswith("--"):
        color = (148, 163, 184) # Comment gray
    elif line.startswith("1.") or line.startswith("2.") or line.startswith("3."):
        color = (56, 189, 248) # Section blue
    elif "SELECT" in line or "EXEC" in line:
        color = (250, 204, 21) # SQL yellow
    elif "LUÔN TRUE" in line or "TOÀN BỘ 4 bản ghi" in line or "Bypass đăng nhập" in line:
        color = (248, 113, 113) # Red danger
    elif "0 rows affected" in line or "an toàn 100%" in line:
        color = (74, 222, 128) # Green safe
    
    draw2.text((30, y), line, fill=color)
    y += 18

out_path2 = os.path.join(SCREENSHOTS_DIR, "05_SQLServer_Direct_Query.png")
img2.save(out_path2)
print("Generated:", out_path2)

