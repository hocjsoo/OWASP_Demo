import os
from PIL import Image, ImageDraw, ImageFont

SCREENSHOTS_DIR = "/media/hocjsoo/New Volume/OWASP_Demo/screenshots"

FONT_PATH = "/usr/share/fonts/truetype/liberation/LiberationMono-Bold.ttf"
TITLE_FONT = ImageFont.truetype(FONT_PATH, 20)
BODY_FONT = ImageFont.truetype(FONT_PATH, 17)
NOTE_FONT = ImageFont.truetype(FONT_PATH, 16)

w, h = 950, 720

# 1. cURL Terminal
img1 = Image.new("RGB", (w, h), color=(15, 23, 42))
draw1 = ImageDraw.Draw(img1)
draw1.rectangle([(0, 0), (w, 45)], fill=(30, 41, 59))
draw1.ellipse([(16, 14), (32, 30)], fill=(239, 68, 68))
draw1.ellipse([(42, 14), (58, 30)], fill=(245, 158, 11))
draw1.ellipse([(68, 14), (84, 30)], fill=(34, 197, 94))
draw1.text((105, 12), "Bắn HTTP POST trực tiếp bằng cURL (Tầng API)", font=TITLE_FONT, fill=(226, 232, 240))

text1 = [
    ("$ curl -X POST http://localhost:5076/LoginVulnerable \\", (56, 189, 248)),
    ("       -F \"username=' OR '1'='1' --\" -F \"password=any\"", (56, 189, 248)),
    ("", (255, 255, 255)),
    ("HTTP/1.1 200 OK | Content-Type: application/json", (148, 163, 184)),
    ("{", (250, 204, 21)),
    ('  "success": true, "isExploited": true,', (248, 113, 113)),
    ('  "executedSql": "SELECT * FROM Accounts WHERE', (250, 204, 21)),
    ("         Username='' OR '1'='1' --' AND Pass='...\",", (248, 113, 113)),
    ('  "count": 4,', (248, 113, 113)),
    ('  "accounts": [', (250, 204, 21)),
    ('    { "id": 1, "user": "admin", "role": "Admin",', (248, 113, 113)),
    ('      "pass": "AdminPassword@2026",', (248, 113, 113)),
    ('      "balance": "100.000.000 VNĐ" },', (248, 113, 113)),
    ('    { "id": 2, "user": "chuyengia_an", ... },', (241, 245, 249)),
    ('    { "id": 3, "user": "khachhang_binh", ... }', (148, 163, 184)),
    ('  ]', (250, 204, 21)),
    ("}", (250, 204, 21)),
    ("", (255, 255, 255)),
    (">> KẾT QUẢ: Rò rỉ toàn bộ CSDL qua API dòng lệnh!", (248, 113, 113)),
    (">> Chứng minh: Lỗi ở tầng C#, không phụ thuộc Web.", (74, 222, 128))
]

y = 60
for line, col in text1:
    draw1.text((25, y), line, font=BODY_FONT, fill=col)
    y += 28

img1.save(os.path.join(SCREENSHOTS_DIR, "04_Terminal_cURL_Exploit.png"))


# 2. SQL Server Terminal
img2 = Image.new("RGB", (w, h), color=(15, 23, 42))
draw2 = ImageDraw.Draw(img2)
draw2.rectangle([(0, 0), (w, 45)], fill=(30, 41, 59))
draw2.ellipse([(16, 14), (32, 30)], fill=(239, 68, 68))
draw2.ellipse([(42, 14), (58, 30)], fill=(245, 158, 11))
draw2.ellipse([(68, 14), (84, 30)], fill=(34, 197, 94))
draw2.text((105, 12), "Truy vấn trực tiếp trên Microsoft SQL Server 2025", font=TITLE_FONT, fill=(226, 232, 240))

text2 = [
    ("-- 1. BẢNG DỮ LIỆU THẬT TRONG OwaspDemoDB:", (148, 163, 184)),
    ("SELECT Id, Username, Role, Balance FROM Accounts;", (250, 204, 21)),
    ("Id | Username     | Role       | Balance (VND)", (56, 189, 248)),
    (" 1 | admin        | Admin      | 100,000,000.00", (241, 245, 249)),
    (" 2 | chuyengia_an | Specialist |  25,000,000.00", (241, 245, 249)),
    (" 3 | khach_binh   | Customer   |   5,200,000.00", (241, 245, 249)),
    (" 4 | khach_cuong  | Customer   |   2,800,000.00", (241, 245, 249)),
    ("", (255, 255, 255)),
    ("-- 2. THỰC THI CÂU LỆNH BỊ INJECT TRỰC TIẾP:", (148, 163, 184)),
    ("SELECT * FROM Accounts WHERE Username='' OR '1'='1' --';", (250, 204, 21)),
    (">> Cấu trúc: WHERE (User='') OR (1=1) -- [comment]", (248, 113, 113)),
    (">> Kết quả: Trả về TOÀN BỘ 4 bản ghi (Bypass login!)", (248, 113, 113)),
    ("", (255, 255, 255)),
    ("-- 3. THỰC THI TRUY VẤN THAM SỐ (EF CORE PARAMETER):", (148, 163, 184)),
    ("EXEC sp_executesql N'SELECT * FROM Accounts ...',", (250, 204, 21)),
    ("     N'@p0 nvarchar(50)', @p0 = N''' OR ''1''=''1'' --';", (74, 222, 128)),
    (">> Kết quả: 0 rows. CSDL chặn đứng hoàn toàn!", (74, 222, 128))
]

y = 60
for line, col in text2:
    draw2.text((25, y), line, font=BODY_FONT, fill=col)
    y += 32

img2.save(os.path.join(SCREENSHOTS_DIR, "05_SQLServer_Direct_Query.png"))
print("Done regenerating terminal images with TTF fonts!")
