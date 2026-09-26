# KỊCH BẢN THUYẾT TRÌNH & THỰC NGHIỆM OWASP TOP 10 (22 SLIDES - 3 THÀNH VIÊN)
**Học phần:** Lập trình Web nâng cao | **Giảng viên hướng dẫn:** ThS. Lê Hữu Dũng  
**Nhóm sinh viên thực hiện:** WNC.G01 (Đề tài BTL: Website cung cấp dịch vụ tư vấn trực tuyến)  
**Thời lượng chuẩn:** 7 - 9 phút (Mỗi thành viên trình bày Slide & Live Demo 2 - 2.5 phút)

---

## 🎤 PHẦN MỞ ĐẦU (Nguyễn Danh Học - 45 giây) — Chiếu Slide 1 & 2
- *"Kính thưa thầy Lê Hữu Dũng và các bạn sinh viên, hôm nay nhóm WNC.G01 xin trình bày báo cáo và biểu diễn thực nghiệm trực tiếp **Bộ 3 lỗ hổng bảo mật Web kinh điển nhất theo tiêu chuẩn OWASP Top 10**."*
- *"Quán triệt đúng phương châm chỉ đạo của thầy: **Không đọc lý thuyết dịch sách vở, mà phải hiểu sâu cách hacker tấn công và thực nghiệm đối chứng rõ ràng**, nhóm em đã xây dựng một nền tảng thực nghiệm độc lập trên ASP.NET Core .NET 10 kết nối trực tiếp CSDL Microsoft SQL Server 2025."*
- *"Để đảm bảo khối lượng công việc được phân chia đồng đều và nghiêm túc, nhóm chia làm 3 chuyên đề độc lập:*
  1. *Nguyễn Danh Học: Thực nghiệm **SQL Injection** (OWASP A03:2021) — Can thiệp tầng CSDL SQL Server 2025.*
  2. *Nguyễn Thanh Bình: Thực nghiệm **Stored XSS** (OWASP A03:2021) — Tấn công qua tính năng nhận xét chuyên gia.*
  3. *Nguyễn Minh Cường: Thực nghiệm **CSRF Attack** (OWASP A01:2021) — Tấn công giả mạo chuyển tiền ví ngầm.*"

---

## 👤 CHUYÊN ĐỀ 1: SQL INJECTION — NGUYỄN DANH HỌC (2.5 phút)
### 1. Trình bày lý thuyết & phân tích kỹ thuật (Chiếu Slide 3, 4, 5 — 1 phút)
- **Slide 3 (Bản chất lỗi):** *"Nguyên nhân gốc rễ của SQL Injection chỉ nằm ở 1 điểm: Lập trình viên nhầm lẫn giữa Code và Data, dùng chuỗi cộng trực tiếp input người dùng vào câu SQL. Khi hacker truyền vào payload `' OR '1'='1' --`, cấu trúc cú pháp của câu lệnh bị bẻ gãy hoàn toàn."*
- **Slide 4 (Độ nguy hiểm CVSS 9.8):** *"Lỗ hổng này đạt mức độ CRITICAL (9.8/10) vì 3 hậu quả nghiêm trọng: Bypass đăng nhập Admin không cần mật khẩu, Rò rỉ toàn bộ CSDL bảng Accounts và nguy cơ bị xóa sổ CSDL bằng lệnh DROP TABLE."*
- **Slide 5 (Giải phẫu payload):** *"Khi đưa `' OR '1'='1' --` vào ô Username: Dấu nháy đơn đóng chuỗi sớm, `OR 1=1` biến điều kiện WHERE thành luôn đúng cho mọi dòng, và dấu `--` biến toàn bộ vế kiểm tra mật khẩu phía sau thành comment vô hiệu."*

### 2. Biểu diễn thực nghiệm trực tiếp (Slide 6 & 7 — Chuyển Live Demo 1.5 phút)
- **Thao tác Web UI (Tab 1):**
  - Bấm nạp Payload 1 `' OR '1'='1' --` -> Bấm nút đỏ *Gửi Request Khai Thác*:
    - *"Thưa thầy, ngay lập tức hệ thống đăng nhập thành công với quyền Admin, và toàn bộ 4 tài khoản CSDL cùng số dư tiền, mật khẩu plain-text đều bị trích xuất hiển thị ra màn hình."*
  - Bấm nút xanh *Gửi Request Kiểm Thử* (Nhánh đã phòng thủ):
    - *"Hệ thống chặn đứng hoàn toàn, thông báo 'Sai tài khoản/mật khẩu' vì EF Core đã dùng Parameterized Query (`@p0`)."*
- **Thao tác Tầng Sâu (Slide 7):**
  - Mở Terminal chạy: `./test_exploit_cli.sh` -> Server trả JSON rò rỉ CSDL, chứng minh lỗi nằm ở tầng C# Backend.
  - Mở VS Code chạy file `verify_sql_server.sql` đối chiếu trực tiếp trên máy chủ SQL Server 2025.
- **Chốt phòng thủ (Slide 8):** *"Dùng LINQ Parameterized trong EF Core và nguyên tắc đặc quyền tối thiểu (Least Privilege). Sau đây xin mời bạn Nguyễn Thanh Bình tiếp tục với Stored XSS."*

---

## 👤 CHUYÊN ĐỀ 2: STORED XSS — NGUYỄN THANH BÌNH (2.5 phút)
### 1. Trình bày lý thuyết & phân tích kỹ thuật (Chiếu Slide 9, 10, 11 — 1 phút)
- **Slide 9 (Ngữ cảnh nghiệp vụ):** *"Kính thưa thầy, em là Nguyễn Thanh Bình. Trong website tư vấn trực tuyến, tính năng đánh giá chuyên gia là nơi khách hàng tương tác trực tiếp. Lỗ hổng Stored XSS xảy ra khi máy chủ lưu nguyên văn chuỗi script của hacker vào CSDL và View render bằng `@Html.Raw()`."*
- **Slide 10 (Độ nguy hiểm):** *"Stored XSS nguy hiểm ở 3 điểm: Đánh cắp Cookie phiên làm việc (`document.cookie`), Thay đổi toàn bộ nội dung website (DOM Defacement) và lừa người dùng nhập lại thông tin thẻ ngân hàng."*
- **Slide 11 (3 Kịch bản payload):** *"Kẻ tấn công có thể dùng thẻ `<script>`, hoặc bypass bộ lọc bằng thẻ ảnh `<img src=x onerror=...>`, hoặc dùng script chiếm quyền điều khiển trang web."*

### 2. Biểu diễn thực nghiệm trực tiếp (Slide 12 & 13 — Chuyển Live Demo 1.5 phút)
- **Thao tác Web UI (Tab 2 - `/Xss`):**
  - Bấm nạp Payload 1 `<script>alert('Lộ Cookie: ' + document.cookie);</script>` -> Bấm *Đăng Đánh Giá (Vulnerable)*:
    - *"Trình duyệt lập tức kích hoạt mã script, bật popup alert hiển thị chuỗi Cookie phiên `AuthSessionToken` của nạn nhân!"*
  - Bấm nút *Đăng Đánh Giá (Secure)*:
    - *"Nhánh an toàn sử dụng cơ chế tự động HTML Encoding của Razor View (`@comment.Content`). Ký tự `<` và `>` được mã hóa thành `&lt;` và `&gt;`. Đoạn script chỉ hiện ra dưới dạng chữ thường, hoàn toàn vô hại."*
- **Minh chứng CSDL & Cookie (Slide 13):**
  - Mở bảng `dbo.Comments` cho thấy đoạn script được lưu vĩnh viễn trong CSDL.
  - Phân tích cờ `HttpOnly = true` của Cookie để ngăn chặn triệt để JavaScript đọc trộm Session.
- **Chốt phòng thủ (Slide 14):** *"Không dùng `@Html.Raw()`, bật cờ `HttpOnly = true` và triển khai Content Security Policy (CSP Header). Xin mời bạn Nguyễn Minh Cường tiếp tục với CSRF Attack."*

---

## 👤 CHUYÊN ĐỀ 3: CSRF ATTACK — NGUYỄN MINH CƯỜNG (2.5 phút)
### 1. Trình bày lý thuyết & phân tích kỹ thuật (Chiếu Slide 15 & 16 — 1 phút)
- **Slide 15 (Bản chất CSRF):** *"Kính thưa thầy, em là Nguyễn Minh Cường. CSRF là kỹ thuật mượn quyền của người dùng hợp lệ. Khi nạn nhân đang đăng nhập hệ thống tư vấn, trình duyệt đã lưu cookie hợp lệ. Hacker lừa nạn nhân bấm vào trang web bẫy của hacker ở tab bên cạnh."*
- **Slide 16 (Kịch bản bẫy trúng thưởng):** *"Trang web bẫy (`AttackerSite.cshtml`) chứa form ẩn tự động POST sang hệ thống mục tiêu. Máy chủ thấy Cookie hợp lệ của nạn nhân tự động gửi kèm nên thực hiện lệnh trừ tiền mà không biết request đó bắt nguồn từ trang web của hacker."*

### 2. Biểu diễn thực nghiệm trực tiếp (Slide 17 & 18 — Chuyển Live Demo 1.5 phút)
- **Thao tác Web UI (Tab 3 - `/Csrf`):**
  - Cho thầy xem số dư ban đầu: Ví Nạn nhân có 50 triệu VNĐ, ví Hacker có 0 VNĐ.
  - Bấm nút đỏ *Mở trang web bẫy của Hacker* -> Bấm *Nhận thưởng iPhone*:
    - *"Ngay lập tức, số dư ví nạn nhân bị trừ 20 triệu VNĐ chuyển sang ví của hacker!"*
  - Quay lại hệ thống, thử form bên phải có Token:
    - *"Giao dịch chính chủ có Anti-Forgery Token được bảo vệ an toàn."*
- **Minh chứng CSDL & Token (Slide 18 & 19):**
  - Đối chiếu bảng `dbo.UserWallets` trong CSDL SQL Server 2025 thấy rõ số dư biến động từ 50 triệu xuống 30 triệu.
  - Phân tích cơ chế Synchronizer Token Pattern (`@Html.AntiForgeryToken()` và `[ValidateAntiForgeryToken]`).
- **Phòng thủ trong AJAX Buổi 9 (Slide 20):** *"Truyền token qua Header `RequestVerificationToken` trong Fetch API."*

---

## 🎯 TỔNG KẾT & CAM KẾT BTL (Nguyễn Danh Học - 45 giây) — Chiếu Slide 21 & 22
- **Slide 21 (Bảng so sánh):** *"Tổng kết lại: Cả 3 lỗ hổng đều nguy hiểm nhưng có cơ chế và tầng phòng thủ khác nhau: SQLi ở tầng CSDL, XSS ở tầng DOM Trình duyệt, và CSRF ở tầng xác thực Request."*
- **Slide 22 (Cam kết đồ án BTL):** *"Trong đồ án Website Tư vấn trực tuyến của nhóm WNC.G01, chúng em cam kết tích hợp đầy đủ cả 3 tầng bảo vệ: EF Core LINQ, Data Annotations / Razor Encoding và Anti-Forgery Token để bảo vệ an toàn tuyệt đối cho người dùng."*
- *"Nhóm WNC.G01 xin trân trọng cảm ơn thầy và các bạn đã lắng nghe!"*
