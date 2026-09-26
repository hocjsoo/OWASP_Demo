# KỊCH BẢN THUYẾT TRÌNH & THỰC NGHIỆM OWASP (PHÂN CÔNG 3 THÀNH VIÊN)
**Học phần:** Lập trình Web nâng cao | **Giảng viên:** ThS. Lê Hữu Dũng  
**Nhóm:** WNC.G01 (Đề tài: Website cung cấp dịch vụ tư vấn trực tuyến)  
**Thời lượng:** 5 - 6 phút (Mỗi thành viên trình bày và demo 1.5 - 2 phút)

---

## 🎤 MỞ ĐẦU (Nguyễn Danh Học - 30 giây)
- *"Kính thưa thầy và các bạn, hôm nay nhóm WNC.G01 xin trình bày và thực nghiệm trực tiếp **Bộ 3 lỗ hổng bảo mật Web kinh điển nhất theo chuẩn OWASP Top 10**."*
- *"Để đảm bảo tính khách quan và đóng góp đồng đều, nhóm em chia đều 3 chủ đề cho 3 thành viên:*
  1. *Nguyễn Danh Học: Thực nghiệm **SQL Injection** (A03:2021).*
  2. *Nguyễn Thanh Bình: Thực nghiệm **Stored XSS** (A03:2021).*
  3. *Nguyễn Minh Cường: Thực nghiệm **CSRF Attack** (A01:2021).*
- *"Chúng em sẽ đi thẳng vào thực nghiệm sống (Live Demo) trên chính ứng dụng ASP.NET Core kết nối SQL Server 2025 do nhóm tự xây dựng."*

---

## 👤 PHẦN 1: NGUYỄN DANH HỌC — SQL INJECTION (1.5 phút)
### 1. Thuyết trình trên Slide (Slide 2 - 30 giây)
- *"Phần 1 do em - Nguyễn Danh Học phụ trách về lỗ hổng **SQL Injection**."*
- *"Nguyên nhân cốt lõi: Lập trình viên nối chuỗi dữ liệu người dùng trực tiếp vào câu SQL. Khi hacker truyền vào payload `' OR '1'='1' --`, dấu nháy đơn đóng chuỗi sớm, mệnh đề `OR 1=1` luôn đúng và dấu `--` loại bỏ hoàn toàn bước kiểm tra mật khẩu."*
- *"Hậu quả: Hacker đăng nhập thẳng vào tài khoản Admin mà không cần mật khẩu, đánh cắp toàn bộ cơ sở dữ liệu."*

### 2. Thực nghiệm sống trên máy (Alt + Tab sang Web UI Tab 1 - 1 phút)
- **Bước 1 (Khai thác):** Chọn nút Payload 1 `' OR '1'='1' --` -> Bấm nút đỏ *Gửi Request Khai Thác*:
  - *"Thưa thầy, hệ thống lập tức bypass đăng nhập, rò rỉ toàn bộ 4 tài khoản cùng số dư tiền và mật khẩu gốc."*
- **Bước 2 (Phòng thủ):** Đưa cùng payload sang bên Xanh -> Bấm nút xanh *Gửi Request Kiểm Thử*:
  - *"Hệ thống chặn đứng ngay, báo 'Sai tài khoản/mật khẩu'. Lý do: Nhánh an toàn sử dụng **EF Core LINQ**, SQL Server sử dụng cơ chế tham số hóa Parameterized Query (`@p0`), cô lập payload thành chuỗi văn bản thuần túy."*
- *"Sau đây xin mời bạn Nguyễn Thanh Bình tiếp tục với phần Stored XSS."*

---

## 👤 PHẦN 2: NGUYỄN THANH BÌNH — STORED XSS (1.5 phút)
### 1. Thuyết trình trên Slide (Slide 3 - 30 giây)
- *"Kính thưa thầy, em là Nguyễn Thanh Bình, phụ trách chủ đề **Stored XSS (Cross-Site Scripting lưu trữ)**."*
- *"Ngữ cảnh bài toán: Trong website tư vấn trực tuyến của nhóm, khách hàng có chức năng gửi nhận xét và đánh giá chuyên gia."*
- *"Bản chất lỗi: Nếu Server lưu nguyên văn chuỗi script của người dùng và View sử dụng `@Html.Raw()` để hiển thị, trình duyệt của mọi người dùng khác khi truy cập trang web sẽ tự động thực thi đoạn mã độc của hacker."*
- *"Hậu quả: Đánh cắp Cookie phiên làm việc (`document.cookie`), giả mạo danh tính hoặc chuyển hướng người dùng sang trang web lừa đảo."*

### 2. Thực nghiệm sống trên máy (Alt + Tab sang Web UI Tab 2 - 1 phút)
- **Bước 1 (Khai thác):** Bấm nút Payload 1 `<script>alert('Lộ Cookie: ' + document.cookie);</script>` -> Bấm *Đăng Đánh Giá (Vulnerable)*:
  - *"Ngay khi bài viết được đăng, trình duyệt lập tức kích hoạt mã script, bật popup alert hiển thị Cookie nhạy cảm `AuthSessionToken` của nạn nhân."*
- **Bước 2 (Phòng thủ):** Bấm nút *Đăng Đánh Giá (Secure)*:
  - *"Nhánh an toàn sử dụng cơ chế tự động **HTML Encoding của Razor View** (`@comment.Content`). Toàn bộ ký tự `<` và `>` được chuyển thành `&lt;` và `&gt;`. Đoạn script chỉ hiện ra dưới dạng chữ thường, hoàn toàn vô hại."*
- *"Sau đây xin mời bạn Nguyễn Minh Cường tiếp tục với phần CSRF Attack."*

---

## 👤 PHẦN 3: NGUYỄN MINH CƯỜNG — CSRF ATTACK (1.5 phút)
### 1. Thuyết trình trên Slide (Slide 4 - 30 giây)
- *"Kính thưa thầy, em là Nguyễn Minh Cường, phụ trách chủ đề **Cross-Site Request Forgery (CSRF - Giả mạo yêu cầu từ trang khác)**."*
- *"Bản chất lỗi: Khi nạn nhân đã đăng nhập vào hệ thống tư vấn trực tuyến, trình duyệt của nạn nhân đã lưu cookie phiên làm việc hợp lệ. Kẻ tấn công lừa nạn nhân mở một trang web bẫy (ví dụ: trang trúng thưởng iPhone)."*
- *"Cơ chế tấn công: Trang web bẫy tự động kích hoạt một form POST ngầm chuyển tiền sang ví của hacker. Vì cookie của nạn nhân tự động được đính kèm theo request, máy chủ tưởng đây là yêu cầu chính chủ và thực hiện trừ tiền."*

### 2. Thực nghiệm sống trên máy (Alt + Tab sang Web UI Tab 3 - 1 phút)
- **Bước 1 (Xem số dư ban đầu):** *"Số dư ví nạn nhân hiện tại là 50.000.000 VNĐ, ví hacker là 0 VNĐ."*
- **Bước 2 (Khai thác):** Bấm vào nút đỏ *Mở trang web bẫy của Hacker (Attacker Site)* -> Bấm *Nhận thưởng*:
  - *"Ngay lập tức, số dư của nạn nhân bị trừ 20.000.000 VNĐ và chuyển thẳng sang ví của hacker!"*
- **Bước 3 (Phòng thủ):** Quay lại hệ thống, chỉ sang Form bên phải:
  - *"Để phòng thủ, ASP.NET Core sử dụng cơ chế **Anti-Forgery Token** với thuộc tính `[ValidateAntiForgeryToken]` và thẻ `@Html.AntiForgeryToken()`. Kẻ tấn công ở trang web khác không thể đọc được token bí mật này, mọi request giả mạo đều bị chặn đứng với mã lỗi 400 Bad Request."*

---

## 🎯 TỔNG KẾT & CAM KẾT ĐỒ ÁN (Nguyễn Danh Học - 30 giây) — Slide 5
- *"Kính thưa thầy, thông qua 3 bài thực nghiệm trực tiếp:*
  1. *Nguyễn Danh Học chứng minh phòng thủ SQL Injection bằng EF Core LINQ.*
  2. *Nguyễn Thanh Bình chứng minh phòng thủ XSS bằng Razor HTML Encoding.*
  3. *Nguyễn Minh Cường chứng minh phòng thủ CSRF bằng Anti-Forgery Token.*
- *"Trong đồ án Bài tập lớn Website Tư vấn trực tuyến của nhóm WNC.G01, chúng em cam kết tích hợp đầy đủ cả 3 lớp phòng thủ này để bảo vệ dữ liệu người dùng một cách an toàn tuyệt đối."*
- *"Nhóm WNC.G01 xin trân trọng cảm ơn thầy đã theo dõi!"*
