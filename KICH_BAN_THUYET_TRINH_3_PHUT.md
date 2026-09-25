# KỊCH BẢN THUYẾT TRÌNH & DEMO THỰC NGHIỆM (3 - 5 PHÚT)
**Chủ đề:** SQL Injection (OWASP A03:2021) — Môn Lập trình Web nâng cao  
**Giảng viên:** ThS. Lê Hữu Dũng | **Nhóm:** WNC.G01 (Nguyễn Danh Học trình bày)

---

### PHẦN 1: MỞ ĐẦU (30 giây) — Chiếu Slide 1 & 2
- *"Kính thưa thầy và các bạn, hôm nay nhóm WNC.G01 xin trình bày và thực nghiệm trực tiếp lỗ hổng bảo mật **SQL Injection** theo tiêu chuẩn **OWASP Top 10 - A03:2021**."*
- *"Chúng em không đọc lý thuyết dịch sách vở, mà sẽ đi thẳng vào: **Bản chất lỗi**, **Độ nguy hiểm thực tế**, và **Thực nghiệm bắn payload tấn công trực tiếp** trên ứng dụng ASP.NET Core kết nối CSDL."*
- *"Nguyên nhân gốc rễ của SQL Injection chỉ nằm ở 1 điểm: **Lập trình viên không tách biệt giữa Dữ liệu (Data) và Câu lệnh (Code)**, mà dùng chuỗi cộng trực tiếp dữ liệu người dùng vào câu SQL."*

---

### PHẦN 2: CHỨNG MINH ĐỘ NGUY HIỂM (30 giây) — Chiếu Slide 3
- *"Lỗ hổng này cực kỳ nguy hiểm vì 3 cấp độ tác động:*
  1. ***Bypass xác thực (Authentication Bypass):** Kẻ tấn công không cần biết mật khẩu vẫn chiếm quyền Quản trị viên (Admin).
  2. ***Rò rỉ toàn bộ CSDL (Data Exfiltration):** Trích xuất mật khẩu, số dư tài khoản, hồ sơ khách hàng.
  3. ***Phá hoại máy chủ:** Kẻ xấu có thể DROP TABLE hoặc gọi lệnh can thiệp hệ điều hành.*"

---

### PHẦN 3: THỰC NGHIỆM DEMO TRỰC TIẾP TRÊN MÁY (2 phút) — Mở Trình Duyệt Web
*(Mở trang web `http://127.0.0.1:5076/SqlInjection` chia 2 cột đối chứng)*

#### Bước 3.1: Demo Tấn Công (Cột Đỏ - Vulnerable)
1. Chỉ vào ô code trên màn hình:
   - *"Ở nhánh Bị Lỗi bên trái, em viết câu lệnh ghép chuỗi thuần: `SELECT * FROM Accounts WHERE Username = '{user}' AND Password = '{pass}'`."*
2. Nhập payload: `' OR '1'='1' --` (hoặc bấm nút màu đỏ trên thanh công cụ). Mật khẩu: gõ bừa hoặc để trống.
3. Bấm nút **Gửi Request Khai Thác**:
   - *"Thưa thầy, ngay lập tức hệ thống đăng nhập thành công với quyền Admin, và toàn bộ 4 tài khoản trong CSDL cùng số dư tiền, mật khẩu plain-text đều bị trích xuất hiển thị ra màn hình."*
   - *"Giải thích câu lệnh SQL thực tế vừa chạy (chỉ vào khung màu vàng phía dưới):*
     - *Dấu nháy đơn `'` đóng chuỗi sớm.*
     - *Mệnh đề `OR '1'='1'` biến điều kiện thành luôn ĐÚNG (True) cho mọi bản ghi.*
     - *Ký tự `--` biến toàn bộ phần kiểm tra mật khẩu phía sau thành comment vô hiệu lực."*

#### Bước 3.2: Demo Phòng Thủ (Cột Xanh - Secure)
1. Chỉ sang cột An Toàn bên phải:
   - *"Bây giờ, em đưa **chính xác payload `' OR '1'='1' --` đó** sang nhánh Đã Vá (dùng EF Core LINQ: `_context.Accounts.Where(a => a.Username == user && a.Password == pass)`)."*
2. Bấm nút **Gửi Request Kiểm Thử**:
   - *"Kết quả: Hệ thống chặn đứng hoàn toàn, thông báo 'Sai tài khoản hoặc mật khẩu'."*
   - *"Lý do an toàn (chỉ vào khung câu truy vấn màu xanh): SQL Server đã dùng tham số hóa (Parameterized Query với `@p0`, `@p1`). Cây truy vấn được biên dịch trước, và chuỗi payload của hacker chỉ được coi là một chuỗi văn bản thuần túy, hoàn toàn không thể làm biến đổi cấu trúc lệnh."*

---

### PHẦN 4: KẾT LUẬN & KHUYẾN NGHỊ (30 giây) — Chiếu Slide 6
- *"Tóm lại, để phòng thủ triệt để SQL Injection trong ASP.NET Core MVC:*
  1. *Tuyệt đối không ghép chuỗi SQL thủ công. Luôn dùng ORM (EF Core) hoặc Parameterized Query.*
  2. *Áp dụng nguyên tắc đặc quyền tối thiểu (Least Privilege) cho tài khoản kết nối CSDL.*
  3. *Validate dữ liệu đầu vào và luôn mã hóa một chiều mật khẩu với ASP.NET Core Identity.*"
- *"Nhóm chúng em xin kết thúc bài thực nghiệm, cảm ơn thầy và các bạn đã lắng nghe!"*
