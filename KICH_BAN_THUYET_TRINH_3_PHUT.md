# KỊCH BẢN THUYẾT TRÌNH & DEMO THỰC NGHIỆM ĐA TẦNG (3 - 5 PHÚT)
**Chủ đề:** SQL Injection (OWASP A03:2021) — Môn Lập trình Web nâng cao  
**Giảng viên:** ThS. Lê Hữu Dũng | **Nhóm:** WNC.G01 (Nguyễn Danh Học trình bày)

---

### PHẦN 1: MỞ ĐẦU (30 giây) — Chiếu Slide 1 & 2
- *"Kính thưa thầy và các bạn, hôm nay nhóm WNC.G01 xin trình bày và thực nghiệm trực tiếp lỗ hổng bảo mật **SQL Injection** theo tiêu chuẩn **OWASP Top 10 - A03:2021**."*
- *"Chúng em không đọc lý thuyết dịch sách vở, mà sẽ đi thẳng vào: **Bản chất lỗi**, **Độ nguy hiểm thực tế**, và **Thực nghiệm bắn payload tấn công trực tiếp** trên cả 3 tầng: Giao diện Web, Tầng mạng/API dòng lệnh, và Tầng CSDL SQL Server 2025."*
- *"Nguyên nhân gốc rễ của SQL Injection chỉ nằm ở 1 điểm: **Lập trình viên không tách biệt giữa Dữ liệu (Data) và Câu lệnh (Code)**, mà dùng chuỗi cộng trực tiếp dữ liệu người dùng vào câu SQL."*

---

### PHẦN 2: CHỨNG MINH ĐỘ NGUY HIỂM (30 giây) — Chiếu Slide 3
- *"Lỗ hổng này cực kỳ nguy hiểm vì 3 cấp độ tác động:*
  1. ***Bypass xác thực (Authentication Bypass):** Kẻ tấn công không cần biết mật khẩu vẫn chiếm quyền Quản trị viên (Admin).
  2. ***Rò rỉ toàn bộ CSDL (Data Exfiltration):** Trích xuất mật khẩu, số dư tài khoản, hồ sơ khách hàng.
  3. ***Phá hoại máy chủ:** Kẻ xấu có thể DROP TABLE hoặc gọi lệnh can thiệp hệ điều hành.*"

---

### PHẦN 3: THỰC NGHIỆM DEMO TRỰC TIẾP (2 phút)

#### Bước 3.1: Minh Chứng Trên Web UI (Trình Duyệt)
1. Chỉ vào ô code trên màn hình:
   - *"Ở nhánh Bị Lỗi bên trái, em viết câu lệnh ghép chuỗi thuần: `SELECT * FROM Accounts WHERE Username = '{user}' AND Password = '{pass}'`."*
2. Nhập payload: `' OR '1'='1' --` (hoặc bấm nút màu đỏ trên thanh công cụ). Mật khẩu: gõ bừa.
3. Bấm **Gửi Request Khai Thác**:
   - *"Ngay lập tức hệ thống đăng nhập thành công với quyền Admin, và toàn bộ 4 tài khoản trong CSDL cùng số dư tiền, mật khẩu plain-text đều bị trích xuất hiển thị ra màn hình."*
4. Chuyển sang nhánh Đã Vá bên phải:
   - *"Bây giờ em đưa chính xác payload đó sang nhánh Đã Vá (dùng EF Core LINQ). Kết quả: Hệ thống chặn đứng hoàn toàn, thông báo 'Sai tài khoản hoặc mật khẩu' vì SQL Server đã dùng Parameterized Query (@p0, @p1)."*

#### Bước 3.2: MINH CHỨNG TẦNG SÂU (NẾU THẦY HỎI "CÓ PHẢI WEB TỰ CODE CỨNG FAKE KHÔNG?")
- **Minh chứng qua dòng lệnh cURL (Tầng API):**
  - Mở Terminal, gõ: `./test_exploit_cli.sh`
  - *"Thưa thầy, đây là kịch bản kẻ tấn công không dùng trình duyệt mà dùng lệnh cURL bắn HTTP POST trực tiếp vào server. Server trả về JSON trích xuất sạch sẽ toàn bộ CSDL, chứng minh lỗi nằm ở tầng C# Backend chứ không phải do giao diện Web."*
- **Minh chứng trực tiếp trên Microsoft SQL Server 2025 (Tầng CSDL):**
  - Mở file `verify_sql_server.sql` trên VS Code (kết nối CSDL `OwaspDemoDB`).
  - Chạy câu 1: Ghép chuỗi `SELECT * FROM Accounts WHERE Username = '' OR '1'='1' --' ...` -> SQL Server thực sự trả về 4 dòng.
  - Chạy câu 2: Dùng `EXEC sp_executesql` với tham số `@p0` -> SQL Server trả về 0 dòng.
  - *"Điều này chứng minh triệt để cơ chế phòng thủ ở cấp độ động cơ cơ sở dữ liệu."*

---

### PHẦN 4: KẾT LUẬN & BÀI HỌC CHO BTL (30 giây) — Chiếu Slide 7 & 8
- *"Tóm lại, để phòng thủ triệt để SQL Injection trong ASP.NET Core MVC:*
  1. *Tuyệt đối không ghép chuỗi SQL thủ công. Luôn dùng ORM (EF Core) hoặc Parameterized Query.*
  2. *Áp dụng nguyên tắc đặc quyền tối thiểu (Least Privilege) cho tài khoản kết nối CSDL.*
  3. *Validate dữ liệu đầu vào và luôn mã hóa mật khẩu với ASP.NET Core Identity.*"
- *"Nhóm WNC.G01 xin cảm ơn thầy và các bạn đã lắng nghe!"*
