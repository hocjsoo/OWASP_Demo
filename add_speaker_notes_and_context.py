import os
from pptx import Presentation

pptx_path = '/media/hocjsoo/New Volume/OWASP_Demo/Slide_OWASP_Top3_Nhom_WNC_G01.pptx'
prs = Presentation(pptx_path)

# Speaker Notes mapping tailored for WNC.G01 Online Consultation Website BTL
notes = {
    1: """[PHẦN MỞ ĐẦU - TỔNG QUAN ĐỀ TÀI]
- Nhóm WNC.G01 xin chào thầy và các bạn.
- Báo cáo thực nghiệm chuyên đề an toàn ứng dụng Web (OWASP Top 10) gắn liền trực tiếp với bài toán đề tài BTL: "Website cung cấp dịch vụ tư vấn trực tuyến".
- Hệ thống phân định 3 bề mặt nguy cơ trọng yếu:
  1. Tầng CSDL (Database Engine): Nguy cơ SQL Injection khi đăng nhập/truy vấn. Phụ trách: Nguyễn Danh Học.
  2. Tầng Trình duyệt (Client DOM & Cookie): Nguy cơ Stored XSS trong chức năng nhận xét/đánh giá chuyên gia. Phụ trách: Nguyễn Thanh Bình.
  3. Tầng Ứng dụng & Giao dịch (Web Server & Session): Nguy cơ CSRF trong chức năng chuyển tiền ví & đặt lịch tư vấn trực tuyến. Phụ trách: Nguyễn Minh Cường.""",

    2: """[BỐI CẢNH & MÔI TRƯỜNG THỰC NGHIỆM]
- Môi trường kỹ thuật: Linux Ubuntu 24.04 LTS, máy chủ CSDL Microsoft SQL Server 2025 Developer kết nối qua EF Core 10.0 Code-First, giao diện Razor View Engine.
- Lab Scope: Thực nghiệm hoàn toàn khép kín trên ứng dụng do nhóm tự xây dựng (http://localhost:5076).
- Mọi dữ liệu tài khoản, số dư ví và bình luận đều là dữ liệu giả lập có kiểm soát nhằm chứng minh tính rủi ro và hiệu quả của các biện pháp phòng vệ.""",

    3: """[CHỦ ĐỀ 1 - SQL INJECTION: BẢN CHẤT LỖ HỔNG (HỌC)]
- Lỗi xảy ra ở đâu trong BTL: Chức năng Đăng nhập hệ thống (cho cả 3 Role: Customer, Specialist, Admin).
- Luồng dữ liệu: Input từ trình duyệt -> Controller nhận tham số -> Ghép chuỗi SQL thô bằng string interpolation -> SQL Server 2025 thực thi.
- Root Cause: Dữ liệu người dùng không được phân tách ranh giới với cú pháp lệnh (Data vs Code). Ký tự nháy đơn (') đóng chuỗi sớm, biến vế mật khẩu thành chú thích (--).""",

    4: """[CHỦ ĐỀ 1 - SQL INJECTION: 3 KỊCH BẢN KHAI THÁC (HỌC)]
- Kịch bản 1 (' OR '1'='1' --): Bypass tổng quát, mệnh đề OR 1=1 luôn đúng với mọi bản ghi, đăng nhập vào tài khoản đầu tiên (Admin).
- Kịch bản 2 (admin' --): Tấn công có chủ đích (Targeted Attack), chỉ định đích danh tài khoản Admin, bỏ qua hoàn toàn bước kiểm tra mật khẩu.
- Kịch bản 3 (admin / AdminPassword@2026): Kịch bản chuẩn đối chứng, chứng minh cả 2 nhánh (Lỗi và Phòng thủ) đều xác thực đúng khi người dùng nhập dữ liệu chuẩn.""",

    5: """[CHỦ ĐỀ 1 - SQL INJECTION: MINH CHỨNG WEB UI (HỌC)]
- Trực quan trên giao diện:
  + Nhánh dính lỗi (Trái): Nhập payload ' OR '1'='1' --, máy chủ xác thực thành công, làm rò rỉ toàn bộ 4 tài khoản gồm Admin, Chuyên gia và Khách hàng kèm số dư ví và thông tin nhạy cảm.
  + Nhánh phòng thủ (Phải): Hệ thống áp dụng EF Core LINQ, từ chối truy cập vì coi toàn bộ payload là chuỗi văn bản thuần (0 kết quả).""",

    6: """[CHỦ ĐỀ 1 - SQL INJECTION: MINH CHỨNG DEVTOOLS NETWORK (HỌC)]
- Soi gói tin mạng thực tế:
  + Request POST gửi lên /SqlInjection/LoginVulnerable mang payload username=' OR '1'='1' --.
  + Phản hồi JSON từ Server: Trả về HTTP 200 OK với danh sách 4 tài khoản CSDL. Điều này chứng minh lỗ hổng nằm ở tầng C# Backend chứ không phụ thuộc vào giao diện.""",

    7: """[CHỦ ĐỀ 1 - SQL INJECTION: PHÒNG THỦ CHUẨN TRONG BTL (HỌC)]
- Giải pháp cốt lõi: Áp dụng Parameterized Query với EF Core LINQ (FirstOrDefault(a => a.Username == user && a.Password == pass)). SQL Server biên dịch Execution Plan trước khi gán tham số @p0, @p1.
- Bảo vệ mật khẩu (Defense-in-Depth): Tách biệt rõ ràng với SQLi. Dùng ASP.NET Core Identity băm mật khẩu 1 chiều (PBKDF2) kèm Salt ngẫu nhiên để bảo vệ người dùng kể cả khi CSDL bị lộ.
- Áp dụng quyền tối thiểu (Least Privilege) cho tài khoản kết nối CSDL.""",

    8: """[CHỦ ĐỀ 2 - STORED XSS: BẢN CHẤT LỖ HỔNG (BÌNH)]
- Lỗi xảy ra ở đâu trong BTL: Chức năng Đánh giá / Nhận xét chất lượng chuyên gia sau buổi tư vấn.
- Luồng phát tán: Kẻ tấn công tiêm script -> Server lưu trực tiếp vào bảng dbo.Comments -> Nạn nhân (khách hàng khác hoặc chuyên gia) truy cập trang xem nhận xét -> Trình duyệt nạn nhân tự động thực thi script.
- Root Cause: Sử dụng @Html.Raw(content) vô hiệu hóa cơ chế phòng thủ mã hóa tự động của Razor View.""",

    9: """[CHỦ ĐỀ 2 - STORED XSS: 3 KỊCH BẢN KHAI THÁC (BÌNH)]
- Kịch bản 1 (<script>alert(document.cookie)</script>): Đánh cắp Session Cookie (Cookie Stealing) để chiếm đoạt phiên làm việc của chuyên gia/quản trị viên.
- Kịch bản 2 (<img src=x onerror=...>): Vượt qua bộ lọc từ khóa đơn giản (Filter Bypass) bằng sự kiện HTML khi đường dẫn ảnh lỗi.
- Kịch bản 3: Bình luận văn bản hợp lệ đối chứng, hiển thị an toàn trên cả 2 nhánh.""",

    10: """[CHỦ ĐỀ 2 - STORED XSS: MINH CHỨNG WEB UI (BÌNH)]
- Trực quan trên giao diện:
  + Nhánh dính lỗi (Trái): Khi tải trang, đoạn script nằm sẵn trong CSDL lập tức kích hoạt hộp thoại Alert hiển thị Cookie AuthSessionToken.
  + Nhánh phòng thủ (Phải): Razor Engine tự động mã hóa chuỗi thành văn bản an toàn (&lt;script&gt;), triệt tiêu hoàn toàn khả năng thực thi mã.""",

    11: """[CHỦ ĐỀ 2 - STORED XSS: MINH CHỨNG DEVTOOLS COOKIE & CONSOLE (BÌNH)]
- Soi bộ nhớ trình duyệt & Console:
  + Cửa sổ Application > Cookies: Cookie AuthSessionToken bị để cờ HttpOnly = false (cấu hình mô phỏng lab có chủ đích để minh chứng lỗ hổng).
  + Cửa sổ Console: Lệnh document.cookie đọc trọn vẹn chuỗi token nhạy cảm. Đây là cách hacker trích xuất phiên và gửi ngầm về server điều khiển.""",

    12: """[CHỦ ĐỀ 2 - STORED XSS: PHÒNG THỦ CHUẨN TRONG BTL (BÌNH)]
- Phòng thủ tầng View: Luôn sử dụng cú pháp Razor mặc định @comment.Content (tự động HTML Encode biến đổi <, > thành &lt;, &gt;).
- Phòng thủ tầng Cookie: Trong sản phẩm thực tế, bắt buộc cấu hình HttpOnly = true, Secure = true, SameSite = Strict để cấm JavaScript đọc Cookie.
- Bổ sung Content Security Policy (CSP Header) với directive script-src 'self'.""",

    13: """[CHỦ ĐỀ 3 - CSRF: BẢN CHẤT LỖ HỔNG (CƯỜNG)]
- Lỗi xảy ra ở đâu trong BTL: Chức năng Giao dịch Ví tiền tư vấn & Đặt lịch hẹn trực tuyến.
- Luồng tấn công: Nạn nhân có phiên đăng nhập hợp lệ -> Bị dẫn dụ mở trang web bẫy của hacker ở tab bên cạnh -> Trang bẫy gửi POST ngầm sang hệ thống -> Trình duyệt tự động đính kèm Cookie của nạn nhân -> Hệ thống thực hiện lệnh chuyển tiền.
- Phân loại chuẩn OWASP: CSRF liên quan trực tiếp đến A01:2021 – Broken Access Control và cơ chế bảo toàn yêu cầu phiên.""",

    14: """[CHỦ ĐỀ 3 - CSRF: KỊCH BẢN KHAI THÁC FORM ẨN (CƯỜNG)]
- Kịch bản lừa đảo: Nạn nhân được dẫn dụ vào trang bẫy trúng thưởng (AttackerSite.cshtml).
- Kỹ thuật khai thác: Trang bẫy chứa biểu mẫu ẩn <form action="/Csrf/TransferVulnerable" method="POST"> kèm giá trị amount=20.000.000. Đoạn script tự động submit form mà nạn nhân không hề hay biết.
- Hậu quả: Tiền trong ví nạn nhân bị rút sạch sang tài khoản hacker.""",

    15: """[CHỦ ĐỀ 3 - CSRF: MINH CHỨNG WEB UI (CƯỜNG)]
- Trực quan trên giao diện:
  + Nhánh dính lỗi: Mở Attacker Site và nhấn nhận quà -> Quay lại trang hệ thống thấy số dư ví nạn nhân bị trừ từ 50.000.000 VNĐ xuống 30.000.000 VNĐ, ví hacker tăng lên 20.000.000 VNĐ.
  + Nhánh phòng thủ: Form chuyển tiền chính chủ có tích hợp Token bảo vệ -> Chặn đứng mọi yêu cầu giả mạo từ trang web khác.""",

    16: """[CHỦ ĐỀ 3 - CSRF: MINH CHỨNG DEVTOOLS HEADERS (CƯỜNG)]
- Soi gói tin HTTP POST gửi ngầm:
  + Request Headers: Header Origin là tên miền hacker (http://attacker-site.com), nhưng Cookie xác thực phiên vẫn được trình duyệt tự động đính kèm.
  + Điểm mù mấu chốt: Form hoàn toàn KHÔNG CÓ trường __RequestVerificationToken. Vì Action thiếu [ValidateAntiForgeryToken], máy chủ chỉ kiểm tra Cookie hợp lệ rồi trừ tiền ngay mà không xác thực nguồn gốc request.""",

    17: """[CHỦ ĐỀ 3 - CSRF: PHÒNG THỦ CHUẨN TRONG BTL (CƯỜNG)]
- Phòng thủ Form truyền thống: Áp dụng Synchronizer Token Pattern bằng @Html.AntiForgeryToken() trong View và kiểm tra [ValidateAntiForgeryToken] trên Controller. Máy chủ phát hành cặp token ngẫu nhiên, trang web ngoài không thể đọc trộm do chính sách Same-Origin Policy (SOP).
- Phòng thủ AJAX Fetch API (Slide Buổi 9): Đọc token từ trường ẩn và đính kèm vào HTTP Header RequestVerificationToken cho mọi thao tác đặt lịch/thanh toán không reload trang.""",

    18: """[TỔNG HỢP: BÀI HỌC KỸ THUẬT & MA TRẬN ĐỐI CHIẾU]
- Bảng tổng hợp đối chiếu 3 chiều cho cả 3 thành viên:
  1. SQL Injection (Học): Root Cause = Ghép chuỗi SQL -> Phòng thủ = Parameterized Query (EF Core LINQ).
  2. Stored XSS (Bình): Root Cause = Render script thô qua @Html.Raw() -> Phòng thủ = Context-Aware HTML Encoding + HttpOnly Cookie.
  3. CSRF (Cường): Root Cause = Thiếu xác thực nguồn gốc request -> Phòng thủ = Synchronizer Token Pattern (@Html.AntiForgeryToken / Header).""",

    19: """[KẾT LUẬN & CAM KẾT CHUẨN BẢO MẬT BTL]
- Kết luận thực nghiệm: Tái lập thành công lỗ hổng trên cả 3 bề mặt (UI, DevTools, Database) và kiểm chứng giải pháp phòng thủ triệt để.
- Triết lý an ninh cốt lõi:
  1. Never trust user-controlled input (Không bao giờ tin tưởng dữ liệu người dùng).
  2. Separate data from executable instructions (Luôn phân tách dữ liệu khỏi cú pháp lệnh).
- Cam kết áp dụng: Toàn bộ các cơ chế phòng thủ đã được tích hợp thành Security View trong kiến trúc Đồ án BTL Website Tư Vấn Trực Tuyến của nhóm WNC.G01.
- Nhóm xin trân trọng cảm ơn thầy và sẵn sàng bước vào phần Q&A!"""
}

for idx, slide in enumerate(prs.slides):
    slide_num = idx + 1
    if slide_num in notes:
        slide.notes_slide.notes_text_frame.text = notes[slide_num]

prs.save(pptx_path)
print("Updated all 19 slides with rich speaker notes and BTL context!")
