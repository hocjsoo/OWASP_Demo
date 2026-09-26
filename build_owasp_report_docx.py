import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

doc = docx.Document()

# Page Margins A4
for s in doc.sections:
    s.top_margin = Inches(0.8)
    s.bottom_margin = Inches(0.8)
    s.left_margin = Inches(0.8)
    s.right_margin = Inches(0.8)

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def add_callout(text, title="LƯU Ý QUAN TRỌNG", border_color="1D4ED8", bg_color="EFF6FF"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    set_cell_background(cell, bg_color)
    cell.width = Inches(6.8)
    
    # Border left
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:top w:val="none"/>
            <w:left w:val="single" w:sz="36" w:space="0" w:color="{border_color}"/>
            <w:bottom w:val="none"/>
            <w:right w:val="none"/>
        </w:tcBorders>
    ''')
    tcPr.append(tcBorders)
    
    p = cell.paragraphs[0]
    r_title = p.add_run(f"📌 {title}: ")
    r_title.bold = True
    r_title.font.size = Pt(10.5)
    r_title.font.color.rgb = RGBColor(15, 23, 42)
    
    r_text = p.add_run(text)
    r_text.font.size = Pt(10)
    r_text.font.color.rgb = RGBColor(51, 65, 85)

# -------------------------------------------------------------
# TRANG BÌA
# -------------------------------------------------------------
p_hou = doc.add_paragraph()
p_hou.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p_hou.add_run("TRƯỜNG ĐẠI HỌC MỞ HÀ NỘI\nKHOA CÔNG NGHỆ THÔNG TIN\n―o0o―\n\n")
r.bold = True
r.font.size = Pt(13)
r.font.color.rgb = RGBColor(70, 70, 70)

p_title = doc.add_paragraph()
p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_t = p_title.add_run("BÁO CÁO THỰC NGHIỆM AN TOÀN WEB (OWASP TOP 10)\n")
r_t.bold = True
r_t.font.size = Pt(16)
r_t.font.color.rgb = RGBColor(190, 18, 60)

r_sub = p_title.add_run("CHUYÊN ĐỀ: BỘ 3 LỖ HỔNG KINH ĐIỂN & GIẢI PHÁP PHÒNG THỦ ĐA TẦNG\n(SQL Injection • Stored XSS • Cross-Site Request Forgery)\n\n")
r_sub.bold = True
r_sub.font.size = Pt(14)
r_sub.font.color.rgb = RGBColor(29, 78, 216)

p_info = doc.add_paragraph()
p_info.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_i = p_info.add_run(
    "Học phần: Lập trình Web nâng cao\n"
    "Giảng viên hướng dẫn: ThS. Lê Hữu Dũng\n"
    "Nhóm thực hiện: WNC.G01\n"
    "1. Nguyễn Danh Học  — MSV: 23A1001D0158 (Trưởng nhóm)\n"
    "2. Nguyễn Thanh Bình — MSV: 23A1001D0041\n"
    "3. Nguyễn Minh Cường — MSV: 23A1001D0058\n\n"
    "Hà Nội, Năm 2026\n"
)
r_i.font.size = Pt(11.5)

doc.add_page_break()

# -------------------------------------------------------------
# PHẦN MỞ ĐẦU
# -------------------------------------------------------------
h1 = doc.add_heading("I. GIỚI THIỆU CHUNG & PHÂN CÔNG THỰC NGHIỆM", level=1)
h1.paragraph_format.space_before = Pt(12)

doc.add_paragraph(
    "Trong khuôn khổ học phần Lập trình Web nâng cao (ThS. Lê Hữu Dũng), thực hiện theo chỉ đạo của giảng viên "
    "về việc nghiên cứu thực nghiệm an toàn thông tin theo chuẩn OWASP Top 10 thay vì đọc lý thuyết dịch sách vở, "
    "nhóm WNC.G01 đã xây dựng một nền tảng thực nghiệm độc lập (OWASP_Demo) trên nền ASP.NET Core MVC (.NET 10) "
    "kết nối máy chủ cơ sở dữ liệu Microsoft SQL Server 2025. Nhóm lựa chọn 3 lỗ hổng tiêu biểu nhất tương ứng với 3 thành viên:"
)

# Bảng phân công
table = doc.add_table(rows=4, cols=4)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
headers = ["STT", "Họ và tên & MSV", "Chủ đề OWASP", "Môi trường & Trọng trách"]
for i, h in enumerate(headers):
    cell = table.cell(0, i)
    cell.paragraphs[0].add_run(h).bold = True
    set_cell_background(cell, "E2E8F0")

data = [
    ("1", "Nguyễn Danh Học\n(23A1001D0158)", "SQL Injection\n(OWASP A03:2021)", "Trưởng nhóm: Phân tích tầng CSDL SQL Server 2025, kịch bản bẻ gãy cú pháp, đối chứng cURL API và phòng thủ EF Core LINQ."),
    ("2", "Nguyễn Thanh Bình\n(23A1001D0041)", "Stored XSS\n(OWASP A03:2021)", "Thành viên: Phân tích ngữ cảnh nhận xét đánh giá chuyên gia, kịch bản đánh cắp Cookie Session và phòng thủ Razor HTML Encoding."),
    ("3", "Nguyễn Minh Cường\n(23A1001D0058)", "CSRF Attack\n(OWASP A01:2021)", "Thành viên: Phân tích cơ chế giả mạo yêu cầu từ trang web bẫy của hacker, rút tiền ví nạn nhân và phòng thủ AntiForgeryToken.")
]

for row_idx, row_data in enumerate(data, start=1):
    for col_idx, text in enumerate(row_data):
        cell = table.cell(row_idx, col_idx)
        cell.paragraphs[0].add_run(text)
        set_cell_background(cell, "F8FAFC" if row_idx % 2 == 1 else "FFFFFF")

doc.add_paragraph().paragraph_format.space_after = Pt(10)

# -------------------------------------------------------------
# CHỦ ĐỀ 1: SQL INJECTION (NGUYỄN DANH HỌC)
# -------------------------------------------------------------
doc.add_heading("II. CHỦ ĐỀ 1: THỰC NGHIỆM SQL INJECTION (NGUYỄN DANH HỌC)", level=1)

doc.add_heading("1. Bản chất kỹ thuật & Mức độ nguy hiểm", level=2)
doc.add_paragraph(
    "SQL Injection (A03:2021) là lỗ hổng xảy ra khi dữ liệu đầu vào do người dùng cung cấp bị lập trình viên ghép chuỗi trực tiếp "
    "vào câu lệnh truy vấn SQL mà không qua cơ chế tham số hóa (Parameterized Query). Hệ quản trị CSDL SQL Server khi thông dịch "
    "không thể phân biệt được đâu là dữ liệu của người dùng và đâu là cú pháp lệnh của hệ thống."
)
doc.add_paragraph(
    "Đánh giá rủi ro theo tiêu chuẩn quốc tế CVSS v3.1: Điểm số 9.8 / 10.0 (Mức độ CRITICAL - Cực kỳ nguy hiểm). "
    "Hậu quả: Kẻ tấn công có thể vượt qua bước đăng nhập (Authentication Bypass), trích xuất toàn bộ dữ liệu bảng Accounts "
    "(bao gồm mật khẩu gốc, số dư tài khoản ngân hàng và ghi chú bí mật), hoặc phá hoại dữ liệu bằng lệnh DROP TABLE."
)

doc.add_heading("2. Kịch bản khai thác đối chứng", level=2)
doc.add_paragraph(
    "Trong mã nguồn Controllers/SqlInjectionController.cs, nhóm đã dựng song song 2 Action đối chứng:"
)
doc.add_paragraph(
    "• Đoạn mã dính lỗi (LoginVulnerable):\n"
    "   string rawSql = $\"SELECT * FROM Accounts WHERE Username = '{username}' AND Password = '{password}'\";\n"
    "   var accounts = _context.Accounts.FromSqlRaw(rawSql).ToList();\n\n"
    "• Payload tấn công: username = ' OR '1'='1' --  | password = (bất kỳ)\n"
    "• Câu lệnh thực thi trên SQL Server:\n"
    "   SELECT * FROM Accounts WHERE Username = '' OR '1'='1' --' AND Password = 'xyz'\n"
    "• Cơ chế giải phẫu payload:\n"
    "   - Dấu nháy đơn (') đóng sớm chuỗi ký tự của trường Username.\n"
    "   - Mệnh đề OR '1'='1' biến điều kiện WHERE thành luôn đúng (TRUE) cho mọi dòng trong bảng.\n"
    "   - Ký tự chú thích (--) biến toàn bộ phần kiểm tra mật khẩu phía sau thành chú thích vô hiệu lực."
)

doc.add_heading("3. Cơ chế phòng thủ chuẩn với EF Core LINQ", level=2)
doc.add_paragraph(
    "Nhóm chứng minh cơ chế phòng thủ triệt để tại Action LoginSecure:\n"
    "   var matched = _context.Accounts.Where(a => a.Username == username && a.Password == password).ToList();\n\n"
    "Nguyên lý: EF Core tự động biên dịch câu truy vấn thành Parameterized Query. Khi gửi sang SQL Server, lệnh được gọi qua "
    "thủ tục sp_executesql với tham số @p0 và @p1. SQL Server biên dịch cây thực thi (Execution Plan) trước, sau đó mới gán giá trị "
    "của payload vào tham số @p0 như một chuỗi văn bản (String Literal) thuần túy. Toàn bộ dấu nháy đơn và từ khóa OR/-- mất hoàn toàn "
    "khả năng can thiệp vào logic của CSDL."
)

add_callout(
    "Minh chứng đã được thực nghiệm trên 3 tầng: Giao diện Web (Tab 1), Dòng lệnh Terminal cURL (test_exploit_cli.sh) "
    "và Cơ sở dữ liệu thật Microsoft SQL Server 2025 (Database: OwaspDemoDB, bảng dbo.Accounts).",
    title="MINH CHỨNG THỰC NGHIỆM ĐA TẦNG"
)

doc.add_paragraph().paragraph_format.space_after = Pt(10)

# -------------------------------------------------------------
# CHỦ ĐỀ 2: STORED XSS (NGUYỄN THANH BÌNH)
# -------------------------------------------------------------
doc.add_heading("III. CHỦ ĐỀ 2: THỰC NGHIỆM STORED XSS (NGUYỄN THANH BÌNH)", level=1)

doc.add_heading("1. Ngữ cảnh nghiệp vụ & Bản chất lỗ hổng", level=2)
doc.add_paragraph(
    "Trong đề tài Website cung cấp dịch vụ tư vấn trực tuyến, chức năng gửi đánh giá và nhận xét chuyên gia là tính năng "
    "tương tác trực tiếp của người dùng. Lỗ hổng Stored XSS xảy ra khi ứng dụng lưu trực tiếp chuỗi HTML/JavaScript độc hại "
    "của kẻ tấn công vào cơ sở dữ liệu (bảng dbo.Comments) mà không thực hiện làm sạch (Sanitization). "
    "Khi người dùng khác (hoặc Quản trị viên) vào xem trang đánh giá, trình duyệt sẽ biên dịch và thực thi đoạn script độc hại đó."
)

doc.add_heading("2. Kịch bản tấn công đánh cắp Session Cookie", level=2)
doc.add_paragraph(
    "• Payload tấn công:\n"
    "   <script>alert('Lộ Cookie: ' + document.cookie);</script>\n"
    "   <img src=x onerror=alert('Kích hoạt XSS qua thẻ ảnh!')>\n\n"
    "• Hậu quả thực tế:\n"
    "   Hacker có thể đọc được cookie phiên AuthSessionToken của nạn nhân và gửi ngầm về máy chủ của kẻ tấn công (Cookie Stealing). "
    "   Hacker sau đó chiếm đoạt phiên làm việc (Session Hijacking) mà không cần đánh cắp mật khẩu."
)

doc.add_heading("3. Giải pháp phòng thủ trong ASP.NET Core", level=2)
doc.add_paragraph(
    "• Phòng thủ tầng View: Mặc định Razor View sử dụng cú pháp @comment.Content sẽ tự động HTML Encode các ký tự nhạy cảm "
    "(< biến thành &lt;, > biến thành &gt;). Lỗi chỉ xảy ra khi lập trình viên lạm dụng @Html.Raw().\n"
    "• Phòng thủ tầng Controller: Sử dụng System.Text.Encodings.Web.HtmlEncoder.Default.Encode(content) để mã hóa trước khi lưu.\n"
    "• Phòng thủ tầng Cookie: Thiết lập thuộc tính HttpOnly = true cho cookie phiên, cấm hoàn toàn mã JavaScript truy cập document.cookie."
)

doc.add_paragraph().paragraph_format.space_after = Pt(10)

# -------------------------------------------------------------
# CHỦ ĐỀ 3: CSRF ATTACK (NGUYỄN MINH CƯỜNG)
# -------------------------------------------------------------
doc.add_heading("IV. CHỦ ĐỀ 3: THỰC NGHIỆM CSRF ATTACK (NGUYỄN MINH CƯỜNG)", level=1)

doc.add_heading("1. Bản chất lỗ hổng Cross-Site Request Forgery", level=2)
doc.add_paragraph(
    "CSRF (A01:2021) là kỹ thuật tấn công lừa trình duyệt của nạn nhân (đang có phiên đăng nhập hợp lệ tại hệ thống mục tiêu) "
    "thực hiện một hành động ngoài ý muốn (chuyển tiền ví, đổi mật khẩu) thông qua một trang web giả mạo do hacker kiểm soát."
)

doc.add_heading("2. Kịch bản thực nghiệm rút tiền ví ngầm", level=2)
doc.add_paragraph(
    "• Nạn nhân có số dư ví 50.000.000 VNĐ đang đăng nhập tại hệ thống.\n"
    "• Hacker dựng một trang web trúng thưởng quà tặng (Views/Csrf/AttackerSite.cshtml) chứa một form ẩn:\n"
    "   <form action=\"http://localhost:5076/Csrf/TransferVulnerable\" method=\"post\">\n"
    "       <input type=\"hidden\" name=\"amount\" value=\"20000000\" />\n"
    "   </form>\n"
    "• Khi nạn nhân bấm vào nút nhận thưởng, form ẩn lập tức gửi HTTP POST sang hệ thống mục tiêu. "
    "Do trình duyệt tự động đính kèm Cookie đăng nhập hợp lệ của nạn nhân, Action TransferVulnerable thiếu kiểm tra token "
    "sẽ thực hiện trừ 20.000.000 VNĐ từ ví nạn nhân sang ví hacker."
)

doc.add_heading("3. Cơ chế phòng thủ Anti-Forgery Token", level=2)
doc.add_paragraph(
    "ASP.NET Core cung cấp cơ chế phòng thủ Synchronizer Token Pattern hoàn chỉnh:\n"
    "• Trong Form View: Nhúng cú pháp @Html.AntiForgeryToken() (sinh 1 input ẩn mang token bí mật).\n"
    "• Trong Controller: Bổ sung thuộc tính [ValidateAntiForgeryToken] trước Action nhận POST.\n"
    "• Cơ chế: Máy chủ phát hành 1 token lưu trong cookie và 1 token tương ứng trong form. Do chính sách Same-Origin Policy, "
    "trang web của hacker không thể đọc được token trong form của nạn nhân. Mọi request giả mạo thiếu token đều bị máy chủ từ chối ngay lập tức (Lỗi 400 Bad Request)."
)

doc.add_paragraph().paragraph_format.space_after = Pt(10)

# -------------------------------------------------------------
# PHẦN KẾT LUẬN & LIÊN HỆ ĐỀ TÀI BTL
# -------------------------------------------------------------
doc.add_heading("V. TỔNG KẾT & CAM KẾT BẢO MẬT CHO ĐỀ TÀI BTL", level=1)
doc.add_paragraph(
    "Thông qua việc xây dựng và thực nghiệm trực tiếp 3 chuyên đề an toàn web, nhóm WNC.G01 rút ra bài học sâu sắc: "
    "An ninh thông tin không phải là tính năng bổ sung sau cùng mà phải được thiết kế ngay từ kiến trúc ban đầu (Security by Design - Buổi 8).\n\n"
    "Trong đồ án Website cung cấp dịch vụ tư vấn trực tuyến, nhóm cam kết áp dụng triệt để:\n"
    "1. 100% truy vấn dữ liệu thông qua Entity Framework Core Code-First với LINQ Parameterized (Nguyễn Danh Học).\n"
    "2. 100% dữ liệu đầu vào được kiểm soát chặt chẽ qua Data Annotations và Razor HTML Encoding tự động (Nguyễn Thanh Bình).\n"
    "3. 100% các form giao dịch (đặt lịch tư vấn, thanh toán, đổi mật khẩu) được bảo vệ bằng Anti-Forgery Token và Cookie SameSite (Nguyễn Minh Cường)."
)

output_docx = "/media/hocjsoo/New Volume/OWASP_Demo/Bao_Cao_Thuc_Nghiem_OWASP_WNC_G01.docx"
doc.save(output_docx)
print(f"OWASP Word Report successfully saved to: {output_docx}")
