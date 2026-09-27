import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

doc = docx.Document()

# Thiết lập lề trang A4 chuẩn (Top/Bottom: 2cm, Left: 2.5cm, Right: 2cm)
for section in doc.sections:
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(0.79)
    section.bottom_margin = Inches(0.79)
    section.left_margin = Inches(0.98)
    section.right_margin = Inches(0.79)

# Màu sắc chuẩn nhận diện
COLOR_PRIMARY = RGBColor(29, 78, 216)    # Blue 700 (HOU)
COLOR_DANGER = RGBColor(190, 18, 60)     # Rose 700
COLOR_SUCCESS = RGBColor(4, 120, 87)     # Emerald 700
COLOR_TEXT = RGBColor(15, 23, 42)        # Slate 900
COLOR_MUTED = RGBColor(100, 116, 139)    # Slate 500

def set_cell_shading(cell, color_hex):
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'''
        <w:tcMar {nsdecls("w")}>
            <w:top w:w="{top}" w:type="dxa"/>
            <w:bottom w:w="{bottom}" w:type="dxa"/>
            <w:left w:w="{left}" w:type="dxa"/>
            <w:right w:w="{right}" w:type="dxa"/>
        </w:tcMar>
    ''')
    tcPr.append(tcMar)

def add_callout(doc, title, text, border_color="1D4ED8", bg_color="EFF6FF"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    set_cell_shading(cell, bg_color)
    set_cell_margins(cell, top=140, bottom=140, left=200, right=200)
    
    # Left border only
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:top w:val="none"/>
            <w:left w:val="single" w:sz="24" w:space="0" w:color="{border_color}"/>
            <w:bottom w:val="none"/>
            <w:right w:val="none"/>
        </w:tcBorders>
    ''')
    tcPr.append(tcBorders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    run_t = p.add_run(f"📌 {title}\n")
    run_t.font.name = "Arial"
    run_t.font.size = Pt(10.5)
    run_t.font.bold = True
    if border_color == "1D4ED8":
        run_t.font.color.rgb = COLOR_PRIMARY
    elif border_color == "BE123C":
        run_t.font.color.rgb = COLOR_DANGER
    else:
        run_t.font.color.rgb = COLOR_SUCCESS
        
    run_b = p.add_run(text)
    run_b.font.name = "Arial"
    run_b.font.size = Pt(10)
    run_b.font.color.rgb = COLOR_TEXT

# ==============================================================================
# PHẦN TIÊU ĐỀ BÁO CÁO & THÔNG TIN CHUNG
# ==============================================================================
p_top = doc.add_paragraph()
p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_top.paragraph_format.space_after = Pt(2)
r_top1 = p_top.add_run("TRƯỜNG ĐẠI HỌC MỞ HÀ NỘI — KHOA CÔNG NGHỆ THÔNG TIN\n")
r_top1.font.name = "Arial"
r_top1.font.size = Pt(11)
r_top1.font.bold = True
r_top1.font.color.rgb = COLOR_PRIMARY

r_top2 = p_top.add_run("BỘ MÔN: LẬP TRÌNH WEB NÂNG CAO (ThS. LÊ HỮU DŨNG)")
r_top2.font.name = "Arial"
r_top2.font.size = Pt(10)
r_top2.font.bold = True
r_top2.font.color.rgb = COLOR_MUTED

p_main_title = doc.add_paragraph()
p_main_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_main_title.paragraph_format.space_before = Pt(14)
p_main_title.paragraph_format.space_after = Pt(6)
r_mt = p_main_title.add_run("KỊCH BẢN THUYẾT TRÌNH & THỰC NGHIỆM CHI TIẾT\nCHUYÊN ĐỀ BẢO MẬT WEB (OWASP TOP 10)")
r_mt.font.name = "Arial"
r_mt.font.size = Pt(17)
r_mt.font.bold = True
r_mt.font.color.rgb = COLOR_DANGER

p_sub = doc.add_paragraph()
p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_sub.paragraph_format.space_after = Pt(14)
r_sub = p_sub.add_run("(Tài liệu điều hành bảo vệ đề tài — Đồng bộ 100% với Slide Master 19 trang)\nĐề tài BTL liên hệ: Website Cung Cấp Dịch Vụ Tư Vấn Trực Tuyến • Nhóm WNC.G01")
r_sub.font.name = "Arial"
r_sub.font.size = Pt(10.5)
r_sub.font.italic = True
r_sub.font.color.rgb = COLOR_TEXT

# Bảng phân công 3 thành viên
t_mem = doc.add_table(rows=4, cols=4)
t_mem.alignment = WD_TABLE_ALIGNMENT.CENTER
t_mem.autofit = False

headers = ["Thành viên", "MSV / Vai trò", "Chủ đề phụ trách", "Gắn kết bài toán BTL"]
col_widths = [Inches(1.8), Inches(1.4), Inches(1.8), Inches(2.3)]

for idx, h in enumerate(headers):
    cell = t_mem.cell(0, idx)
    cell.width = col_widths[idx]
    set_cell_shading(cell, "1D4ED8")
    set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(h)
    r.font.name = "Arial"
    r.font.size = Pt(10)
    r.font.bold = True
    r.font.color.rgb = RGBColor(255, 255, 255)

m_data = [
    ("Nguyễn Danh Học", "23A1001D0158\nTrưởng nhóm", "Chủ đề 1: SQL Injection\n(OWASP A03:2021)", "Module Đăng nhập & Xác thực 3 Role (Customer, Specialist, Admin)"),
    ("Nguyễn Thanh Bình", "23A1001D0041\nThành viên", "Chủ đề 2: Stored XSS\n(OWASP A03:2021)", "Module Nhận xét & Đánh giá chất lượng chuyên gia (Reviews/Rating)"),
    ("Nguyễn Minh Cường", "23A1001D0058\nThành viên", "Chủ đề 3: CSRF Attack\n(A01:2021 / Session)", "Module Giao dịch Ví tiền tư vấn & Đặt lịch hẹn qua AJAX Fetch API")
]

for r_idx, row in enumerate(m_data, start=1):
    for c_idx, val in enumerate(row):
        cell = t_mem.cell(r_idx, c_idx)
        cell.width = col_widths[c_idx]
        set_cell_shading(cell, "F8FAFC" if r_idx % 2 == 1 else "FFFFFF")
        set_cell_margins(cell, top=70, bottom=70, left=90, right=90)
        p = cell.paragraphs[0]
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(val)
        r.font.name = "Arial"
        r.font.size = Pt(9.5)
        if c_idx == 0:
            r.font.bold = True
            r.font.color.rgb = COLOR_TEXT
        elif c_idx == 2:
            r.font.bold = True
            r.font.color.rgb = COLOR_PRIMARY

doc.add_paragraph().paragraph_format.space_after = Pt(10)

# ==============================================================================
# HÀM TẠO MỤC SLIDE TRONG WORD
# ==============================================================================
def add_slide_section(doc, slide_no, title, speaker, intent, speech_text, live_action, btl_context):
    h = doc.add_heading(level=2)
    h.paragraph_format.space_before = Pt(14)
    h.paragraph_format.space_after = Pt(4)
    r = h.add_run(f"SLIDE {slide_no:02d}: {title}")
    r.font.name = "Arial"
    r.font.size = Pt(12)
    r.font.bold = True
    r.font.color.rgb = COLOR_PRIMARY

    # Bảng siêu dữ liệu Slide
    tbl = doc.add_table(rows=3, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    
    meta_rows = [
        ("Người trình bày:", speaker),
        ("Mục tiêu truyền tải:", intent),
        ("Gắn kết bài toán BTL:", btl_context)
    ]
    for idx, (lbl, val) in enumerate(meta_rows):
        c1, c2 = tbl.cell(idx, 0), tbl.cell(idx, 1)
        c1.width = Inches(1.8)
        c2.width = Inches(5.5)
        set_cell_shading(c1, "F1F5F9")
        set_cell_shading(c2, "F8FAFC")
        set_cell_margins(c1, 50, 50, 80, 80)
        set_cell_margins(c2, 50, 50, 80, 80)
        
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(lbl)
        r1.font.name = "Arial"
        r1.font.size = Pt(9.5)
        r1.font.bold = True
        r1.font.color.rgb = COLOR_MUTED
        
        p2 = c2.paragraphs[0]
        p2.paragraph_format.space_after = Pt(0)
        r2 = p2.add_run(val)
        r2.font.name = "Arial"
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = COLOR_TEXT

    # Lời thoại chi tiết
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    add_callout(doc, f"Lời nói thuyết trình (Khuyến nghị 30-45 giây):", speech_text, border_color="1D4ED8", bg_color="EFF6FF")
    
    # Thao tác máy chiếu / Live Demo đi kèm
    if live_action:
        doc.add_paragraph().paragraph_format.space_after = Pt(2)
        add_callout(doc, f"Thao tác chỉ dẫn màn hình / Live Demo (Khi thầy yêu cầu):", live_action, border_color="047857", bg_color="ECFDF5")
    
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

# ==============================================================================
# NỘI DUNG 19 SLIDE ĐƯỢC BIÊN SOẠN CÔNG PHU
# ==============================================================================
slides_content = [
    (1, "TRANG BÌA — TỔNG QUAN BÁO CÁO THỰC NGHIỆM AN TOÀN WEB", 
     "Nguyễn Danh Học (Trưởng nhóm)", 
     "Giới thiệu thành viên, học phần, GVHD và đặt vấn đề an toàn ứng dụng web trong kiến trúc 3 tầng.",
     "\"Kính thưa ThS. Lê Hữu Dũng và các bạn! Nhóm WNC.G01 xin phép trình bày báo cáo chuyên đề Thực nghiệm An toàn Web với 3 lỗ hổng tiêu biểu trong OWASP Top 10: SQL Injection, Stored XSS và CSRF. Thay vì tiếp cận lý thuyết thuần túy, nhóm chúng em đặt cả 3 lỗ hổng này vào chính bài toán đề tài BTL: 'Website cung cấp dịch vụ tư vấn trực tuyến'. Nhóm phân định 3 bề mặt nguy cơ trọng yếu tương ứng với 3 thành viên: Em - Nguyễn Danh Học phụ trách tầng CSDL và API với SQL Injection; bạn Nguyễn Thanh Bình phụ trách tầng Client DOM và Cookie với Stored XSS; và bạn Nguyễn Minh Cường phụ trách tầng Giao dịch và Session Token với CSRF.\"",
     "Chỉ tay vào khung bên phải slide 'Mô hình 3 tầng & 3 điểm nguy cơ' để người nghe nắm ngay cấu trúc bài.",
     "Toàn bộ nền tảng hệ thống tư vấn trực tuyến phân chia 3 vai trò (Customer, Specialist, Admin)."),

    (2, "MÔI TRƯỜNG CÔNG NGHỆ & PHẠM VI THỰC NGHIỆM (TECH STACK & LAB SCOPE)", 
     "Nguyễn Danh Học", 
     "Minh bạch môi trường kỹ thuật thực tế và cam kết đạo đức an toàn thông tin (môi trường cô lập).",
     "\"Để phục vụ báo cáo thực nghiệm, nhóm đã dựng một phòng lab độc lập chạy trên máy chủ nội bộ localhost:5076. Backend sử dụng ASP.NET Core MVC trên nền tảng .NET 10 mới nhất, cơ sở dữ liệu kết nối trực tiếp máy chủ Microsoft SQL Server 2025 Developer qua EF Core Code-First. Về phạm vi, chúng em cam kết toàn bộ thực nghiệm chỉ diễn ra trong môi trường nội bộ khép kín, 100% dữ liệu tài khoản và số dư ví là dữ liệu mẫu giả định phục vụ học tập, tuyệt đối không xâm nhập bất kỳ hệ thống thực tế nào bên ngoài.\"",
     "Chỉ rõ 2 cột trên slide: Cột trái Tech Stack (.NET 10, SQL Server 2025), Cột phải Lab Scope (Localhost, Vulnerable vs Secure).",
     "Môi trường này dùng chung đồng bộ với Solution đồ án BTL."),

    (3, "1.1. SQL INJECTION — LỖI XẢY RA Ở ĐÂU & NGUYÊN NHÂN CỐT LÕI", 
     "Nguyễn Danh Học", 
     "Chỉ ra vị trí phát sinh lỗi trong luồng dữ liệu 3 tầng và đoạn code C# dính lỗi.",
     "\"Bắt đầu với chuyên đề 1: SQL Injection. Trong bài toán BTL, lỗ hổng này xuất hiện trực tiếp tại chức năng Đăng nhập hệ thống cho cả 3 vai trò. Luồng đi bình thường: Client gửi thông tin tài khoản lên, Controller nhận tham số và chuyển xuống SQL Server. Nguyên nhân gốc rễ (Root Cause) xảy ra khi lập trình viên sử dụng kỹ thuật ghép chuỗi trực tiếp (String Interpolation) như trên slide: SELECT * FROM Accounts WHERE Username='{username}'. Trình thông dịch CSDL không thể phân định được đâu là Dữ liệu (Data) và đâu là Cú pháp lệnh (Code). Ký tự nháy đơn (') do người dùng nhập vào sẽ đóng sớm chuỗi giá trị và biến phần còn lại thành cấu trúc điều khiển.\"",
     "Nhấn mạnh vào dòng code C# màu đỏ 'FromSqlRaw(rawSql)' và giải thích ký tự nháy đơn (') và chú thích (--).",
     "Form đăng nhập tài khoản hệ thống tư vấn (`/Account/Login`)."),

    (4, "1.2. SQL INJECTION — 3 KỊCH BẢN KHAI THÁC THỰC NGHIỆM", 
     "Nguyễn Danh Học", 
     "Phân tích cơ chế kỹ thuật của 3 trường hợp payload nạp vào form đăng nhập.",
     "\"Nhóm xây dựng 3 kịch bản thực nghiệm điển hình: Kịch bản 1 là Bypass tổng quát với chuỗi nháy đơn OR 1 bằng 1 gạch gạch. Mệnh đề OR 1=1 luôn đúng với mọi bản ghi, còn hai dấu gạch ngang vô hiệu hóa bước kiểm tra mật khẩu phía sau, giúp đăng nhập vào tài khoản đầu tiên trong bảng là Admin. Kịch bản 2 là Tấn công có chủ đích (Targeted Attack) với chuỗi 'admin' gạch gạch', nhắm thẳng vào tài khoản Quản trị viên tối cao mà không cần mật khẩu. Và Kịch bản 3 là Đăng nhập chuẩn hợp lệ dùng để đối chứng, chứng minh hệ thống vẫn xác thực chính xác khi người dùng nhập dữ liệu chuẩn.\"",
     "Chỉ vào 3 thẻ card ngang 01, 02, 03. Đọc rõ từng payload và kết quả thực tế tương ứng.",
     "Cơ chế vượt qua xác thực để chiếm quyền Admin duyệt chuyên gia hoặc quyền Chuyên gia xem bệnh án."),

    (5, "1.3. SQL INJECTION — MINH CHỨNG THỰC NGHIỆM TRỰC QUAN TRÊN WEB UI", 
     "Nguyễn Danh Học", 
     "Trình bày kết quả đối chứng trực quan trên giao diện ứng dụng (Vulnerable vs Secure).",
     "\"Trên màn hình là kết quả thực nghiệm trực tiếp trên web lab. Bên trái là nhánh dính lỗi: Khi chúng em nạp chuỗi ' OR '1'='1' --, máy chủ SQL Server bị bẻ gãy cú pháp, xác thực thành công và làm rò rỉ toàn bộ danh sách 4 tài khoản CSDL gồm cả Admin, Chuyên gia tâm lý và Khách hàng kèm số dư tài khoản ngân hàng và ghi chú mật. Ngược lại, ở bên phải là nhánh đã được phòng thủ bằng EF Core LINQ Parameterized: Khi gửi cùng một chuỗi khai thác, máy chủ coi toàn bộ chuỗi đó là văn bản thuần túy, truy vấn trả về 0 kết quả và từ chối đăng nhập tuyệt đối.\"",
     "Chỉ vào bảng 4 tài khoản bị lộ ở bên trái (màu đỏ) và đối chiếu với thông báo an toàn màu xanh ở bên phải.",
     "Hậu quả rò rỉ thông tin cá nhân và số dư ví của khách hàng trong hệ thống tư vấn."),

    (6, "1.4. SQL INJECTION — MINH CHỨNG GÓI TIN MẠNG (CHROME DEVTOOLS NETWORK)", 
     "Nguyễn Danh Học", 
     "Chứng minh lỗ hổng nằm ở tầng Backend API thông qua công cụ Chrome DevTools (F12).",
     "\"Để chứng minh tính khách quan dưới góc độ kỹ thuật mạng, chúng em bật Chrome DevTools tab Network. Nhìn vào gói tin POST gửi lên /SqlInjection/LoginVulnerable, tham số payload gửi đi nguyên văn chuỗi nháy đơn. Phản hồi Server trả về mã HTTP 200 OK với khối dữ liệu JSON chứa toàn bộ thông tin tài khoản bị trích xuất. Điều này chứng minh lỗ hổng nằm sâu ở tầng xử lý C# Backend chứ không phụ thuộc vào giao diện người dùng, kẻ tấn công hoàn toàn có thể dùng cURL hoặc Postman để rút sạch dữ liệu mà không cần thông qua trình duyệt.\"",
     "Chỉ vào tab Network F12: Dòng request `LoginVulnerable 200 OK`, tab Payload chứa payload và tab Response chứa JSON.",
     "Kiểm thử bảo mật API Endpoint trong bài toán BTL (Buổi 9)."),

    (7, "1.5. SQL INJECTION — CƠ CHẾ PHÒNG THỦ CHUẨN HÓA TRONG BTL", 
     "Nguyễn Danh Học", 
     "Trình bày giải pháp phòng thủ triệt để với EF Core LINQ và phân định rõ lớp bảo vệ mật khẩu.",
     "\"Giải pháp khắc phục chuẩn mực được nhóm áp dụng: Sử dụng LINQ Parameterized Query thông qua EF Core: FirstOrDefault(a => a.Username == user && a.Password == pass). Cơ chế bảo vệ cốt lõi là Data tách biệt hoàn toàn với Cú pháp lệnh. SQL Server sẽ biên dịch Execution Plan trước, sau đó mới truyền giá trị người dùng vào tham số @p0. Kẻ tấn công không thể chèn lệnh được nữa. Một điểm lưu ý học thuật quan trọng: Parameterized Query là giải pháp cho SQL Injection, còn để bảo vệ mật khẩu khi CSDL bị lộ, nhóm áp dụng thêm lớp phòng thủ chiều sâu (Defense-in-Depth) bằng cách băm mật khẩu qua ASP.NET Core Identity (PBKDF2) kèm Salt ngẫu nhiên và cấp quyền tối thiểu (Least Privilege) cho tài khoản kết nối CSDL.\"",
     "Chỉ rõ 2 cột: Cột xanh là code LINQ phòng thủ SQLi, Cột xám bên phải là nguyên tắc tách biệt Hashing mật khẩu.",
     "Áp dụng trực tiếp vào `AppDbContext` và `AccountController` của đồ án BTL."),

    (8, "2.1. STORED XSS — LỖI PHÁT TÁN NHƯ THẾ NÀO & NGUYÊN NHÂN GỐC", 
     "Nguyễn Thanh Bình", 
     "Trình bày bản chất Stored XSS trong nghiệp vụ nhận xét chuyên gia và nguy cơ từ @Html.Raw.",
     "\"Em là Nguyễn Thanh Bình, xin phép trình bày Chủ đề 2: Stored XSS. Trong website tư vấn trực tuyến của nhóm, người dùng có chức năng gửi nhận xét, đánh giá chất lượng chuyên gia sau buổi tư vấn. Luồng phát tán nguy hiểm của Stored XSS như sau: Kẻ tấn công nhập mã độc JavaScript vào ô nhận xét -> Máy chủ lưu nguyên văn vào bảng dbo.Comments trong CSDL. Khi các nạn nhân khác (như chuyên gia hoặc khách hàng khác) vào trang xem phản hồi, trình duyệt của họ tải mã HTML về và tự động thực thi script độc hại đó. Nguyên nhân cốt lõi là ở tầng View, lập trình viên muốn hiển thị chữ in đậm hoặc xuống dòng nên đã lạm dụng cú pháp @Html.Raw(content), vô tình phá bỏ cơ chế mã hóa bảo vệ tự động của Razor Engine.\"",
     "Chỉ vào sơ đồ luồng 4 bước (Attacker -> Web Server -> CSDL -> Nạn nhân) và dòng code lạm dụng @Html.Raw().",
     "Chức năng Đánh giá chuyên gia (`Reviews`) sau phiên tư vấn trực tuyến (P04)."),

    (9, "2.2. STORED XSS — 3 KỊCH BẢN TIÊM MÃ ĐỘC THỰC NGHIỆM", 
     "Nguyễn Thanh Bình", 
     "Phân tích 3 biến thể kịch bản XSS: Trộm cookie, Bypass bộ lọc và Đánh giá hợp lệ.",
     "\"Nhóm xây dựng 3 kịch bản thực nghiệm: Kịch bản 1 là Đánh cắp Cookie phiên làm việc (Cookie Stealing) bằng đoạn mã script alert document.cookie. Khi script chạy, nó có thể đọc chuỗi token phiên và âm thầm gửi về máy chủ của hacker. Kịch bản 2 là Kỹ thuật vượt qua bộ lọc (Filter Bypass) bằng thẻ hình ảnh <img src=x onerror=alert('XSS')>. Nếu hệ thống chỉ chặn từ khóa <script>, hacker dùng thuộc tính onerror khi ảnh lỗi để kích hoạt JavaScript. Kịch bản 3 là Đánh giá hợp lệ thông thường, chứng minh cả 2 nhánh đều hiển thị nội dung trơn tru khi người dùng không chèn mã độc.\"",
     "Đọc rõ 3 thẻ card 01, 02, 03. Giải thích ngắn gọn tại sao thẻ img lại kích hoạt được mã JavaScript.",
     "Bảo vệ phiên đăng nhập của Chuyên gia và Quản trị viên khỏi nguy cơ bị chiếm quyền (Session Hijacking)."),

    (10, "2.3. STORED XSS — MINH CHỨNG THỰC NGHIỆM TRỰC QUAN TRÊN WEB UI", 
     "Nguyễn Thanh Bình", 
     "Trình bày kết quả đối chứng thực tế trên màn hình chức năng nhận xét chuyên gia.",
     "\"Trên giao diện thực nghiệm: Ở nhánh dính lỗi bên trái, ngay khi người dùng mở trang danh sách nhận xét, đoạn mã độc nằm sẵn trong CSDL lập tức kích hoạt hộp thoại Alert đọc trộm chuỗi Session Cookie của người dùng đang đăng nhập. Trong khi đó ở nhánh an toàn bên phải, Razor View tự động nhận diện dữ liệu đầu vào và chuyển đổi các ký tự nguy hiểm thành thực thể an toàn, đoạn mã chỉ hiển thị như văn bản chữ thông thường và tuyệt đối không thể tự kích hoạt.\"",
     "Chỉ vào popup Alert trên ảnh bên trái và đối chiếu với phần text an toàn ở bên phải.",
     "Trang chi tiết chuyên gia hiển thị danh sách nhận xét của khách hàng."),

    (11, "2.4. STORED XSS — MINH CHỨNG CHROME DEVTOOLS (COOKIE & CONSOLE)", 
     "Nguyễn Thanh Bình", 
     "Phân tích cơ chế đánh cắp Cookie qua DevTools và làm rõ cấu hình mô phỏng Lab.",
     "\"Quan sát sâu hơn qua công cụ Chrome DevTools: Tại tab Application mục Cookies, cookie AuthSessionToken của hệ thống bị để cờ HttpOnly bằng false. Đây là cấu hình có chủ đích trong phòng lab để chúng em minh họa nguy cơ: Khi mở tab Console và gõ lệnh document.cookie, toàn bộ chuỗi token phiên làm việc hiện ra rõ ràng. Điều này chứng minh: nếu không có cờ bảo vệ, đoạn mã XSS có thể gửi token này về cho kẻ tấn công để chiếm đoạt phiên làm việc mà không cần biết mật khẩu của nạn nhân.\"",
     "Chỉ vào cột HttpOnly = false màu đỏ trong bảng Cookies và dòng lệnh document.cookie in ra chuỗi token trong Console.",
     "Cơ chế bảo vệ Cookie xác thực của ASP.NET Core Identity trong đồ án BTL."),

    (12, "2.5. STORED XSS — CƠ CHẾ PHÒNG THỦ CHUẨN TRONG ASP.NET CORE", 
     "Nguyễn Thanh Bình", 
     "Nêu giải pháp phòng thủ 2 lớp: Context-Aware Encoding ở View và HttpOnly ở Cookie.",
     "\"Để phòng thủ triệt để Stored XSS, nhóm áp dụng giải pháp 2 lớp: Thứ nhất ở tầng View, tuyệt đối không dùng @Html.Raw(), mà sử dụng cú pháp mặc định @comment.Content hoặc hàm HtmlEncoder.Default.Encode(). Bộ mã hóa ngữ cảnh sẽ tự động chuyển đổi ký tự nhỏ hơn (<) thành &lt; và lớn hơn (>) thành &gt;, triệt tiêu hoàn toàn khả năng thực thi của trình duyệt. Thứ hai ở tầng Quản lý Cookie: Trong môi trường triển khai thực tế của BTL, toàn bộ Cookie xác thực bắt buộc phải bật cờ HttpOnly = true để cấm JavaScript truy cập, Secure = true để bắt buộc truyền qua HTTPS, và bổ sung Header Content-Security-Policy (CSP) ngăn chặn inline script.\"",
     "Chỉ vào cú pháp Razor chuẩn bên trái và đoạn cấu hình CookieOptions (HttpOnly, Secure, SameSite) bên phải.",
     "Chuẩn hóa cấu hình Cookie trong file `Program.cs` và mã hóa Razor trong toàn bộ View BTL."),

    (13, "3.1. CSRF ATTACK — CƠ CHẾ MƯỢN QUYỀN & PHÂN LOẠI CHUẨN OWASP", 
     "Nguyễn Minh Cường", 
     "Phân tích bản chất CSRF, cơ chế mượn quyền người dùng và định vị chuẩn trong OWASP Top 10.",
     "\"Em là Nguyễn Minh Cường, xin phép trình bày Chủ đề 3: Tấn công CSRF (Cross-Site Request Forgery). Về phân loại học thuật chuẩn: CSRF không phải một mục riêng rẽ trong OWASP Top 10 2021; nó liên quan trực tiếp đến A01:2021 – Broken Access Control và cơ chế bảo toàn yêu cầu phiên. Bản chất của CSRF là kỹ thuật tấn công mượn quyền: Nạn nhân đang đăng nhập hợp lệ tại ứng dụng mục tiêu. Sau đó, nạn nhân bị dẫn dụ truy cập một trang web độc hại do hacker kiểm soát. Trang web độc hại này tự động kích hoạt một HTTP POST request ngầm sang ứng dụng mục tiêu. Vì nạn nhân đã đăng nhập, trình duyệt tự động đính kèm Cookie của nạn nhân gửi sang. Điểm mù là máy chủ chỉ xác thực Cookie hợp lệ mà không hề kiểm tra nguồn gốc phát sinh yêu cầu.\"",
     "Chỉ vào sơ đồ luồng 4 bước (Victim logged in -> Mở tab bẫy -> Gửi POST ngầm -> Máy chủ trừ tiền).",
     "Chức năng Giao dịch nạp/rút tiền ví điện tử và Đặt lịch hẹn tư vấn."),

    (14, "3.2. CSRF ATTACK — KỊCH BẢN KHAI THÁC QUA BIỂU MẪU ẨN", 
     "Nguyễn Minh Cường", 
     "Phân tích kịch bản tấn công bằng trang web bẫy trúng thưởng và biểu mẫu ẩn tự động submit.",
     "\"Kịch bản khai thác thực tế: Nạn nhân đang mở tab website tư vấn và nhận được đường link lừa đảo trúng thưởng iPhone ở một tab khác (AttackerSite.cshtml). Trang web bẫy này chứa một biểu mẫu HTML ẩn trỏ trực tiếp đến địa chỉ nhận lệnh chuyển tiền: /Csrf/TransferVulnerable kèm số tiền 20.000.000 VNĐ. Ngay khi nạn nhân bấm nút nhận quà, đoạn JavaScript tự động submit form. Vì action bên Controller không có cơ chế kiểm tra token chống giả mạo, yêu cầu lập tức được máy chủ thực thi hợp lệ và chuyển sạch 20 triệu của nạn nhân sang tài khoản hacker mà nạn nhân không hề hay biết.\"",
     "Chỉ vào khối code HTML Form ẩn bên phải và dòng action trỏ thẳng sang URL dính lỗi.",
     "Các giao dịch nhạy cảm thay đổi số dư tài khoản hoặc hủy lịch tư vấn trái phép."),

    (15, "3.3. CSRF ATTACK — MINH CHỨNG THỰC NGHIỆM TRỰC QUAN TRÊN WEB UI", 
     "Nguyễn Minh Cường", 
     "Trình bày minh chứng số dư ví bị biến động thực tế giữa 2 tab trình duyệt.",
     "\"Minh chứng thực nghiệm trực quan trên lab: Trước khi tấn công, ví nạn nhân có 50.000.000 VNĐ, ví kẻ tấn công có 0 đồng. Khi kích hoạt form ẩn từ trang bẫy của hacker, máy chủ thực thi lệnh POST ngầm. F5 lại trang hệ thống: Số dư ví nạn nhân lập tức bị trừ xuống còn 30.000.000 VNĐ, và ví kẻ tấn công tăng vọt lên 20.000.000 VNĐ. Ngược lại ở nhánh phòng thủ bên phải: Form chuyển tiền chính chủ được tích hợp Token bảo vệ, từ chối 100% mọi yêu cầu xuất phát từ các trang web bên ngoài.\"",
     "Chỉ vào 2 ô thẻ ví tiền ở phần đầu ảnh: Số dư nạn nhân giảm còn 10M-30M và ví hacker tăng lên tương ứng.",
     "Bảo vệ tài chính và số dư ví của khách hàng và chuyên gia trong BTL."),

    (16, "3.4. CSRF ATTACK — MINH CHỨNG GÓI TIN MẠNG (CHROME DEVTOOLS HEADERS)", 
     "Nguyễn Minh Cường", 
     "Chứng minh cơ chế mượn Cookie và thiếu Token qua DevTools Network Inspector.",
     "\"Kiểm tra chi tiết gói tin mạng qua Chrome DevTools: Yêu cầu POST được kích hoạt có mã trạng thái 302 Found. Nhìn vào mục Request Headers: Header Origin xuất phát từ tên miền của kẻ tấn công (http://attacker-site.com), nhưng trình duyệt vẫn tự động đính kèm Cookie phiên làm việc hợp lệ. Điểm yếu mấu chốt được khoanh đỏ bên dưới: Gói tin gửi đi hoàn toàn KHÔNG CÓ trường __RequestVerificationToken. Vì Action xử lý ở Backend thiếu thuộc tính kiểm tra token, máy chủ ngộ nhận đây là thao tác tự nguyện của người dùng và thực thi giao dịch trái phép.\"",
     "Chỉ vào phần Header: Origin kẻ tấn công, Cookie nạn nhân và khung cảnh báo đỏ 'Anti-Forgery Token: ABSENT'.",
     "Bắt buộc kiểm tra tính hợp lệ của nguồn gốc request trong kiến trúc Web API."),

    (17, "3.5. CSRF ATTACK — CƠ CHẾ PHÒNG THỦ BẰNG ANTI-FORGERY TOKEN", 
     "Nguyễn Minh Cường", 
     "Trình bày cơ chế Synchronizer Token Pattern cho cả Form truyền thống và AJAX Fetch API (Buổi 9).",
     "\"Để phòng chống CSRF toàn diện, nhóm triển khai mô hình Synchronizer Token Pattern: Trong Razor View, nhúng cú pháp @Html.AntiForgeryToken() vào trong form. Máy chủ sẽ sinh ra một cặp token bí mật: một lưu trong Cookie và một nhúng vào thẻ ẩn của form. Tại Controller, gắn thuộc tính [ValidateAntiForgeryToken]. Trang web của hacker ở tên miền khác bị chặn đọc token bí mật này bởi chính sách Cùng nguồn gốc (Same-Origin Policy - SOP), do đó không thể giả mạo request được nữa. Đặc biệt bám sát kiến thức Buổi 9, với các thao tác đặt lịch tư vấn dùng AJAX Fetch API không reload trang, chúng em đọc token từ thẻ ẩn và truyền qua HTTP Header RequestVerificationToken, vừa đảm bảo an toàn vừa tối ưu trải nghiệm người dùng.\"",
     "Chỉ rõ 2 phần: Bên trái là @Html.AntiForgeryToken() và [ValidateAntiForgeryToken]; Bên phải là code Fetch API đính kèm token qua Header.",
     "Áp dụng cho toàn bộ các chức năng đặt lịch tư vấn, nạp ví tiền và cập nhật ca làm việc bằng AJAX trong BTL."),

    (18, "TỔNG HỢP: BÀI HỌC KỸ THUẬT & BẢNG ĐỐI CHIẾU 3 LỖ HỔNG", 
     "Nguyễn Danh Học (Trưởng nhóm tổng hợp)", 
     "Tổng kết đối chiếu 3 lỗ hổng thành 3 thẻ Card nổi bật, dễ nhớ, dễ bảo vệ.",
     "\"Kính thưa thầy, để tổng kết lại toàn bộ bài báo cáo thực nghiệm, nhóm chúng em đúc kết thành 3 bài học kỹ thuật cốt lõi: 1. Với SQL Injection (do em phụ trách): Nguyên nhân là do ghép chuỗi SQL trực tiếp -> Cơ chế phòng thủ là dùng Parameterized Query với EF Core LINQ. 2. Với Stored XSS (do bạn Bình phụ trách): Nguyên nhân là do render dữ liệu thô qua @Html.Raw() -> Cơ chế phòng thủ là Context-Aware HTML Encoding kết hợp cờ HttpOnly cho Cookie. 3. Với CSRF Attack (do bạn Cường phụ trách): Nguyên nhân là máy chủ chỉ kiểm tra Cookie mà thiếu xác thực nguồn gốc request -> Cơ chế phòng thủ là triển khai Synchronizer Token Pattern cho cả Form lẫn AJAX Fetch API.\"",
     "Chỉ lần lượt qua 3 Card lớn trên slide: SQL Injection (Xanh) -> Stored XSS (Đỏ) -> CSRF Attack (Xanh lá).",
     "Khung kiến trúc bảo mật toàn diện của nhóm WNC.G01."),

    (19, "KẾT LUẬN THỰC NGHIỆM & PHẦN HỎI ĐÁP (Q&A)", 
     "Nguyễn Danh Học", 
     "Tuyên bố kết luận thực nghiệm, nêu 2 triết lý an ninh cốt lõi và mở đầu phiên hỏi đáp Q&A.",
     "\"Qua quá trình nghiên cứu và thực nghiệm, nhóm WNC.G01 đã đạt được 3 kết quả: Một là tái lập thành công các kịch bản tấn công thực tế trên cả 3 bề mặt (Giao diện Web, Gói tin mạng DevTools và CSDL SQL Server 2025). Hai là kiểm chứng thành công các giải pháp phòng thủ chuẩn hóa. Và ba là đúc kết 2 nguyên tắc an ninh cốt lõi xuyên suốt đồ án BTL: 'Never trust user-controlled input' - Không bao giờ tin tưởng dữ liệu người dùng; và 'Separate data from executable instructions' - Luôn phân tách dữ liệu khỏi cú pháp lệnh thực thi. Toàn bộ giải pháp phòng vệ này đã được nhóm chúng em tích hợp thành Security View trong kiến trúc Đồ án BTL Website Tư Vấn Trực Tuyến. Nhóm WNC.G01 xin trân trọng cảm ơn ThS. Lê Hữu Dũng và các bạn đã lắng nghe. Chúng em rất mong nhận được nhận xét và câu hỏi của thầy ạ!\"",
     "Cả 3 thành viên đứng nghiêm túc, hướng về phía giảng viên và sẵn sàng chuyển sang phần trả lời câu hỏi.",
     "Hoàn tất trọn vẹn chuyên đề an toàn web, chuyển tiếp tự tin sang Đồ án BTL.")
]

for s_data in slides_content:
    s_no, s_title, s_spk, s_int, s_spch, s_act, s_btl = s_data
    add_slide_section(doc, s_no, s_title, s_spk, s_int, s_spch, s_act, s_btl)

# ==============================================================================
# PHẦN PHỤ LỤC: 10 CÂU HỎI PHẢN BIỆN CỦA GIẢNG VIÊN & ĐÁP ÁN MẪU
# ==============================================================================
doc.add_page_break()

h_qa = doc.add_heading(level=1)
h_qa.paragraph_format.space_before = Pt(14)
h_qa.paragraph_format.space_after = Pt(8)
r_qa = h_qa.add_run("PHỤ LỤC: 10 CÂU HỎI VẤN ĐÁP PHẢN BIỆN DỰ KIẾN CỦA GIẢNG VIÊN & ĐÁP ÁN MẪU")
r_qa.font.name = "Arial"
r_qa.font.size = Pt(14)
r_qa.font.bold = True
r_qa.font.color.rgb = COLOR_PRIMARY

qa_list = [
    ("Câu 1 (Học): Tại sao Entity Framework Core LINQ lại chống được SQL Injection?",
     "Thưa thầy, khi sử dụng LINQ (như .Where() hay .FirstOrDefault()), EF Core không ghép chuỗi mà tự động dịch thành câu lệnh SQL có tham số hóa (sp_executesql @p0). SQL Server sẽ phân tích cú pháp và lập Execution Plan trước, sau đó mới truyền giá trị người dùng vào tham số @p0 dưới dạng literal (văn bản thuần). Do đó mọi ký tự như nháy đơn (') hay từ khóa OR đều không thể can thiệp vào logic truy vấn."),

    ("Câu 2 (Học): Nếu tôi bắt buộc phải dùng câu lệnh SQL thuần (Raw SQL) trong EF Core thì làm sao để an toàn?",
     "Thưa thầy, chúng ta tuyệt đối không dùng FromSqlRaw với phép cộng chuỗi ($). Thay vào đó, EF Core cung cấp hàm FromSqlInterpolated() hoặc FromSqlRaw() kèm mảng tham số SqlParameter. Với FromSqlInterpolated, EF Core tự động nhận diện các biến truyền vào trong chuỗi $ và bọc chúng thành tham số @p0, đảm bảo an toàn 100% như dùng LINQ."),

    ("Câu 3 (Học): Password Hashing có ngăn chặn được SQL Injection không?",
     "Thưa thầy, hoàn toàn không ạ. Password Hashing là cơ chế bảo vệ mật khẩu (Defense-in-Depth) phòng trường hợp CSDL bị đánh cắp hoặc trích xuất ra ngoài thì hacker không đọc được mật khẩu gốc. Còn lỗi SQL Injection xảy ra ở mệnh đề logic WHERE, hacker bypass trực tiếp mà không cần khớp mật khẩu, nên để chặn SQLi bắt buộc phải dùng Parameterized Query."),

    ("Câu 4 (Bình): Phân biệt Stored XSS với Reflected XSS và DOM-based XSS?",
     "Thưa thầy: Reflected XSS là mã độc nằm trong URL (query string) và chỉ chạy khi nạn nhân bấm vào link đó; DOM-based XSS là lỗi hoàn toàn ở mã JavaScript phía client tự đọc và ghi dữ liệu không an toàn vào DOM; còn Stored XSS là mã độc được lưu trữ vĩnh viễn trong CSDL của máy chủ, bất kỳ ai truy cập trang hiển thị dữ liệu đó đều tự động bị tấn công mà không cần bấm link lạ."),

    ("Câu 5 (Bình): Nếu tôi viết hàm kiểm tra, cứ thấy chữ '<script>' là xóa đi thì có chặn được XSS không?",
     "Thưa thầy là không an toàn ạ. Đây là cách tiếp cận Blacklist (danh sách đen) rất dễ bị bypass. Kẻ tấn công có thể chèn script qua các thẻ HTML khác có hỗ trợ Event Handler như <img src=x onerror=alert(1)>, <svg onload=...>, hoặc dùng kỹ thuật lồng thẻ như <scr<script>ipt>. Cách phòng thủ chuẩn nhất là Whitelist Sanitization hoặc Context-Aware HTML Encoding ở đầu ra."),

    ("Câu 6 (Bình): Cờ HttpOnly trong Cookie hoạt động thế nào và tại sao nó quan trọng?",
     "Thưa thầy, cờ HttpOnly là một chỉ thị yêu cầu trình duyệt cấm tuyệt đối mọi đoạn mã JavaScript (như document.cookie) truy cập vào cookie đó. Cookie chỉ được đính kèm tự động trong các HTTP Request gửi về máy chủ. Kể cả khi trang web xuất hiện lỗi XSS, kẻ tấn công cũng không thể đọc trộm được Session Cookie để chiếm đoạt phiên làm việc."),

    ("Câu 7 (Cường): CSRF khác gì so với XSS?",
     "Thưa thầy: XSS là kẻ tấn công tiêm mã kịch bản độc hại để chạy TRỰC TIẾP trên trình duyệt của nạn nhân tại trang web mục tiêu; còn CSRF là kẻ tấn công tạo một trang web KHÁC (Cross-Site) để lừa trình duyệt của nạn nhân gửi request mượn quyền (kèm Cookie hợp lệ) sang trang web mục tiêu mà không cần chạy mã trên trang web mục tiêu."),

    ("Câu 8 (Cường): Tại sao trang web của hacker lại không đọc được Anti-Forgery Token của nạn nhân?",
     "Thưa thầy, đó là nhờ Chính sách Cùng nguồn gốc (Same-Origin Policy - SOP) của trình duyệt. Trình duyệt ngăn cấm mã JavaScript ở tên miền A (như attacker-site.com) đọc nội dung DOM hoặc tài nguyên của tên miền B (như website tư vấn localhost:5076). Hacker chỉ có thể kích hoạt gửi request POST sang chứ không thể đọc trộm mã token ẩn trong form của nạn nhân."),

    ("Câu 9 (Cường): Thuộc tính SameSite của Cookie có thay thế hoàn toàn được Anti-Forgery Token không?",
     "Thưa thầy, SameSite=Strict hoặc Lax là một lớp phòng thủ rất tốt ở tầng trình duyệt để hạn chế gửi cookie cross-site. Tuy nhiên, nó không thể thay thế hoàn toàn Token vì vẫn có rủi ro từ các trình duyệt cũ không hỗ trợ SameSite, hoặc các request dạng top-level navigation ở chế độ Lax. Do đó, khuyến nghị bảo mật của OWASP là kết hợp cả hai: SameSite Cookie và Anti-Forgery Token."),

    ("Câu 10 (Cả nhóm): Nhóm đã đưa các giải pháp an ninh này vào Đồ án BTL như thế nào?",
     "Thưa thầy, toàn bộ các giải pháp này đã được tích hợp thành Security View trong đồ án: Toàn bộ truy vấn CSDL trong BTL dùng 100% LINQ qua EF Core Code-First; Form đăng nhập và đánh giá chuyên gia dùng Razor HTML Encoding mặc định và mã hóa Identity; còn toàn bộ Form chuyển tiền ví và gọi AJAX Fetch API đặt lịch hẹn đều bắt buộc xác thực qua Anti-Forgery Token.")
]

for q, a in qa_list:
    p_q = doc.add_paragraph()
    p_q.paragraph_format.space_before = Pt(8)
    p_q.paragraph_format.space_after = Pt(2)
    r_q = p_q.add_run(f"❓ {q}")
    r_q.font.name = "Arial"
    r_q.font.size = Pt(10.5)
    r_q.font.bold = True
    r_q.font.color.rgb = COLOR_PRIMARY
    
    p_a = doc.add_paragraph()
    p_a.paragraph_format.space_before = Pt(0)
    p_a.paragraph_format.space_after = Pt(6)
    r_a = p_a.add_run(f"➔ Trả lời: {a}")
    r_a.font.name = "Arial"
    r_a.font.size = Pt(10)
    r_a.font.color.rgb = COLOR_TEXT

out_docx_path = "/media/hocjsoo/New Volume/OWASP_Demo/Kich_Ban_Thuyet_Trinh_OWASP_WNC_G01.docx"
doc.save(out_docx_path)
print(f"Presentation Speech Word Document successfully created at: {out_docx_path}")
