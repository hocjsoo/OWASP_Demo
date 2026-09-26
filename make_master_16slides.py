import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# Bảng màu Gamma / Canva Light Theme (Chuẩn máy chiếu giảng đường)
BG_WHITE = RGBColor(255, 255, 255)
WHITE = RGBColor(255, 255, 255)
CARD_BG = RGBColor(248, 250, 252)          # Slate 50
CARD_BORDER = RGBColor(226, 232, 240)      # Slate 200

TEXT_HEAD = RGBColor(15, 23, 42)           # Slate 900 (Đen sẫm sắc nét)
TEXT_BODY = RGBColor(51, 65, 85)           # Slate 700 (Dễ đọc từ xa)
TEXT_MUTED = RGBColor(100, 116, 139)       # Slate 500

ACCENT_BLUE = RGBColor(29, 78, 216)        # Blue 700 (HOU)
BLUE_LIGHT = RGBColor(239, 246, 255)       # Blue 50
BLUE_BORDER = RGBColor(191, 219, 254)

ACCENT_RED = RGBColor(190, 18, 60)         # Rose 700 (Danger)
RED_LIGHT = RGBColor(255, 241, 242)        # Rose 50
RED_BORDER = RGBColor(254, 205, 211)

ACCENT_GREEN = RGBColor(4, 120, 87)        # Emerald 700 (Secure)
GREEN_LIGHT = RGBColor(236, 253, 245)      # Emerald 50
GREEN_BORDER = RGBColor(167, 243, 208)

ACCENT_AMBER = RGBColor(180, 83, 9)        # Amber 700
AMBER_LIGHT = RGBColor(255, 251, 235)      # Amber 50
AMBER_BORDER = RGBColor(254, 240, 138)

blank_layout = prs.slide_layouts[6]

def add_header(slide, title_text, category_tag="OWASP TOP 10 • THỰC NGHIỆM AN TOÀN WEB"):
    top_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.08))
    top_line.fill.solid()
    top_line.fill.fore_color.rgb = ACCENT_BLUE
    top_line.line.fill.background()

    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.32), Inches(4.8), Inches(0.38))
    badge.fill.solid()
    badge.fill.fore_color.rgb = BLUE_LIGHT
    badge.line.color.rgb = BLUE_BORDER

    tf_b = badge.text_frame
    tf_b.word_wrap = False
    p_b = tf_b.paragraphs[0]
    p_b.text = category_tag.upper()
    p_b.font.size = Pt(10)
    p_b.font.bold = True
    p_b.font.color.rgb = ACCENT_BLUE

    tx = slide.shapes.add_textbox(Inches(0.8), Inches(0.72), Inches(11.7), Inches(0.6))
    tf = tx.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = TEXT_HEAD

    footer = slide.shapes.add_textbox(Inches(0.8), Inches(7.1), Inches(11.7), Inches(0.3))
    ft = footer.text_frame
    p_ft = ft.paragraphs[0]
    p_ft.text = "Nhóm WNC.G01 (Học - Bình - Cường) • Khoa CNTT - Đại học Mở Hà Nội • GVHD: ThS. Lê Hữu Dũng"
    p_ft.font.size = Pt(9.5)
    p_ft.font.color.rgb = TEXT_MUTED


# ==============================================================================
# SLIDE 1: BÌA BÁO CÁO
# ==============================================================================
s1 = prs.slides.add_slide(blank_layout)
bar1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.12))
bar1.fill.solid()
bar1.fill.fore_color.rgb = ACCENT_BLUE
bar1.line.fill.background()

u_badge = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(0.7), Inches(5.8), Inches(0.42))
u_badge.fill.solid()
u_badge.fill.fore_color.rgb = BLUE_LIGHT
u_badge.line.color.rgb = BLUE_BORDER
p = u_badge.text_frame.paragraphs[0]
p.text = "TRƯỜNG ĐẠI HỌC MỞ HÀ NỘI — KHOA CÔNG NGHỆ THÔNG TIN"
p.font.size = Pt(10.5)
p.font.bold = True
p.font.color.rgb = ACCENT_BLUE

tb1 = s1.shapes.add_textbox(Inches(1.0), Inches(1.25), Inches(11.333), Inches(3.2))
tf1 = tb1.text_frame
tf1.word_wrap = True
p = tf1.paragraphs[0]
p.text = "BÁO CÁO THỰC NGHIỆM AN TOÀN WEB (OWASP TOP 10)"
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = ACCENT_RED

p = tf1.add_paragraph()
p.text = "Bộ 3 Lỗ Hổng Kinh Điển & Cơ Chế Phòng Thủ\n(SQL Injection • Stored XSS • CSRF Attack)"
p.font.size = Pt(30)
p.font.bold = True
p.font.color.rgb = TEXT_HEAD
p.space_before = Pt(8)

p = tf1.add_paragraph()
p.text = "Môn học: Lập trình Web nâng cao  •  Giảng viên hướng dẫn: ThS. Lê Hữu Dũng"
p.font.size = Pt(14)
p.font.color.rgb = TEXT_BODY
p.space_before = Pt(8)

col_w = Inches(3.64)
gap = Inches(0.2)

m1 = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(4.7), col_w, Inches(2.0))
m1.fill.solid()
m1.fill.fore_color.rgb = CARD_BG
m1.line.color.rgb = BLUE_BORDER
tf = m1.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "CHỦ ĐỀ 1: SQL INJECTION"
p.font.size = Pt(11)
p.font.bold = True
p.font.color.rgb = ACCENT_BLUE
p = tf.add_paragraph()
p.text = "Nguyễn Danh Học\nMSV: 23A1001D0158\nTrưởng nhóm • Demo CSDL & API"
p.font.size = Pt(12)
p.font.color.rgb = TEXT_HEAD
p.space_before = Pt(4)

m2 = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0) + col_w + gap, Inches(4.7), col_w, Inches(2.0))
m2.fill.solid()
m2.fill.fore_color.rgb = CARD_BG
m2.line.color.rgb = RED_BORDER
tf = m2.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "CHỦ ĐỀ 2: STORED XSS"
p.font.size = Pt(11)
p.font.bold = True
p.font.color.rgb = ACCENT_RED
p = tf.add_paragraph()
p.text = "Nguyễn Thanh Bình\nMSV: 23A1001D0041\nThành viên • Demo Cookie & DOM"
p.font.size = Pt(12)
p.font.color.rgb = TEXT_HEAD
p.space_before = Pt(4)

m3 = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0) + (col_w + gap)*2, Inches(4.7), col_w, Inches(2.0))
m3.fill.solid()
m3.fill.fore_color.rgb = CARD_BG
m3.line.color.rgb = GREEN_BORDER
tf = m3.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "CHỦ ĐỀ 3: CSRF ATTACK"
p.font.size = Pt(11)
p.font.bold = True
p.font.color.rgb = ACCENT_GREEN
p = tf.add_paragraph()
p.text = "Nguyễn Minh Cường\nMSV: 23A1001D0058\nThành viên • Demo Form ẩn giả mạo"
p.font.size = Pt(12)
p.font.color.rgb = TEXT_HEAD
p.space_before = Pt(4)


# ==============================================================================
# SLIDE 2: MỤC TIÊU & PHƯƠNG CHÂM THỰC NGHIỆM
# ==============================================================================
s2 = prs.slides.add_slide(blank_layout)
add_header(s2, "Mục Tiêu & Phương Châm Báo Cáo Thực Nghiệm OWASP")

c_t1 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.4), Inches(5.65), Inches(5.3))
c_t1.fill.solid()
c_t1.fill.fore_color.rgb = CARD_BG
c_t1.line.color.rgb = CARD_BORDER

tf = c_t1.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "🎯 Yêu Cầu Của Giảng Viên (ThS. Lê Hữu Dũng)"
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = ACCENT_RED
p = tf.add_paragraph()
p.text = "• Không dịch lý thuyết tiếng Anh sách vở.\n• Thấy được lỗ hổng thực sự nguy hiểm đến mức nào trong môi trường thực tế.\n• Bắt buộc phải thực nghiệm demo: Muốn phòng thủ an toàn, phải hiểu hacker tấn công hệ thống theo cách nào.\n• Báo cáo đối chứng: Minh chứng rõ ràng giữa code dính lỗi và code đã vá chuẩn."
p.font.size = Pt(13)
p.font.color.rgb = TEXT_BODY
p.space_before = Pt(10)

c_t2 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.85), Inches(1.4), Inches(5.65), Inches(5.3))
c_t2.fill.solid()
c_t2.fill.fore_color.rgb = BLUE_LIGHT
c_t2.line.color.rgb = BLUE_BORDER

tf = c_t2.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "🚀 Phương Pháp Thực Hiện Của Nhóm WNC.G01"
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = ACCENT_BLUE
p = tf.add_paragraph()
p.text = "• Xây dựng ứng dụng thực nghiệm độc lập (OWASP_Demo) trên ASP.NET Core MVC .NET 10.\n• Kết nối trực tiếp máy chủ CSDL Microsoft SQL Server 2025 Developer trên Localhost.\n• Tạo sẵn 3 phân hệ tương tác riêng biệt cho 3 thành viên: Bấm 1 click nạp payload và quan sát kết quả tức thì.\n• Áp dụng kết quả thực nghiệm vào thiết kế kiến trúc bảo mật (Security View) của Đồ án BTL."
p.font.size = Pt(13)
p.font.color.rgb = TEXT_BODY
p.space_before = Pt(10)


# ==============================================================================
# CHUYÊN ĐỀ 1: SQL INJECTION (NGUYỄN DANH HỌC) — 4 SLIDES
# ==============================================================================
# Slide 3: Bản chất SQLi
s3 = prs.slides.add_slide(blank_layout)
add_header(s3, "1.1. SQL Injection: Bản Chất Kỹ Thuật & Nguyên Nhân Gốc Rễ", "CHỦ ĐỀ 1 • NGUYỄN DANH HỌC (23A1001D0158)")

c3_1 = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.4), Inches(5.65), Inches(5.3))
c3_1.fill.solid()
c3_1.fill.fore_color.rgb = CARD_BG
c3_1.line.color.rgb = CARD_BORDER
tf = c3_1.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "Nguyên Nhân Cốt Lõi: Nhầm Lẫn Giữa Code & Data"
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = ACCENT_RED
p = tf.add_paragraph()
p.text = "• Không phân tách giữa DỮ LIỆU (Data) và CÂU LỆNH (Code).\n• Lập trình viên sử dụng kỹ thuật ghép chuỗi trực tiếp (String Concatenation hoặc Interpolation) để tạo câu SQL từ input của người dùng.\n• Hệ quản trị CSDL khi phân tích cú pháp (Parsing) không thể biết đâu là dữ liệu của user và đâu là cấu trúc lệnh do lập trình viên định nghĩa.\n• Ký tự điều khiển (', --, ;) giúp hacker 'thoát' khỏi phạm vi dữ liệu và ép SQL Server thực thi logic của hacker."
p.font.size = Pt(12.5)
p.font.color.rgb = TEXT_BODY
p.space_before = Pt(8)

c3_2 = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.85), Inches(1.4), Inches(5.65), Inches(5.3))
c3_2.fill.solid()
c3_2.fill.fore_color.rgb = RED_LIGHT
c3_2.line.color.rgb = RED_BORDER
tf = c3_2.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "Đoạn Code Gây Lỗi Thực Tế Trong Controller"
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = ACCENT_RED
p = tf.add_paragraph()
p.text = "// LỖI NỐI CHUỖI TRỰC TIẾP:\nstring rawSql = $\"SELECT * FROM Accounts \" +\n                $\"WHERE Username = '{username}' \" +\n                $\"AND Password = '{password}'\";\n\nvar accounts = _context.Accounts\n                       .FromSqlRaw(rawSql)\n                       .ToList();"
p.font.name = "Courier New"
p.font.size = Pt(11)
p.font.color.rgb = RGBColor(159, 18, 57)
p.space_before = Pt(8)
p = tf.add_paragraph()
p.text = "⚠️ Điểm mù chết người:\nLập trình viên mặc định tin rằng input chỉ chứa chữ cái. Khi hacker truyền vào dấu nháy đơn ('), toàn bộ cây cú pháp của câu truy vấn bị bẻ gãy hoàn toàn."
p.font.size = Pt(12)
p.font.color.rgb = TEXT_HEAD
p.space_before = Pt(14)

# Slide 4: Độ nguy hiểm SQLi
s4 = prs.slides.add_slide(blank_layout)
add_header(s4, "1.2. SQL Injection: Thang Đo Nguy Hiểm & Hậu Quả Thực Tế", "CHỦ ĐỀ 1 • NGUYỄN DANH HỌC (23A1001D0158)")
col_w = Inches(3.64)
gap = Inches(0.39)

c4_1 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.4), col_w, Inches(5.3))
c4_1.fill.solid()
c4_1.fill.fore_color.rgb = RED_LIGHT
c4_1.line.color.rgb = RED_BORDER
tf = c4_1.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "🚨 CHIẾM QUYỀN HỆ THỐNG\n(Authentication Bypass)"
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = ACCENT_RED
p = tf.add_paragraph()
p.text = "• Không cần biết mật khẩu của bất kỳ ai.\n• Chỉ với 1 chuỗi payload ngắn, kẻ tấn công đăng nhập thẳng vào tài khoản Quản trị viên (Admin).\n• Chiếm đoạt phiên làm việc, đổi quyền hạn và kiểm soát toàn bộ hệ thống."
p.font.size = Pt(12)
p.font.color.rgb = TEXT_BODY
p.space_before = Pt(10)

c4_2 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8) + col_w + gap, Inches(1.4), col_w, Inches(5.3))
c4_2.fill.solid()
c4_2.fill.fore_color.rgb = AMBER_LIGHT
c4_2.line.color.rgb = AMBER_BORDER
tf = c4_2.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "📂 ĐÁNH CẮP TOÀN BỘ CSDL\n(Data Exfiltration)"
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = ACCENT_AMBER
p = tf.add_paragraph()
p.text = "• Trích xuất thông tin khách hàng, hồ sơ chuyên gia, số dư tài khoản ngân hàng.\n• Rò rỉ mật khẩu và bệnh án cá nhân (Secret Notes).\n• Vi phạm nghiêm trọng an ninh dữ liệu, phá vỡ uy tín của nền tảng tư vấn trực tuyến."
p.font.size = Pt(12)
p.font.color.rgb = TEXT_BODY
p.space_before = Pt(10)

c4_3 = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8) + (col_w + gap)*2, Inches(1.4), col_w, Inches(5.3))
c4_3.fill.solid()
c4_3.fill.fore_color.rgb = BLUE_LIGHT
c4_3.line.color.rgb = BLUE_BORDER
tf = c4_3.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "💣 PHÁ HỦY / THỰC THI LỆNH\n(System Takeover)"
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = ACCENT_BLUE
p = tf.add_paragraph()
p.text = "• Xóa sạch CSDL với lệnh DROP TABLE hoặc TRUNCATE.\n• Giả mạo số dư tiền trong tài khoản.\n• Trong các CSDL cấu hình lỏng lẻo, hacker có thể bật xp_cmdshell để chiếm quyền điều khiển cả máy chủ hệ điều hành."
p.font.size = Pt(12)
p.font.color.rgb = TEXT_BODY
p.space_before = Pt(10)

# Slide 5: Kịch bản khai thác SQLi & Live Demo
s5 = prs.slides.add_slide(blank_layout)
add_header(s5, "1.3. SQL Injection: Phân Tích Cú Pháp Khai Thác & Live Demo", "CHỦ ĐỀ 1 • NGUYỄN DANH HỌC (23A1001D0158)")

c5_main = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.4), Inches(11.733), Inches(4.2))
c5_main.fill.solid()
c5_main.fill.fore_color.rgb = CARD_BG
c5_main.line.color.rgb = CARD_BORDER
tf = c5_main.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "Kịch bản: Bắn Payload Bypass Xác Thực Đăng Nhập Admin"
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = ACCENT_RED

p = tf.add_paragraph()
p.text = "Payload đưa vào ô Username: ' OR '1'='1' --  | Mật khẩu: (nhập bất kỳ hoặc bỏ trống)\nCâu truy vấn thực tế được SQL Server thông dịch:\nSELECT * FROM Accounts WHERE Username = '' OR '1'='1' --' AND Password = 'xyz'"
p.font.name = "Courier New"
p.font.size = Pt(12)
p.font.bold = True
p.font.color.rgb = ACCENT_BLUE
p.space_before = Pt(6)

p = tf.add_paragraph()
p.text = "Giải phẫu 3 thành phần payload:\n1. Dấu nháy đơn (') : Đóng sớm chuỗi Username, đưa ngữ cảnh trở lại câu lệnh SQL.\n2. Mệnh đề OR '1'='1' : Biểu thức logic chân lý (luôn TRUE cho mọi bản ghi trong bảng Accounts).\n3. Ký tự chú thích (--) : Vô hiệu hóa hoàn toàn bước kiểm tra mật khẩu phía sau (AND Password = '...')."
p.font.size = Pt(12)
p.font.color.rgb = TEXT_BODY
p.space_before = Pt(10)

callout5 = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(5.8), Inches(11.733), Inches(0.95))
callout5.fill.solid()
callout5.fill.fore_color.rgb = BLUE_LIGHT
callout5.line.color.rgb = BLUE_BORDER
tf = callout5.text_frame
p = tf.paragraphs[0]
p.text = "💻 [HỌC LIVE DEMO]: Chuyển sang Web UI Tab 1 & VS Code SQL Server"
p.font.size = Pt(12.5)
p.font.bold = True
p.font.color.rgb = ACCENT_BLUE
p = tf.add_paragraph()
p.text = "Thao tác trực tiếp: Nạp payload ' OR '1'='1' -- -> Bấm submit bên Đỏ (rò rỉ 4 tài khoản) -> Chuyển VS Code đối chiếu query."
p.font.size = Pt(11)
p.font.color.rgb = TEXT_BODY
p.space_before = Pt(2)

# Slide 6: Cơ chế phòng thủ SQLi
s6 = prs.slides.add_slide(blank_layout)
add_header(s6, "1.4. SQL Injection: Cơ Chế Phòng Thủ Chuẩn EF Core LINQ", "CHỦ ĐỀ 1 • NGUYỄN DANH HỌC (23A1001D0158)")

c6_1 = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.4), Inches(5.65), Inches(5.3))
c6_1.fill.solid()
c6_1.fill.fore_color.rgb = GREEN_LIGHT
c6_1.line.color.rgb = GREEN_BORDER
tf = c6_1.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "Cách Viết Code Chuẩn Hóa Với EF Core"
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = ACCENT_GREEN
p = tf.add_paragraph()
p.text = "// DÙNG LINQ PARAMETERIZED (KHUYÊN DÙNG):\nvar account = _context.Accounts\n    .FirstOrDefault(a => a.Username == user \n                      && a.Password == pass);\n\n// HOẶC DÙNG FROMSQLINTERPOLATED (NẾU DÙNG SQL THUẦN):\nvar account = _context.Accounts\n    .FromSqlInterpolated($\"SELECT * FROM Accounts \" +\n                         $\"WHERE Username={user} \" +\n                         $\"AND Password={pass}\")\n    .FirstOrDefault();"
p.font.name = "Courier New"
p.font.size = Pt(10)
p.font.color.rgb = RGBColor(6, 95, 70)
p.space_before = Pt(8)
p = tf.add_paragraph()
p.text = "Tại sao cách này an toàn tuyệt đối?\nSQL Server biên dịch cây thực thi (Execution Plan) TRƯỚC khi gán dữ liệu. Biến @p0 nhận toàn bộ payload như một chuỗi chữ bình thường, vô hiệu hóa hoàn toàn ý đồ chèn lệnh."
p.font.size = Pt(12)
p.font.color.rgb = TEXT_HEAD
p.space_before = Pt(12)

c6_2 = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.85), Inches(1.4), Inches(5.65), Inches(5.3))
c6_2.fill.solid()
c6_2.fill.fore_color.rgb = CARD_BG
c6_2.line.color.rgb = CARD_BORDER
tf = c6_2.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "Bộ Quy Tắc Phòng Thủ Toàn Diện (Defense-in-Depth)"
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = ACCENT_BLUE
p = tf.add_paragraph()
p.text = "1. Luôn sử dụng Parameterized Query / ORM:\nTuyệt đối không ghép chuỗi SQL thủ công dưới bất kỳ hình thức nào.\n\n2. Nguyên tắc đặc quyền tối thiểu (Least Privilege):\nTài khoản kết nối CSDL của ứng dụng chỉ có quyền SELECT/INSERT/UPDATE trên các bảng cần thiết, không bao giờ dùng tài khoản 'sa' hay quyền DDL (DROP, ALTER).\n\n3. Xác thực dữ liệu đầu vào (Input Validation):\nSử dụng Data Annotations ([RegularExpression], [StringLength]) kiểm tra chặt chẽ khuôn dạng dữ liệu.\n\n4. Mã hóa mật khẩu một chiều (Password Hashing):\nSử dụng ASP.NET Core Identity (PBKDF2/BCrypt) để nếu CSDL có bị lộ, mật khẩu vẫn được bảo vệ an toàn."
p.font.size = Pt(12)
p.font.color.rgb = TEXT_BODY
p.space_before = Pt(10)


# ==============================================================================
# CHUYÊN ĐỀ 2: STORED XSS (NGUYỄN THANH BÌNH) — 4 SLIDES
# ==============================================================================
# Slide 7: Bản chất XSS
s7 = prs.slides.add_slide(blank_layout)
add_header(s7, "2.1. Stored XSS: Bản Chất Kỹ Thuật & Ngữ Cảnh Bài Toán", "CHỦ ĐỀ 2 • NGUYỄN THANH BÌNH (23A1001D0041)")

c7_1 = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.4), Inches(5.65), Inches(5.3))
c7_1.fill.solid()
c7_1.fill.fore_color.rgb = CARD_BG
c7_1.line.color.rgb = CARD_BORDER
tf = c7_1.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "Ngữ Cảnh Nghiệp Vụ Trong Website Tư Vấn"
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = ACCENT_RED
p = tf.add_paragraph()
p.text = "• Trong website tư vấn trực tuyến, khách hàng có chức năng đăng đánh giá, nhận xét và gửi câu hỏi tư vấn.\n• Bản chất Stored XSS (Persistent XSS): Kẻ tấn công nhúng mã JavaScript độc hại vào nội dung nhận xét.\n• Máy chủ không lọc mà lưu thẳng mã độc vào CSDL (bảng dbo.Comments).\n• Nguy hiểm: Mỗi khi bất kỳ người dùng khác, chuyên gia hoặc quản trị viên mở trang đánh giá, trình duyệt của họ sẽ tự động biên dịch và thực thi đoạn script độc hại đó."
p.font.size = Pt(12.5)
p.font.color.rgb = TEXT_BODY
p.space_before = Pt(8)

c7_2 = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.85), Inches(1.4), Inches(5.65), Inches(5.3))
c7_2.fill.solid()
c7_2.fill.fore_color.rgb = RED_LIGHT
c7_2.line.color.rgb = RED_BORDER
tf = c7_2.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "Mã Nguồn C# Gây Lỗi & Sự Lạm Dụng @@Html.Raw"
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = ACCENT_RED
p = tf.add_paragraph()
p.text = "// 1. CONTROLLER LƯU NGUYÊN VĂN KHÔNG LÀM SẠCH:\nvar comment = new Comment {\n    Author = author,\n    Content = content // Chứa thẻ <script> độc hại\n};\n_context.Comments.Add(comment);\n\n// 2. RAZOR VIEW LẠM DỤNG HTML.RAW:\n<div class=\"comment-body\">\n    @Html.Raw(comment.Content) \n</div>"
p.font.name = "Courier New"
p.font.size = Pt(10.5)
p.font.color.rgb = RGBColor(159, 18, 57)
p.space_before = Pt(8)
p = tf.add_paragraph()
p.text = "⚠️ Điểm mù lập trình viên:\nSử dụng @Html.Raw() với mục đích cho phép hiển thị format HTML nhưng vô tình mở toang cánh cửa cho hacker thực thi JavaScript trên máy nạn nhân."
p.font.size = Pt(12)
p.font.color.rgb = TEXT_HEAD
p.space_before = Pt(12)

# Slide 8: Độ nguy hiểm XSS
s8 = prs.slides.add_slide(blank_layout)
add_header(s8, "2.2. Stored XSS: Mối Đe Dọa Đánh Cắp Cookie & Session", "CHỦ ĐỀ 2 • NGUYỄN THANH BÌNH (23A1001D0041)")

c8_1 = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.4), col_w, Inches(5.3))
c8_1.fill.solid()
c8_1.fill.fore_color.rgb = RED_LIGHT
c8_1.line.color.rgb = RED_BORDER
tf = c8_1.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "🍪 ĐÁNH CẮP SESSION COOKIE\n(Cookie Stealing)"
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = ACCENT_RED
p = tf.add_paragraph()
p.text = "• Mã độc truy cập document.cookie để trích xuất Session Token của nạn nhân.\n• Tự động gửi token về máy chủ của hacker.\n• Hacker giả mạo danh tính nạn nhân (Session Hijacking) đăng nhập vào tài khoản mà không cần mật khẩu."
p.font.size = Pt(12)
p.font.color.rgb = TEXT_BODY
p.space_before = Pt(10)

c8_2 = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8) + col_w + gap, Inches(1.4), col_w, Inches(5.3))
c8_2.fill.solid()
c8_2.fill.fore_color.rgb = AMBER_LIGHT
c8_2.line.color.rgb = AMBER_BORDER
tf = c8_2.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "🎣 THAY ĐỔI NỘI DUNG GIAO DIỆN\n(DOM Defacement & Phishing)"
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = ACCENT_AMBER
p = tf.add_paragraph()
p.text = "• Hacker dùng JavaScript chèn form đăng nhập giả mạo đè lên giao diện website thật.\n• Lừa người dùng nhập lại tài khoản ngân hàng hoặc mật khẩu.\n• Chuyển hướng người dùng sang trang web lừa đảo (Open Redirect)."
p.font.size = Pt(12)
p.font.color.rgb = TEXT_BODY
p.space_before = Pt(10)

c8_3 = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8) + (col_w + gap)*2, Inches(1.4), col_w, Inches(5.3))
c8_3.fill.solid()
c8_3.fill.fore_color.rgb = BLUE_LIGHT
c8_3.line.color.rgb = BLUE_BORDER
tf = c8_3.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "🤖 THỰC HIỆN THAO TÁC NGẦM\n(XSS-Keylogger)"
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = ACCENT_BLUE
p = tf.add_paragraph()
p.text = "• Lắng nghe sự kiện bàn phím của nạn nhân (Keylogging) khi họ chat với chuyên gia.\n• Tự động gửi tin nhắn hoặc đánh giá 5 sao giả mạo danh nghĩa của người dùng.\n• Biến trình duyệt nạn nhân thành botnet tấn công từ chối dịch vụ."
p.font.size = Pt(12)
p.font.color.rgb = TEXT_BODY
p.space_before = Pt(10)

# Slide 9: Kịch bản XSS & Live Demo
s9 = prs.slides.add_slide(blank_layout)
add_header(s9, "2.3. Stored XSS: Kịch Bản Khai Thác & Live Demo Tương Tác", "CHỦ ĐỀ 2 • NGUYỄN THANH BÌNH (23A1001D0041)")

c9_main = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.4), Inches(11.733), Inches(4.2))
c9_main.fill.solid()
c9_main.fill.fore_color.rgb = CARD_BG
c9_main.line.color.rgb = CARD_BORDER
tf = c9_main.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "Kịch bản: Chèn Script Đánh Cắp Cookie Vào Chức Năng Đánh Giá"
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = ACCENT_RED

p = tf.add_paragraph()
p.text = "Payload 1 (Thẻ Script): <script>alert('Lộ Cookie: ' + document.cookie);</script>\nPayload 2 (Bypass qua thẻ ảnh): <img src=x onerror=alert('XSS kích hoạt qua onerror!')>\nPayload 3 (Defacement): <script>document.body.innerHTML='<h1>WEB ĐÃ BỊ HACK</h1>';</script>"
p.font.name = "Courier New"
p.font.size = Pt(11.5)
p.font.bold = True
p.font.color.rgb = ACCENT_BLUE
p.space_before = Pt(6)

p = tf.add_paragraph()
p.text = "Cơ chế kích hoạt thực nghiệm:\n1. Hacker gửi bình luận chứa payload lên máy chủ qua form dính lỗi.\n2. Máy chủ lưu trực tiếp đoạn mã vào bảng dbo.Comments.\n3. Khi tải lại trang, trình duyệt đọc thấy thẻ <script> và lập tức thực thi: Popup Alert bật lên hiển thị giá trị AuthSessionToken giả lập được lưu trong cookie!"
p.font.size = Pt(12)
p.font.color.rgb = TEXT_BODY
p.space_before = Pt(10)

callout9 = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(5.8), Inches(11.733), Inches(0.95))
callout9.fill.solid()
callout9.fill.fore_color.rgb = RED_LIGHT
callout9.line.color.rgb = RED_BORDER
tf = callout9.text_frame
p = tf.paragraphs[0]
p.text = "💬 [BÌNH LIVE DEMO]: Chuyển sang Web UI Tab 2 (Stored XSS)"
p.font.size = Pt(12.5)
p.font.bold = True
p.font.color.rgb = ACCENT_RED
p = tf.add_paragraph()
p.text = "Thao tác: Bấm nạp payload đánh cắp Cookie -> Đăng bên Lỗi (bật popup alert lộ Session) -> Đăng bên An toàn (chữ hiển thị an toàn)."
p.font.size = Pt(11)
p.font.color.rgb = TEXT_BODY
p.space_before = Pt(2)

# Slide 10: Phòng thủ XSS
s10 = prs.slides.add_slide(blank_layout)
add_header(s10, "2.4. Stored XSS: Cơ Chế Phòng Thủ Chuẩn Trong ASP.NET Core", "CHỦ ĐỀ 2 • NGUYỄN THANH BÌNH (23A1001D0041)")

c10_1 = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.4), Inches(5.65), Inches(5.3))
c10_1.fill.solid()
c10_1.fill.fore_color.rgb = GREEN_LIGHT
c10_1.line.color.rgb = GREEN_BORDER
tf = c10_1.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "Cơ Chế HTML Encoding Tự Động Của Razor View"
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = ACCENT_GREEN
p = tf.add_paragraph()
p.text = "// PHÒNG THỦ MẶC ĐỊNH (KHÔNG DÙNG HTML.RAW):\n<p>@comment.Content</p>\n\n// HOẶC MÃ HÓA TẠI CONTROLLER TRƯỚC KHI LƯU:\nstring encoded = HtmlEncoder.Default.Encode(content);\ncomment.Content = encoded;"
p.font.name = "Courier New"
p.font.size = Pt(10.5)
p.font.color.rgb = RGBColor(6, 95, 70)
p.space_before = Pt(8)
p = tf.add_paragraph()
p.text = "Cơ chế biến đổi ký tự nguy hiểm:\n• Ký tự '<' biến thành &lt;\n• Ký tự '>' biến thành &gt;\n• Dấu ngoặc kép '\"' biến thành &quot;\nTrình duyệt thông dịch đây là ký tự văn bản thuần túy để hiển thị, tuyệt đối không coi là thẻ thực thi mã JavaScript."
p.font.size = Pt(12)
p.font.color.rgb = TEXT_HEAD
p.space_before = Pt(12)

c10_2 = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.85), Inches(1.4), Inches(5.65), Inches(5.3))
c10_2.fill.solid()
c10_2.fill.fore_color.rgb = CARD_BG
c10_2.line.color.rgb = CARD_BORDER
tf = c10_2.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "Bộ Giải Pháp Phòng Vệ Chiều Sâu (Defense-in-Depth)"
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = ACCENT_BLUE
p = tf.add_paragraph()
p.text = "1. Bật cờ HttpOnly = true cho Cookie phiên:\nNgăn chặn tuyệt đối mọi đoạn mã JavaScript (kể cả khi đã inject thành công) truy cập vào document.cookie.\n\n2. Cấu hình Content Security Policy (CSP Header):\nThiết lập HTTP Header chỉ cho phép thực thi script từ nguồn đáng tin cậy, chặn đứng inline script (<script>...</script>).\n\n3. Làm sạch dữ liệu đầu vào (Input Sanitization):\nSử dụng thư viện HtmlSanitizer để loại bỏ triệt để các thẻ nguy hiểm (<script>, <iframe>, <object>, thuộc tính onerror/onload)."
p.font.size = Pt(12)
p.font.color.rgb = TEXT_BODY
p.space_before = Pt(10)


# ==============================================================================
# CHUYÊN ĐỀ 3: CSRF ATTACK (NGUYỄN MINH CƯỜNG) — 4 SLIDES
# ==============================================================================
# Slide 11: Bản chất CSRF
s11 = prs.slides.add_slide(blank_layout)
add_header(s11, "3.1. CSRF Attack: Bản Chất Kỹ Thuật & Cơ Chế Giả Mạo", "CHỦ ĐỀ 3 • NGUYỄN MINH CƯỜNG (23A1001D0058)")

c11_1 = s11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.4), Inches(5.65), Inches(5.3))
c11_1.fill.solid()
c11_1.fill.fore_color.rgb = CARD_BG
c11_1.line.color.rgb = CARD_BORDER
tf = c11_1.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "Bản Chất Của Lỗ Hổng Cross-Site Request Forgery"
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = ACCENT_RED
p = tf.add_paragraph()
p.text = "• CSRF (A01:2021 - Broken Access Control) là kỹ thuật tấn công mượn quyền của người dùng hợp lệ.\n• Điều kiện xảy ra:\n  1. Nạn nhân đã đăng nhập vào hệ thống tư vấn trực tuyến (phiên làm việc Session/Cookie vẫn đang hợp lệ).\n  2. Trình duyệt tự động đính kèm Cookie xác thực vào mọi HTTP Request gửi tới hệ thống đó.\n  3. Kẻ tấn công lừa nạn nhân truy cập một trang web độc hại ở một tab khác (Attacker Site).\n• Điểm mù: Máy chủ chỉ kiểm tra Cookie mà không kiểm tra xem Request đó bắt nguồn từ đâu!"
p.font.size = Pt(12.5)
p.font.color.rgb = TEXT_BODY
p.space_before = Pt(8)

c11_2 = s11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.85), Inches(1.4), Inches(5.65), Inches(5.3))
c11_2.fill.solid()
c11_2.fill.fore_color.rgb = RED_LIGHT
c11_2.line.color.rgb = RED_BORDER
tf = c11_2.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "Action Bị Lỗi: Thiếu Thuộc Tính Phòng Vệ"
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = ACCENT_RED
p = tf.add_paragraph()
p.text = "// ACTION NHẬN TIỀN KHÔNG CÓ BẢO VỆ TOKEN:\n[HttpPost]\npublic IActionResult TransferVulnerable(decimal amount)\n{\n    var victim = _context.Wallets.Find(1);\n    var hacker = _context.Wallets.Find(2);\n    \n    // Trừ tiền nạn nhân, cộng tiền hacker:\n    victim.Balance -= amount;\n    hacker.Balance += amount;\n    _context.SaveChanges();\n    return RedirectToAction(\"Index\");\n}"
p.font.name = "Courier New"
p.font.size = Pt(10)
p.font.color.rgb = RGBColor(159, 18, 57)
p.space_before = Pt(8)
p = tf.add_paragraph()
p.text = "⚠️ Nguy hiểm:\nBất kỳ trang web nào trên Internet cũng có thể gửi form POST tới URL này. Máy chủ thấy Cookie hợp lệ sẽ trừ tiền ngay lập tức!"
p.font.size = Pt(12)
p.font.color.rgb = TEXT_HEAD
p.space_before = Pt(10)

# Slide 12: Kịch bản bẫy trúng thưởng CSRF
s12 = prs.slides.add_slide(blank_layout)
add_header(s12, "3.2. CSRF Attack: Kịch Bản Lừa Đảo Chuyển Tiền Ví Ngầm", "CHỦ ĐỀ 3 • NGUYỄN MINH CƯỜNG (23A1001D0058)")

c12_main = s12.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.4), Inches(11.733), Inches(5.3))
c12_main.fill.solid()
c12_main.fill.fore_color.rgb = RED_LIGHT
c12_main.line.color.rgb = RED_BORDER
tf = c12_main.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "Kịch bản: Trang Web Giả Mạo Trúng Thưởng (Attacker Site)"
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = ACCENT_RED

p = tf.add_paragraph()
p.text = "1. Nạn nhân có 50.000.000 VNĐ trong ví đang mở tab làm việc tại nền tảng tư vấn trực tuyến.\n2. Kẻ tấn công gửi link: 'Chúc mừng bạn trúng thưởng iPhone 16 Pro Max! Click để nhận quà!'.\n3. Khi nạn nhân click vào link, một trang web bẫy (AttackerSite.cshtml) được mở ra, bên trong chứa mã HTML ẩn:"
p.font.size = Pt(13)
p.font.color.rgb = TEXT_BODY
p.space_before = Pt(8)

p = tf.add_paragraph()
p.text = "<!-- FORM ẨN TỰ ĐỘNG GỬI SANG HỆ THỐNG MỤC TIÊU -->\n<form id=\"exploit\" action=\"http://localhost:5076/Csrf/TransferVulnerable\" method=\"POST\">\n    <input type=\"hidden\" name=\"amount\" value=\"20000000\" />\n</form>\n<script>document.getElementById('exploit').submit();</script>"
p.font.name = "Courier New"
p.font.size = Pt(11)
p.font.bold = True
p.font.color.rgb = RGBColor(159, 18, 57)
p.space_before = Pt(6)

p = tf.add_paragraph()
p.text = "Hậu quả thực tế: Nạn nhân không hề bấm chuyển tiền, nhưng ví tiền bị trừ ngay 20.000.000 VNĐ sang ví hacker!"
p.font.size = Pt(13)
p.font.bold = True
p.font.color.rgb = ACCENT_RED
p.space_before = Pt(8)

# Slide 13: Live demo CSRF
s13 = prs.slides.add_slide(blank_layout)
add_header(s13, "3.3. CSRF Attack: Thực Nghiệm Trực Tiếp & Kiểm Thử Rút Tiền", "CHỦ ĐỀ 3 • NGUYỄN MINH CƯỜNG (23A1001D0058)")

c13_main = s13.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.4), Inches(11.733), Inches(4.2))
c13_main.fill.solid()
c13_main.fill.fore_color.rgb = CARD_BG
c13_main.line.color.rgb = CARD_BORDER
tf = c13_main.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "Quy Trình 3 Bước Thực Nghiệm Trên Giao Diện Web"
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = ACCENT_GREEN

p = tf.add_paragraph()
p.text = "• Bước 1: Kiểm tra số dư ví ban đầu tại Tab 3:\n  - Ví Nạn nhân: 50.000.000 VNĐ | Ví Hacker: 0 VNĐ.\n• Bước 2: Bấm nút đỏ 'Mở trang web bẫy của Hacker (Attacker Site)' ở tab mới -> Bấm nút 'Nhận thưởng ngay':\n  - Ngay lập tức, số dư của nạn nhân bị trừ 20.000.000 VNĐ và chuyển sang ví hacker!\n• Bước 3: Thử nghiệm chuyển tiền tại Form bên phải (Đã phòng thủ):\n  - Giao dịch chỉ hợp lệ khi được gửi từ chính form trên hệ thống với Anti-Forgery Token hợp lệ."
p.font.size = Pt(13)
p.font.color.rgb = TEXT_BODY
p.space_before = Pt(8)

callout13 = s13.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(5.8), Inches(11.733), Inches(0.95))
callout13.fill.solid()
callout13.fill.fore_color.rgb = GREEN_LIGHT
callout13.line.color.rgb = GREEN_BORDER
tf = callout13.text_frame
p = tf.paragraphs[0]
p.text = "💳 [CƯỜNG LIVE DEMO]: Chuyển sang Web UI Tab 3 (CSRF)"
p.font.size = Pt(12.5)
p.font.bold = True
p.font.color.rgb = ACCENT_GREEN
p = tf.add_paragraph()
p.text = "Thao tác: Mở Attacker Site -> Bấm nhận thưởng -> Quay lại thấy số dư ví nạn nhân bị trừ 20 triệu -> Thử lại form có Token."
p.font.size = Pt(11)
p.font.color.rgb = TEXT_BODY
p.space_before = Pt(2)

# Slide 14: Phòng thủ CSRF
s14 = prs.slides.add_slide(blank_layout)
add_header(s14, "3.4. CSRF Attack: Cơ Chế Phòng Thủ Bằng Anti-Forgery Token", "CHỦ ĐỀ 3 • NGUYỄN MINH CƯỜNG (23A1001D0058)")

c14_1 = s14.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.4), Inches(5.65), Inches(5.3))
c14_1.fill.solid()
c14_1.fill.fore_color.rgb = GREEN_LIGHT
c14_1.line.color.rgb = GREEN_BORDER
tf = c14_1.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "Cơ Chế Synchronizer Token Pattern Trong ASP.NET Core"
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = ACCENT_GREEN
p = tf.add_paragraph()
p.text = "// 1. TRONG RAZOR VIEW: NHÚNG TOKEN VÀO FORM\n<form asp-action=\"TransferSecure\" method=\"post\">\n    @Html.AntiForgeryToken()\n    <input type=\"number\" name=\"amount\" />\n    <button type=\"submit\">Chuyển tiền</button>\n</form>\n\n// 2. TRONG CONTROLLER: BẮT BUỘC KIỂM TRA TOKEN\n[HttpPost]\n[ValidateAntiForgeryToken]\npublic IActionResult TransferSecure(decimal amount)\n{\n    // Chỉ thực thi khi Token từ Form khớp với Token trong Cookie\n}"
p.font.name = "Courier New"
p.font.size = Pt(9.5)
p.font.color.rgb = RGBColor(6, 95, 70)
p.space_before = Pt(8)

c14_2 = s14.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.85), Inches(1.4), Inches(5.65), Inches(5.3))
c14_2.fill.solid()
c14_2.fill.fore_color.rgb = CARD_BG
c14_2.line.color.rgb = CARD_BORDER
tf = c14_2.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "Tại Sao Kẻ Tấn Công Không Thể Vượt Qua?"
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = ACCENT_BLUE
p = tf.add_paragraph()
p.text = "1. Nguyên tắc cặp Token:\nMáy chủ phát hành 1 token mã hóa trong Cookie và 1 token bí mật nhúng vào trường input ẩn của Form.\n\n2. Bảo vệ bằng Same-Origin Policy (SOP):\nTrang web độc hại của hacker ở tên miền khác (Cross-Origin) TUYỆT ĐỐI KHÔNG THỂ đọc trộm token bí mật trong form của nạn nhân.\n\n3. Cấu hình Cookie SameSite = Strict:\nNgăn chặn trình duyệt tự động gửi cookie trong các cross-site request, triệt tiêu tận gốc tiền đề tấn công của CSRF."
p.font.size = Pt(12)
p.font.color.rgb = TEXT_BODY
p.space_before = Pt(10)


# ==============================================================================
# TỔNG KẾT & ĐỒ ÁN BTL — 2 SLIDES
# ==============================================================================
# Slide 15: Bảng so sánh tổng kết 3 lỗ hổng
s15 = prs.slides.add_slide(blank_layout)
add_header(s15, "Tổng Kết: So Sánh Đối Chứng 3 Lỗ Hổng Web Kinh Điển")

table_shape = s15.shapes.add_table(4, 4, Inches(0.8), Inches(1.5), Inches(11.733), Inches(5.2))
tbl = table_shape.table

headers = ["Tiêu chí so sánh", "SQL Injection (Học)", "Stored XSS (Bình)", "CSRF Attack (Cường)"]
for i, h in enumerate(headers):
    cell = tbl.cell(0, i)
    cell.fill.solid()
    cell.fill.fore_color.rgb = ACCENT_BLUE
    p = cell.text_frame.paragraphs[0]
    p.text = h
    p.font.bold = True
    p.font.size = Pt(12)
    p.font.color.rgb = WHITE

rows_data = [
    ("Mục tiêu tấn công", "Tầng Cơ sở dữ liệu (SQL Server Engine)", "Tầng Trình duyệt người dùng (Client DOM)", "Quyền hạn của phiên làm việc (Session/Cookie)"),
    ("Mức độ nghiêm trọng", "CRITICAL (9.8 / 10.0)", "HIGH (7.5 - 8.5 / 10.0)", "MEDIUM / HIGH (6.5 - 8.0 / 10.0)"),
    ("Giải pháp cốt lõi", "Parameterized Query / EF Core LINQ", "Razor HTML Encoding + HttpOnly Cookie", "Anti-Forgery Token + SameSite Cookie")
]

for row_idx, rdata in enumerate(rows_data, start=1):
    for col_idx, text in enumerate(rdata):
        cell = tbl.cell(row_idx, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = CARD_BG if row_idx % 2 == 1 else WHITE
        p = cell.text_frame.paragraphs[0]
        p.text = text
        p.font.size = Pt(11)
        p.font.color.rgb = TEXT_HEAD if col_idx == 0 else TEXT_BODY
        if col_idx == 1 and "CRITICAL" in text:
            p.font.bold = True
            p.font.color.rgb = ACCENT_RED

# Slide 16: Cam kết đồ án BTL & Lời cảm ơn
s16 = prs.slides.add_slide(blank_layout)
add_header(s16, "Áp Dụng Kiến Trúc An Toàn Vào Đồ Án Website Tư Vấn Trực Tuyến")

main_c16 = s16.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.4), Inches(11.733), Inches(5.3))
main_c16.fill.solid()
main_c16.fill.fore_color.rgb = CARD_BG
main_c16.line.color.rgb = CARD_BORDER

tf = main_c16.text_frame
tf.word_wrap = True

p = tf.paragraphs[0]
p.text = "CAM KẾT CHUẨN AN TOÀN ĐA TẦNG CHO BTL (NHÓM WNC.G01):"
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = ACCENT_BLUE

p = tf.add_paragraph()
p.text = "1. Kiến trúc bảo mật (Security View - Buổi 8): Toàn bộ hệ thống được xây dựng theo nguyên tắc 'Security by Design'. Không xem an ninh là tính năng chắp vá mà tích hợp sâu vào kiến trúc 3 tầng.\n\n2. Bảo vệ CSDL tuyệt đối (Nguyễn Danh Học): 100% các câu truy vấn đặt lịch, hồ sơ chuyên gia và đánh giá chất lượng được quản lý bằng Entity Framework Core Code-First với LINQ Parameterized.\n\n3. Bảo vệ dữ liệu người dùng (Nguyễn Thanh Bình): 100% nội dung phản hồi tư vấn và bình luận được Razor View mã hóa tự động, Cookie cấu hình HttpOnly = true và Secure = true.\n\n4. Bảo vệ giao dịch (Nguyễn Minh Cường): Bắt buộc kiểm tra Anti-Forgery Token trên toàn bộ các form đặt lịch hẹn, nạp tiền ví và đổi mật khẩu."
p.font.size = Pt(12.5)
p.font.color.rgb = TEXT_BODY
p.space_before = Pt(10)

p = tf.add_paragraph()
p.text = "XIN TRÂN TRỌNG CẢM ƠN THẦY VÀ CÁC BẠN ĐÃ THEO DÕI!\nNhóm WNC.G01 sẵn sàng lắng nghe câu hỏi và nhận xét từ giảng viên."
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = ACCENT_RED
p.space_before = Pt(16)

out_file = "/media/hocjsoo/New Volume/OWASP_Demo/Slide_OWASP_Top3_Nhom_WNC_G01.pptx"
prs.save(out_file)
print(f"Master 16-slides presentation saved successfully to: {out_file}")
