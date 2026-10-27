# Nhật ký thay đổi
## 23/10/2026
- Đối chiếu hồ sơ với Giấy chứng nhận đăng ký doanh nghiệp (thay đổi lần 3, 09/06/2026): MST, địa chỉ, điện thoại khớp. Bổ sung loại hình, vốn điều lệ 10 tỷ, chủ sở hữu, địa điểm kinh doanh số 02 tại TP HCM (không phải chi nhánh; địa chỉ đầy đủ có Đường Hà Huy Giáp).
- Anh Trung duyệt phương án sửa 9 mục lệch/sai (xem SO_LOI.md): đã sửa trong ho-so/.
- Tạo 8 mẫu mới trong mau-giay-to/ bằng cong-cu/tao_mau.py. Người ký bên Trung Kiên là Giám đốc Nguyễn Việt Trung; thông tin người được ủy quyền, khách hàng để trống. Hồ sơ năng lực đã có (HSNL-Trung-Kien-2026.docx do anh gửi), không làm lại.
- Anh Trung chuyển kho sang riêng tư; đã kiểm lại: private = true.
- Theo yêu cầu anh Trung: bỏ số CCCD, ngày sinh, địa chỉ liên lạc cá nhân khỏi kho; giấy ủy quyền để trống các ô này.
- Chưa lưu bản gốc GPKD (PDF) và HSNL (docx) vào kho; cần thì báo.
- Anh Trung nhắc: công ty của anh chỉ là Trung Kiên, các công ty khác là khách hàng. Đã rà: 8 mẫu không có thông tin khách nào; bỏ việc của khách (địa chỉ giao hàng Tiên Sơn, giấy ủy quyền Phó Giám đốc Thủy sản Quảng Ninh) khỏi danh sách việc của công ty. Chờ anh xác nhận Thịnh Vượng (mục 1.2 hồ sơ) là của anh hay của khách.
- 27/10/2026: làm Giấy đề nghị thanh toán phần còn lại HĐ HĐKT102026-TK-H (Trung Kiên bán, Công ty CP Công nghệ FATA mua): 160.812.000 đồng = 291.168.000 (sau Phụ lục 01) - 130.356.000 (đã trả 50% hợp đồng gốc). Khớp bảng công nợ FATA (hóa đơn 00002771 ngày 17/10/2026). Tệp: giay-to-khach-hang/FATA/. Công cụ: cong-cu/tao_denghi_fata.py.
## 27/10/2026 - anh Trung chốt (nguyên văn tóm tắt)
- Điện thoại in giấy tờ: 096.432.9898. Không in email, website lên mẫu.
- REDPUMP: Trung Kiên là nhà phân phối của hãng.
- Thịnh Vượng là công ty của anh: "khi nào a bảo làm thịnh vượng thì e lấy thông tin này thay cho trung kiên".
- Mã ngành: bỏ qua. Giấy đề nghị FATA: đúng (đã trả 130.356.000, hàng đã giao đủ, ngày 27/10/2026). Cho phép push lên GitHub.
- 27/10/2026: chuẩn hóa bố cục theo Nghị định 30/2020/NĐ-CP (Phụ lục I): A4, lề trên/dưới 20 mm, trái 30 mm, phải 20 mm. Khối ký: hai cột bằng nhau, tiêu đề căn giữa cột. Có dấu: dấu tròn đường kính 40 mm (giả định, chờ anh đo dấu thật), chữ ký và dấu chức danh căn giữa khối ký bên Giám đốc. Chạy: python3 cong-cu/tao_mau.py anh_dau.png thu_muc_ra (ảnh không lưu trong kho).
- 27/10/2026: thêm cong-cu/chen_dau_pdf.py (dấu giáp lai + dấu chữ ký vào PDF). Đã chèn cho Hợp đồng HĐKT102026-TK-H (3 trang) và Phụ lục 01 (2 trang) của FATA; bản có dấu không lưu trong kho.
- 27/10/2026: anh Trung bảo lưu hết lên đám mây để phiên mới làm được như phiên này: lưu dau-chu-ky/, bản có dấu và bản gốc của FATA, hồ sơ năng lực, công cụ (tach_nen_muc.py, chen_dau_pdf.py), CLAUDE.md, HUONG_DAN_PHIEN_MOI.md, requirements.txt. Không lưu giấy phép đăng ký kinh doanh gốc (có số CCCD).
