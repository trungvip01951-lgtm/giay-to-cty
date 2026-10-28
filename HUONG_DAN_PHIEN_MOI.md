# Hướng dẫn làm việc cho phiên mới (cập nhật 27/10/2026)
## Việc anh Trung hay nhờ và cách làm
| Việc | Cách làm |
|---|---|
| Đề nghị thanh toán (hợp đồng chia 2 lần, hoặc phần còn lại) | Lấy mẫu giay-to-khach-hang/FATA/ và cong-cu/tao_denghi_fata.py làm gốc: đổi số hợp đồng, ngày, số tiền; tiền tính bằng mã, đối chiếu bảng công nợ nếu anh gửi; chữ bằng hàm doc_so. Có dấu: `python3 cong-cu/tao_denghi_fata.py dau-chu-ky/dau_chu_ky_moi_ghep.png thu_muc_ra` |
| Mẫu giấy tờ Trung Kiên (thanh lý, nghiệm thu, ủy quyền, giới thiệu, phụ lục, đối chiếu công nợ, đơn đặt hàng, nhận hàng NK) | `python3 cong-cu/tao_mau.py` (không dấu, ra mau-giay-to/). Có dấu: `python3 cong-cu/tao_mau.py dau-chu-ky/dau_chu_ky_moi_ghep.png thu_muc_ra` |
| Hợp đồng/phụ lục của khách (Word) cần dấu | Đổi sang PDF: `soffice --headless --convert-to pdf --outdir ra file.docx`; rồi `python3 cong-cu/chen_dau_pdf.py vao.pdf ra.pdf dau-chu-ky/dau_chu_ky_moi_ghep.png dau-chu-ky/dau_tron_sach.png`. Công cụ tự tìm khối "ĐẠI DIỆN BÊN BÁN" ở trang cuối, chèn dấu giáp lai mọi trang. Không sửa nội dung hợp đồng của khách; chỉ báo nếu thấy lỗi (ví dụ "PDG" thay vì "PGD") |
| Bản trông như mới quét (làm cho MỌI bản có dấu) | `python3 cong-cu/lam_ban_scan.py ten_CoDau.pdf ten_Scan.pdf` (cần có PDF; Word thì đổi sang PDF trước). Kết quả là PDF dạng ảnh |
| Đề nghị thanh toán theo báo giá (100% hoặc tỷ lệ khác) | Tạo giay-to-khach-hang/<khach>/don_<số báo giá>.json (mẫu: giay-to-khach-hang/HoangVienQuangBa/don_BG-28102026-006.json) rồi `python3 cong-cu/tao_denghi_100.py don.json dau-chu-ky/dau_chu_ky_moi_ghep.png`; đổi sang PDF rồi làm bản Scan |
| Ảnh dấu/chữ ký mới | `python3 cong-cu/tach_nen_muc.py anh.jpg ra.png do|xanh [tam_x,tam_y,ban_kinh]` |
## Thông số đã chốt
- A4, lề trên/dưới 20 mm, trái 30 mm, phải 20 mm (NĐ 30/2020/NĐ-CP, Phụ lục I). Times New Roman 12.
- Dấu tròn 40 mm (giả định: anh chưa đo dấu thật; đo ra khác thì sửa SEAL_CM trong cong-cu/tao_mau.py và tham số mm của chen_dau_pdf.py). Chữ ký rộng khoảng 1,7 lần dấu như ảnh gốc.
- Khối ký: hai cột bằng nhau, tiêu đề căn giữa cột. Giáp lai: dấu chia thành n lát bằng nhau (n = số trang), mỗi trang một lát sát mép phải, giữa trang.
- Điện thoại in giấy tờ 096.432.9898; không in email, website. Ngân hàng ghi một cách theo QUY_TAC.md.
- Ngày lập giấy: ngày làm, nếu anh không nói khác.
## Cài đặt môi trường
`pip install -r requirements.txt`; cần libreoffice và poppler-utils. Kiểm tra bố cục: xuất PDF, xem từng trang bằng pdftoppm; giấy nhỏ cố gắng gọn 1 trang.
## Chưa có / cần anh bổ sung khi dùng đến
Số CCCD, ngày cấp, nơi cấp của anh (không lưu trong kho); đường kính dấu thật; mã ngành kinh doanh; bộ mẫu Thịnh Vượng (làm khi anh bảo, theo mục 1.2 hồ sơ). Giấy phép đăng ký kinh doanh gốc (PDF) không lưu trong kho vì có số CCCD; thông tin đã có trong du-lieu/.
