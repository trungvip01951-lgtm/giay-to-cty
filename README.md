# Giấy tờ công ty Trung Kiên

Kho riêng tư (private) lưu thông tin công ty và mẫu giấy tờ của CÔNG TY TNHH THƯƠNG MẠI VÀ CÔNG NGHIỆP TRUNG KIÊN (MST 0109546164).

| Thư mục / tệp | Nội dung |
|---|---|
| `du-lieu/cong_ty_trung_kien.json` | Nơi gốc duy nhất của thông tin công ty (đối chiếu GPKD thay đổi lần 3, 09/06/2026) và danh sách việc cần anh chốt |
| `ho-so/HoSo_ThongTinCongTy_VaMauGiayTo_TrungKien.md` | Hồ sơ thông tin và nội dung 12 mẫu đã làm (bản đã sửa 9 mục ngày 23/10/2026) |
| `mau-giay-to/*.docx` | 8 mẫu mới: thanh lý hợp đồng, nghiệm thu/chạy thử, giấy ủy quyền, giấy giới thiệu, phụ lục hợp đồng, đối chiếu công nợ, đơn đặt hàng nhà cung cấp, nhận hàng nhập khẩu |
| `cong-cu/tao_mau.py` | Tạo lại 8 mẫu từ tệp JSON: `python3 cong-cu/tao_mau.py` |
| `dau-chu-ky/` | Ảnh dấu tròn, chữ ký, dấu chức danh (PNG nền trong suốt) và bản ghép. Chỉ trong kho riêng tư |
| `giay-to-khach-hang/FATA/` | Hợp đồng gốc, phụ lục, đề nghị thanh toán (bản không dấu và có dấu) của FATA |
| `ho-so/goc/` | Hồ sơ năng lực Trung Kiên 2026 (Word) |
| `cong-no/CongNo_TrungKien.xlsx` | Bảng công nợ: đơn hàng, mã hàng, giao hàng, xuất hóa đơn, thu tiền, công nợ từng khách |
| `cong-no/CongNo_ThuLien.xlsx` | Đơn hàng và công nợ riêng Công ty TNHH vật tư thiết bị Thu Liên: mỗi đơn một sheet, trang TongQuan có thẻ số liệu, màu cảnh báo, biểu đồ |
| `CLAUDE.md`, `HUONG_DAN_PHIEN_MOI.md`, `requirements.txt` | Hướng dẫn cho phiên làm việc mới |
| `QUY_TAC.md`, `SO_LOI.md`, `NHAT_KY.md`, `KHOA_VIEC.md` | Quy tắc dự án, sổ lỗi, nhật ký thay đổi, sổ khoá việc |

Cần sửa thông tin công ty: sửa trong `du-lieu/cong_ty_trung_kien.json` rồi chạy lại `tao_mau.py`.
