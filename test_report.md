# Báo cáo kiểm thử — Bộ sắp xếp tiếng Việt trong Excel

## Lần chạy xác nhận

- Ngày: 2026-10-17
- Python: 3.14.0 (64-bit, Windows)
- `openpyxl`: 3.1.2
- `pytest`: 8.4.2
- PyICU: chưa cài; locale ICU `vi_VN` chưa được xác minh

## Lệnh và kết quả

| Lệnh | Kết quả |
|---|---|
| `python -m pytest -q` | **39 passed, 0 failed**, 77 cảnh báo deprecation, 0.66 giây |
| `python -m compileall -q ui.py excel_processor.py unicode_utils.py vietnamese_sort.py tests` | Thành công |

Các cảnh báo pytest đến từ `openpyxl` 3.1.2 gọi `datetime.datetime.utcnow()`,
API đã deprecate trong Python 3.14. Không có cảnh báo nào là lỗi kiểm thử.

## Phạm vi đã kiểm tra

- Unicode NFC/NFD, giữ nguyên chuỗi gốc và khóa dự phòng chữ cái/dấu thanh.
- Bộ tên mẫu, họ tên đầy đủ/tên gọi cuối, chữ hoa/thường, dấu câu, khoảng trắng,
  giá trị trùng và thứ tự ổn định.
- Excel `.xlsx`: sắp xếp nguyên hàng, hàng/cột ngoài phạm vi, header/footer,
  chiều tăng/giảm, ô trống, dữ liệu text/số/ngày.
- Bảo toàn nội dung, kiểu dữ liệu, định dạng, chiều cao hàng, độ rộng cột,
  comment và hyperlink; workbook nguồn không bị thay đổi; không ghi đè tệp đích.
- Chặn công thức trong các hàng di chuyển, ô gộp, bảng, và kiểu dữ liệu/tính năng
  không hỗ trợ đã liệt kê; yêu cầu xác nhận với công thức ngoài vùng.
- UI helpers kiểm tra định dạng giá trị xem trước và thao tác lưu tệp không ghi đè.

## Giới hạn chưa kiểm chứng

- Chưa có workbook thật do người dùng cung cấp; toàn bộ workbook tích hợp được
  tạo trong test bằng `openpyxl`.
- Chưa chạy kiểm thử tương tác cửa sổ tkinter trên desktop; UI helper/import được
  kiểm thử tự động.
- Chưa kiểm chứng đối chiếu ICU `vi_VN`; kết quả hiện tại dùng fallback và không
  được khẳng định là tương đương ICU/chuẩn từ điển chính thức.
- Chưa đóng gói `.exe`.
