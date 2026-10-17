# Công cụ sắp xếp tiếng Việt trong Excel

Ứng dụng Windows dùng giao diện Tkinter để sắp xếp **nguyên hàng** trong tệp
`.xlsx` theo thứ tự tiếng Việt. Tệp nguồn không bị ghi đè; kết quả được lưu
thành tệp mới.

## Trạng thái quy tắc sắp xếp

Ứng dụng ưu tiên ICU/PyICU locale `vi_VN` nếu có và xác minh được runtime locale.
Môi trường phát triển hiện không có PyICU, vì vậy mặc định đang dùng khóa
fallback được định nghĩa trong `vietnamese_sort.py`, không phải thứ tự đã được
chứng nhận tương đương ICU hay một chuẩn từ điển chính thức.

Fallback phân biệt các chữ cái cơ sở theo thứ tự:

`A, Ă, Â, B, C, D, Đ, E, Ê, G, H, I, K, L, M, N, O, Ô, Ơ, P, Q, R, S, T, U, Ư, V, X, Y`

So sánh toàn bộ chuỗi chữ cái cơ sở trước; khi phần đó bằng nhau, dấu thanh
được xét theo thứ tự ngang, huyền, hỏi, ngã, sắc, nặng. Khóa so sánh chuẩn hóa
NFC, bỏ qua khoảng trắng đầu/cuối và không phân biệt hoa/thường. Nội dung ô
gốc không bị chuẩn hóa lại. Các khóa bằng nhau giữ thứ tự hàng ban đầu.

Mặc định sắp xếp họ tên từ trái sang phải. Tùy chọn tên gọi cuối cùng tách từ
cuối bằng khoảng trắng, rồi dùng toàn bộ họ tên làm khóa phụ. Với cột hỗn hợp,
văn bản đứng trước số/ngày (so sánh theo giá trị serial Excel), tiếp đến ô trống.

## Yêu cầu

- Windows 10/11 và Python 3.10 trở lên
- Tkinter (thường đi kèm cài đặt Python trên Windows)
- Các gói trong `requirements.txt`

Chỉ xử lý `.xlsx`. Không xử lý `.xlsm`.

## Cài đặt và chạy

Mở PowerShell tại thư mục dự án:

```powershell
python -m pip install -r requirements.txt
python ui.py
```

Trong ứng dụng:

1. Chọn workbook `.xlsx`, trang tính, cột khóa và vùng hàng dữ liệu. Hàng tiêu
   đề mặc định là hàng 1; có thể để trống nếu không có tiêu đề.
2. Chọn tăng/giảm, cách sắp xếp tên, vị trí ô trống; nhấn **Xem trước kết quả**.
3. Đọc các cảnh báo workbook. Công thức trong hàng cần di chuyển và một số tính
   năng không an toàn sẽ chặn thao tác.
4. Sau khi xem trước, nhấn **Thực hiện và lưu thành tệp mới...**. Ứng dụng không
   ghi đè tệp nguồn hoặc tệp đích đã tồn tại.

Ứng dụng hiển thị tối đa 50 hàng đầu của kết quả xem trước. Workbook trung gian
được giữ trong thư mục tạm và bị dọn khi đóng ứng dụng hoặc đổi thiết lập.

## Workbook mẫu

- Tệp đầu vào: [`examples/sample_vietnamese_names.xlsx`](examples/sample_vietnamese_names.xlsx)
- Thứ tự kỳ vọng theo **fallback đã mô tả ở trên**:
  [`examples/expected_vietnamese_names.csv`](examples/expected_vietnamese_names.csv)
- Có thể tái tạo workbook mẫu bằng:

```powershell
python examples\create_sample_workbook.py
```

Chọn trang tính `Dữ liệu`, cột B, hàng tiêu đề 1, dữ liệu từ hàng 2 đến 14,
tăng dần, toàn bộ họ tên và ô trống ở cuối. Kết quả cần khớp tệp CSV kỳ vọng
theo mã bản ghi và tên. Workbook mẫu là dữ liệu tổng hợp, không chứa dữ liệu cá
nhân thật.

## Kiểm thử

```powershell
python -m pytest -q
python -m compileall -q ui.py excel_processor.py unicode_utils.py vietnamese_sort.py tests
```

Kết quả kiểm thử gần nhất được ghi trong [`test_report.md`](test_report.md).

## Đóng gói `.exe` (tùy chọn)

Trên Windows, cài PyInstaller rồi đóng gói:

```powershell
python -m pip install -r requirements-build.txt
python -m PyInstaller --clean --onefile --windowed --name VietnameseExcelSorter ui.py
```

Tệp chạy dự kiến được tạo tại `dist\VietnameseExcelSorter.exe`. Việc build thành
công không thay thế cho kiểm thử UI thủ công trên máy Windows đích.

## Giới hạn và an toàn

- PyICU và locale ICU `vi_VN` chưa được xác minh trên môi trường hiện tại; xem
  trạng thái thực tế trong `test_report.md`.
- Công thức trong bất kỳ ô nào thuộc hàng sẽ di chuyển chặn sắp xếp. Công thức
  ngoài vùng, defined names và một số tính năng workbook yêu cầu cảnh báo/xác nhận.
- Workbook có liên kết ngoài, ô gộp giao với vùng hàng, bảng, biểu đồ, ảnh,
  pivot, conditional formatting, data validation hoặc extension không được
  `openpyxl` hỗ trợ sẽ bị chặn để tránh mất/thay đổi dữ liệu.
- Boolean, ô lỗi Excel và kiểu giá trị không được hỗ trợ trong cột khóa sẽ bị từ chối.
- `openpyxl` giữ định dạng ô thông dụng nhưng không hỗ trợ mọi thành phần Excel.
  Hãy giữ bản gốc và kiểm tra workbook đầu ra trước khi sử dụng.
- Bộ kiểm thử dùng workbook tổng hợp; chưa kiểm thử trên workbook thật của người dùng.
