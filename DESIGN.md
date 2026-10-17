# Vietnamese Excel Sorter — Thiết kế Giai đoạn 1

## 1. Mục tiêu và phạm vi

Ứng dụng desktop Windows để chọn tệp `.xlsx`, trang tính, cột và vùng dữ liệu,
sau đó sắp xếp **nguyên hàng** theo thứ tự tiếng Việt. Tệp đầu vào không bị ghi
đè; kết quả luôn được lưu thành tệp mới. Nội dung ô gốc không bị chuẩn hóa lại
hoặc bỏ dấu khi ghi.

Giai đoạn này chỉ chốt thiết kế và tiêu chí nghiệm thu. Chưa có mã ứng dụng,
chưa chạy bộ kiểm thử và chưa tuyên bố kết quả sắp xếp chuẩn. (Trạng thái này
ghi nhận tại thời điểm Giai đoạn 1.)

## 2. Khảo sát môi trường

Khảo sát tại thư mục làm việc `C:\Users\ADMIN\Downloads\TIENGVIEt`:

- Python 3.14.0
- `openpyxl` 3.1.2 đã cài
- `tkinter` có sẵn
- `pytest` chưa cài
- PyICU chưa cài; do đó chưa xác minh được ICU hoặc locale `vi_VN` trong môi
  trường hiện tại
- Thư mục làm việc chưa có tệp dự án

`openpyxl` phù hợp để đọc/ghi `.xlsx` và giữ định dạng ô thông dụng, nhưng không
đảm bảo bảo toàn mọi tính năng Excel nâng cao. `tkinter` đáp ứng giao diện desktop
cơ bản. PyICU sẽ được đánh giá ở Giai đoạn 2; việc cài đặt và kiểm tra locale thực
tế là điều kiện trước khi chọn ICU làm bộ so sánh đang hoạt động.

## 3. Quy tắc so sánh đã thống nhất

1. Chuẩn hóa NFC chỉ cho khóa so sánh. Giữ nguyên giá trị gốc trong workbook;
   không bỏ dấu và không tự sửa chính tả.
2. Ưu tiên ICU Collation với locale `vi_VN` nếu thư viện hoạt động và locale
   thực tế được xác minh. Cấu hình và kết quả phải được kiểm chứng bằng bộ dữ
   liệu kiểm thử, không suy ra chỉ từ tên locale.
3. Nếu không dùng được ICU, dùng khóa dự phòng đa cấp được tài liệu hóa:
   - Mức chữ cái cơ sở theo thứ tự
     `A, Ă, Â, B, C, D, Đ, E, Ê, G, H, I, K, L, M, N, O, Ô, Ơ, P, Q, R, S, T, U, Ư, V, X, Y`.
   - So sánh toàn bộ chuỗi chữ cái cơ sở trước; chỉ khi mức này bằng nhau mới
     xét dấu thanh theo thứ tự `ngang, huyền, hỏi, ngã, sắc, nặng`.
   - Chữ hoa/thường không tạo khác biệt trong khóa chính. Những hàng bằng khóa
     giữ thứ tự ban đầu (sắp xếp ổn định).
   - Dấu câu được xử lý theo mã Unicode sau phần chữ cái theo chính sách fallback
     đã chọn; nội dung ô không bị sửa.
4. Không bỏ khoảng trắng đầu/cuối khi ghi dữ liệu; chỉ bỏ qua chúng khi tạo khóa.
   Giá trị chỉ có khoảng trắng được xem là ô trống khi sắp xếp.
5. Mặc định sắp xếp theo toàn bộ họ tên. Tùy chọn tên gọi cuối cùng lấy từ cuối
   cùng tách bởi khoảng trắng làm khóa chính và toàn bộ họ tên làm khóa phụ.
6. Cho phép chọn ô trống ở đầu hoặc cuối; mặc định ở cuối. Vị trí ô trống không
   đổi khi đảo chiều tăng/giảm.
7. Với cột trộn văn bản, số và ngày, xếp nhóm văn bản trước, tiếp theo là nhóm
   số/ngày theo giá trị, sau đó là ô trống. Hàng có khóa bằng nhau giữ nguyên
   thứ tự. Kiểu ô khác chưa có quy tắc đã thống nhất phải được báo và không tự
   ý xếp bằng một quy ước ngầm.
8. Khi dùng ICU, hành vi dấu câu và các trọng số khác theo collator `vi_VN` sẽ
   được ghi nhận riêng; không giả định rằng ICU và fallback cho cùng kết quả
   trước khi bộ kiểm thử chứng minh điều đó.

### Giới hạn về tính chuẩn

Thứ tự fallback được xác định tường minh để có kết quả tái lập, nhưng không được
quảng bá là tương đương hoàn toàn với ICU hay một chuẩn từ điển chính thức nếu
chưa đối chiếu được. Giai đoạn 2 phải thử ICU, kiểm tra locale thực tế và đối
chiếu bộ dữ liệu chuẩn. Nếu ICU không khả dụng, báo cáo và README phải nêu rõ
fallback cùng giới hạn của nó.

## 4. Xử lý workbook và an toàn dữ liệu

- Chọn worksheet, hàng tiêu đề, vùng dữ liệu và cột khóa; không đưa hàng tiêu đề
  vào sắp xếp.
- Di chuyển nguyên hàng trong phạm vi dữ liệu được chọn, giữ quan hệ giữa các cột.
- Lưu sang đường dẫn mới, không ghi đè nguồn; xác nhận trước khi ghi kết quả.
- Nếu vùng dữ liệu cần sắp xếp có ô công thức, dừng thao tác, báo worksheet/địa
  chỉ ô và yêu cầu người dùng dùng bản sao chỉ có giá trị. Công thức ở nơi khác
  vẫn cần cảnh báo vì có thể tham chiếu vùng bị di chuyển.
- Với liên kết ngoài, ô gộp giao với hàng sắp xếp, bảng dữ liệu, biểu đồ, ảnh,
  pivot, conditional formatting, data validation hoặc thành phần nâng cao có
  nguy cơ bị thay đổi: phát hiện, cảnh báo rõ và chặn thao tác; nêu tính năng
  cần xử lý trước. Không được âm thầm bỏ dữ liệu. Tệp `.xlsm` không thuộc phạm
  vi `.xlsx` và phải bị từ chối thay vì xử lý như `.xlsx`.
- Trước khi lưu, hiển thị bản xem trước và cảnh báo. Sau khi ghi, mở lại tệp
  kết quả và kiểm tra có thể đọc được, số hàng và dữ liệu hàng được bảo toàn.
  Lưu ý: xác nhận tiếp tục với tính năng nâng cao có thể đồng nghĩa chấp nhận
  rủi ro tương thích của `openpyxl`; nguồn luôn được giữ nguyên để có thể phục hồi.

## 5. Kiến trúc và tệp dự kiến

```text
unicode_utils.py          NFC và chuẩn bị dữ liệu cho khóa so sánh
vietnamese_sort.py        ICU, khóa fallback, chế độ họ tên, khóa ổn định
excel_processor.py        đọc/cảnh báo/sắp xếp nguyên hàng/lưu tệp mới
ui.py                     giao diện tkinter, xem trước, nhật ký và thông báo
tests/
  test_unicode_utils.py
  test_vietnamese_sort.py
  test_excel_processor.py
  fixtures/                workbook kiểm thử nhỏ, có thể tái tạo
requirements.txt          phụ thuộc ứng dụng và phụ thuộc kiểm thử
README.md                 cài đặt, chạy, quy tắc và giới hạn
test_report.md            lệnh đã chạy, số đạt/trượt, lỗi và giới hạn còn lại
```

## 6. Kế hoạch kiểm thử

- Unicode: NFC/NFD tương đương về khóa; giá trị gốc không đổi; dấu không bị mất.
- So sánh: toàn bộ chữ cái cơ sở, D/Đ, dấu thanh, chữ hoa/thường, khóa bằng nhau,
  tên có dấu câu/gạch nối và bộ 13 tên mẫu do yêu cầu cung cấp.
- Kiểm chứng ICU: tạo collator `vi_VN`, xác minh locale collator thực tế và đối
  chiếu kết quả với các vector kỳ vọng đã chốt. Nếu không có ICU, chạy vector
  fallback và ghi rõ không phải chứng nhận tương đương ICU.
- Excel: header và hàng nguyên vẹn; nhiều cột đi cùng nhau; tăng/giảm; vị trí ô
  trống; khoảng trắng; trùng lặp và ổn định; văn bản/số/ngày; công thức; workbook
  có cảnh báo; lưu sang tệp mới và mở lại kết quả.
- UI: chọn tệp/worksheet/cột/vùng; xem trước; cảnh báo; thông báo lỗi; lưu kết quả.
- Báo cáo kết quả chỉ được cập nhật sau khi thực sự chạy kiểm thử.

## 7. Tiêu chí nghiệm thu

- [x] Chỉ chấp nhận `.xlsx`; nguồn không bị ghi đè và kết quả lưu được vào tệp mới.
- [x] ICU `vi_VN` được xác minh bằng runtime và bộ vector, hoặc fallback được
  công bố cùng giới hạn; không tuyên bố chuẩn hoàn toàn khi chưa kiểm chứng.
- [x] Quy tắc chữ cái, dấu thanh, hoa/thường, khoảng trắng, tên gọi cuối, dấu câu,
  trống và dữ liệu hỗn hợp có kiểm thử xác định.
- [x] Mọi hàng giữ quan hệ giữa các cột; header, giá trị gốc và thứ tự ổn định
  được kiểm tra.
- [x] Công thức trong vùng sắp xếp chặn thao tác; tính năng rủi ro được phát hiện,
  cảnh báo và chặn lưu khi có khả năng mất/thay đổi dữ liệu.
- [x] Giao diện cung cấp chọn tệp/worksheet/cột/vùng, tăng/giảm, kiểu tên, vị trí
  ô trống, xem trước, sắp xếp, lưu tệp mới và nhật ký.
- [x] Toàn bộ kiểm thử đã chạy đạt; `test_report.md` phản ánh đúng kết quả thực tế.
- [x] README mô tả cài đặt, chạy, dữ liệu mẫu và giới hạn.

## 8. Kết quả Giai đoạn 2

- Tạo `unicode_utils.py`, `vietnamese_sort.py` và các kiểm thử trong `tests/`.
- Tạo `requirements.txt`; môi trường dùng `openpyxl` 3.1.2 và pytest 8.4.2.
- Runtime không có PyICU, nên backend `auto` dùng fallback và báo rõ lý do.
  Locale ICU `vi_VN` chưa thể xác minh trong môi trường này.
- Bộ mẫu 13 tên được kiểm thử theo thứ tự chữ cái cơ sở đã thống nhất; kết quả
  fallback gồm `Nguyễn Đạt` trước `Nguyễn Đăng` do `A` đứng trước `Ă`.
- Lệnh `python -m pytest -q`: 19 passed.
- Lệnh `python -m compileall -q unicode_utils.py vietnamese_sort.py tests`:
  thành công.
- Một lần chạy đầu tiên phát hiện vector kỳ vọng đặt `Đăng` trước `Đạt`; vector
  được sửa theo quy tắc `A < Ă`, sau đó toàn bộ kiểm thử đạt.

Fallback có quy tắc xác định và được kiểm thử, nhưng chưa được chứng nhận tương
đương ICU/chuẩn từ điển chính thức.

## 9. Kết quả Giai đoạn 3

- Tạo `excel_processor.py` với hàm `sort_xlsx`, nhận worksheet, cột (số thứ tự
  hoặc nhãn Excel), hàng tiêu đề và phạm vi dữ liệu; sắp xếp toàn bộ hàng từ
  trái sang cột cuối đang dùng.
- Hỗ trợ tăng/giảm, đặt ô trống đầu/cuối, họ tên đầy đủ/tên cuối, text/số/ngày.
  Nhóm dữ liệu hỗn hợp được xếp văn bản trước, số/ngày cùng nhóm theo giá trị
  serial Excel, rồi đến ô trống. Ô boolean và giá trị Excel lỗi bị từ chối.
- Không ghi đè nguồn hoặc tệp đích có sẵn. Kết quả được mở lại và kiểm tra nội
  dung/kiểu/định dạng của các ô trong hàng đã sắp xếp, liên kết và ghi chú.
- Công thức trong bất kỳ ô nào của hàng di chuyển, liên kết ngoài, ô gộp giao
  với các hàng sắp xếp, bảng, biểu đồ, ảnh, pivot, conditional formatting và
  data validation chặn thao tác. Công thức ngoài các hàng di chuyển tạo cảnh báo;
  API yêu cầu xác nhận rõ qua `acknowledge_warnings=True`.
- Workbook có tính năng openpyxl cảnh báo không hỗ trợ/bị loại bỏ cũng bị chặn.
  Cảnh báo deprecation nội bộ thư viện không bị nhầm là mất dữ liệu.
- Kiểm thử ban đầu tìm ra API external links không tồn tại trên `Workbook` công
  khai; kiểm tra được sửa dựa trên thuộc tính thực tế. Một kiểm thử khác phát
  hiện worksheet có hàng cuối trống có thể có `max_row` ngắn hơn sau khi ghi;
  xác minh đã đổi sang kiểm tra nội dung các ô thay vì chiều cao trang tính.
- Lệnh `python -m pytest -q`: **34 passed**.
- Lệnh `python -m compileall -q excel_processor.py unicode_utils.py
  vietnamese_sort.py tests`: thành công.
- Pytest phát 51 cảnh báo deprecation từ `openpyxl` 3.1.2 gọi
  `datetime.datetime.utcnow()` trên Python 3.14; không có test thất bại.

Giai đoạn 3 đạt trên workbook kiểm thử được tạo tự động; chưa thử với workbook
thực do người dùng cung cấp. Giao diện sẽ nối xác nhận cảnh báo ở Giai đoạn 4.

## 10. Kết quả Giai đoạn 4

- Tạo `ui.py` bằng tkinter: chọn tệp `.xlsx`, worksheet, cột, hàng tiêu đề,
  vùng dữ liệu, chiều sắp xếp, chế độ họ tên, vị trí ô trống và backend.
- Có bản xem trước tối đa 50 hàng; cảnh báo cần xác nhận được hiển thị trước
  khi tạo bản xem trước. Nút lưu chỉ bật khi bản xem trước thành công.
- Kết quả trung gian được đặt trong thư mục tạm và dọn khi đóng ứng dụng hoặc
  thay cấu hình. Tệp nguồn được kiểm tra chưa đổi trước khi lưu. Tệp đích được
  tạo độc quyền; tệp đã tồn tại hoặc trùng nguồn sẽ không bị ghi đè.
- Nhật ký, hộp thoại lỗi dễ hiểu và nút mở thư mục nguồn/kết quả được thêm.
- Chạy `python ui.py` để mở ứng dụng trên máy Windows có desktop.
- Lệnh `python -m pytest -q`: **36 passed**.
- Lệnh `python -m compileall -q ui.py excel_processor.py unicode_utils.py
  vietnamese_sort.py tests`: thành công.
- 51 cảnh báo deprecation của `openpyxl`/Python 3.14 vẫn còn như Giai đoạn 3.
  Bộ kiểm thử UI xác nhận format xem trước và lưu tệp không ghi đè; chưa thực
  hiện kiểm thử tương tác cửa sổ tkinter trên desktop.

Phần mã Giai đoạn 4 đã được xây dựng và các kiểm thử tự động đạt; nghiệm thu
trực quan trên desktop vẫn còn. Giai đoạn 5 chưa bắt đầu.

## 11. Kết quả Giai đoạn 5

- Bổ sung kiểm thử tích hợp cho khóa bằng nhau theo cả hai chiều, giữ hàng tiêu
  đề/footer ngoài vùng và bảo toàn hyperlink cùng hàng.
- Chạy `python -m pytest -q`: **39 passed, 0 failed**, 77 cảnh báo deprecation,
  0.66 giây.
- Chạy `python -m compileall -q ui.py excel_processor.py unicode_utils.py
  vietnamese_sort.py tests`: thành công.
- Kết quả có thể tái lập và phạm vi/giới hạn được ghi tại `test_report.md`.
- Chưa có tệp Excel thực từ người dùng để kiểm thử; UI chưa được kiểm thử tương
  tác trên desktop; PyICU/locale `vi_VN` chưa khả dụng để đối chiếu.

Giai đoạn 5 đạt cho toàn bộ kiểm thử tự động hiện có. Các giới hạn trên còn mở;
không được diễn giải kết quả này là kiểm chứng ICU hoặc nghiệm thu với workbook
thực.

## 12. Kết quả Giai đoạn 6

- Tạo `README.md` với cài đặt, chạy, quy tắc, giới hạn, kiểm thử và lệnh đóng gói.
- Tạo `examples/create_sample_workbook.py`,
  `examples/sample_vietnamese_names.xlsx` và
  `examples/expected_vietnamese_names.csv`. Kiểm thử end-to-end xác nhận bản
  sắp xếp khớp toàn bộ mã bản ghi/tên/địa danh trong CSV.
- Tạo `requirements-build.txt` và build Windows one-file bằng PyInstaller 6.16.0.
  Kết quả: `dist\VietnameseExcelSorter.exe`, 35,098,307 bytes.
- Chạy smoke test EXE: cửa sổ “Sắp xếp tiếng Việt trong Excel” xuất hiện và
  phản hồi; đóng các tiến trình sau kiểm tra. Giao diện Python trực tiếp cũng
  được mở và đóng bình thường.
- Build có cảnh báo một hidden import tùy chọn
  `charset_normalizer.md__mypyc` không tìm thấy; EXE vẫn build và mở được.
- Kiểm thử cuối: `python -m pytest -q` **40 passed, 0 failed**, 82 cảnh báo
  deprecation từ `openpyxl`/Python 3.14, 0.75 giây.
- `python -m compileall -q ui.py excel_processor.py unicode_utils.py
  vietnamese_sort.py examples tests`: thành công.
- `test_report.md` phản ánh các lệnh/kết quả thực tế. README, workbook mẫu và
  kết quả kỳ vọng được đồng bộ lên repository GitHub công khai đã chỉ định.

Các bước xây dựng/test đã hoàn thành trên workbook tổng hợp; chưa nhận được tệp
Excel thực của người dùng. Giao diện đã smoke test nhưng chưa thao tác thủ công
từng control; chưa xác minh PyICU `vi_VN`. Vì vậy đây là bản chạy được theo fallback
đã mô tả, không phải tuyên bố đạt chuẩn ICU/toàn bộ workbook Excel nâng cao.
