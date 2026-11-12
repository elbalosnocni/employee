# Employee Sync Secure v2 – Hướng dẫn triển khai

## 1. Mục tiêu và giới hạn bảo mật

Bộ nâng cấp này giữ lại luồng Excel/VBA → Google Apps Script → Google Sheets, đồng thời hỗ trợ nhiều chi nhánh có cấu trúc Excel khác nhau.

**Sheet `Employees` chỉ nhận 7 trường:** mã chi nhánh, mã nhân viên, họ tên, mã chấm công, phòng ban, bộ phận, chức vụ (kèm thời điểm cập nhật).

**Không gửi lên API và không lưu vào `Employees`:** CCCD, ngày sinh, lương cơ bản, phụ cấp, lương hợp đồng hoặc bất kỳ trường lương nào. CCCD có thể được dùng tùy chọn để kiểm tra đủ 12 chữ số ngay trên máy Excel; giá trị CCCD không được đưa vào JSON.

Lưu ý chính xác: phiên bản này giảm thiểu dữ liệu bằng cách không thu thập/gửi trường nhạy cảm. Nó không mã hóa CCCD/lương tại nguồn. Dữ liệu CCCD/lương hiện có trong file nguồn hoặc Sheet cũ vẫn cần được bảo vệ bằng quyền truy cập riêng và quy trình di chuyển/xóa an toàn.

## 2. Thành phần

- `GAS/Code_employees.gs`: API tiếp nhận và upsert danh bạ nhân viên.
- `Excel/EmployeeSync.bas`: module VBA nhiều chi nhánh.
- `Excel/EmployeeSync.xlsm`: file điều khiển gốc được giữ nguyên để bảo toàn dự án VBA; cần import module BAS mới.
- `Excel/BranchConfig.xlsx`: cấu hình từng nguồn Excel.
- `AutoSync/RunEmployeeSync.vbs`: chạy file điều khiển cùng thư mục.
- `README.md`: hướng dẫn này.

## 3. Tạo Google Sheet trung tâm

1. Tạo một Google Spreadsheet mới hoặc chọn spreadsheet trung tâm dành riêng cho danh bạ nhân viên.
2. Không dùng lại Sheet đang chứa CCCD/lương để làm Sheet chia sẻ công khai.
3. Mở Extensions → Apps Script.
4. Thay nội dung file mặc định bằng `GAS/Code_employees.gs`.
5. Trong Project Settings → Script Properties, hoặc chạy các hàm cấu hình ở bước sau.

### Cấu hình Script Properties

Có thể chạy các hàm sau từ Apps Script Editor (thay giá trị thật):

```javascript
setEmployeeSpreadsheetId('ID_GOOGLE_SHEET_TRUNG_TAM');
setBranchToken('BH', 'TOKEN_BH_DAI_NGAU_NHIEN_TREN_24_KY_TU');
setBranchToken('HN', 'TOKEN_HN_DAI_NGAU_NHIEN_TREN_24_KY_TU');
setupEmployeeSheets();
```

Mỗi token nên là chuỗi ngẫu nhiên tối thiểu 32 ký tự. Token nhập trong BranchConfig.xlsx phải khớp token đã đăng ký cho mã chi nhánh. Script chỉ lưu SHA-256 của token, không lưu token gốc.

## 4. Triển khai Web App

1. Deploy → New deployment → Web app.
2. Execute as: Me.
3. Who has access: Anyone (nếu không, Excel ngoài Google không thể gọi endpoint).
4. Deploy và sao chép URL kết thúc bằng `/exec`.
5. Mỗi lần thay đổi code, cần cập nhật deployment (Manage deployments → Edit → New version).

Endpoint có token riêng theo chi nhánh, kiểm tra timestamp 10 phút và chỉ cho phép schema version 2. HTTPS bảo vệ đường truyền. Không đưa URL/token lên nơi công khai.

## 5. Cấu hình BranchConfig.xlsx

Mỗi dòng là một chi nhánh. Cột:
- A `Active`: YES/NO.
- B `BranchCode`: mã duy nhất, chỉ A-Z, 0-9, `_`, `-`.
- C `FolderPath`: thư mục chứa file.
- D `FilePattern`: ví dụ `DSCNV-*.xlsb`; macro chọn file mới sửa gần nhất.
- E `ExcelPassword`: mật khẩu mở file (để trống nếu không có).
- F `SheetName`: tên sheet nguồn.
- G `HeaderRow`: dòng tiêu đề (tham khảo, dùng để ghi cấu hình).
- H `FirstDataRow`: dòng bắt đầu dữ liệu.
- I `NameCol`: cột họ tên.
- J `EmployeeCodeCol`: cột mã NV.
- K `TimeClockCol`: cột mã chấm công.
- L `DepartmentCol`: cột phòng ban.
- M `SectionCol`: cột bộ phận.
- N `PositionCol`: cột chức vụ.
- O `CCCDCol`: cột CCCD tùy chọn, chỉ dùng lọc 12 chữ số tại máy nguồn; để trống nếu không cần lọc.
- P `ApiUrl`: URL Web App.
- Q `BranchToken`: token riêng chi nhánh.

Có thể nhập ký tự cột Excel (A, B, AF...) hoặc số cột (1, 2, 32...). Timestamp được lấy theo UTC của Windows để khớp kiểm tra thời hạn phía máy chủ. Cột dữ liệu khác nhau giữa các chi nhánh chỉ cần sửa cấu hình, không sửa VBA.

Dòng mẫu đã thiết lập theo mapping trong VBA gốc: tên B, mã NV D, mã chấm công E, CCCD H, phòng ban AF, bộ phận AG, chức vụ AH; dữ liệu bắt đầu dòng 8. Hãy xác nhận mapping thực tế của từng file chi nhánh trước khi chạy.

## 6. Cài VBA

1. Giữ nguyên cấu trúc thư mục trong ZIP: `Excel` chứa XLSM/BAS/BranchConfig.xlsx; `AutoSync` chứa VBS. VBS tự tìm workbook trong thư mục `Excel` cùng cấp.
2. Mở `EmployeeSync.xlsm`.
3. Nhấn Alt+F11 → File → Import File → chọn `EmployeeSync.bas`.
4. Nếu có module cũ cùng tên hàm `CopySpecificColumnsToGoogleSheets`, hãy xóa/đổi tên module cũ để tránh trùng tên macro.
5. Lưu file `.xlsm` dưới định dạng Excel Macro-Enabled Workbook.
6. Cho phép macro theo chính sách bảo mật của công ty.

Không bật chế độ lưu mật khẩu token trong nơi chia sẻ công khai. File BranchConfig.xlsx chứa token và mật khẩu Excel, cần giới hạn quyền NTFS/chia sẻ mạng.

## 7. Chạy thử an toàn

1. Ban đầu chỉ bật `YES` cho một chi nhánh thử nghiệm.
2. Kiểm tra đường dẫn, tên file, sheet, cột và token.
3. Chạy macro `CopySpecificColumnsToGoogleSheets`.
4. Kiểm tra Sheet `Employees`: chỉ có 7 trường nghiệp vụ + UpdatedAt.
5. Kiểm tra `SyncLog` trong Google Sheet và `SyncRunLog` trong file điều khiển.
6. Sau khi xác nhận đúng mới bật các chi nhánh còn lại.

Macro hiện upsert theo khóa `BranchCode + EmployeeCode`. Nếu cùng một nhân viên chuyển chi nhánh, hệ thống coi là hai bản ghi khác nhau; cần thống nhất mã nhân viên toàn công ty hoặc bổ sung quy tắc chuyển chi nhánh trước khi dùng trường hợp này.

## 8. Chia sẻ qua IMPORTRANGE

Chỉ chia sẻ sheet `Employees` (hoặc một spreadsheet trung gian chỉ chứa dữ liệu công khai). Ví dụ:

```excel
=IMPORTRANGE("URL_SPREADSHEET_TRUNG_TAM"; "Employees!A:H")
```

Không cấp quyền truy cập spreadsheet riêng tư đang chứa CCCD/lương. `IMPORTRANGE` không phải cơ chế phân quyền theo cột: người được cấp quyền vào spreadsheet nguồn có thể có khả năng xem dữ liệu trong nguồn. Vì vậy hãy dùng một spreadsheet công khai riêng chỉ chứa danh bạ được phép chia sẻ.

## 9. Các giới hạn cần biết

- Bản này chưa tự động dò header theo tên; mapping mỗi chi nhánh được cấu hình bằng chữ/số cột. Trường `HeaderRow` được giữ trong cấu hình để tiện mở rộng, còn dòng dữ liệu bắt đầu theo `FirstDataRow`.
- Mã VBA cần được import vào file XLSM bằng Excel trên Windows; môi trường bàn giao không thực thi Excel/VBA để kiểm thử end-to-end.
- Nếu nhân viên có mã nhân viên trùng giữa hai chi nhánh, khóa hiện tại phân biệt bằng BranchCode.
- Dữ liệu cũ ở Sheet hiện tại không tự động bị xóa/chuyển. Hãy tạo spreadsheet trung tâm mới chỉ chứa danh bạ công khai để tránh vô tình chia sẻ dữ liệu CCCD/lương lịch sử.
- Apps Script Web App dùng token theo chi nhánh. Token lưu trong BranchConfig.xlsx nên file này phải được bảo vệ như thông tin đăng nhập.

## 10. Checklist nghiệm thu

- [ ] API triển khai đúng URL `/exec`.
- [ ] Đã cấu hình ID spreadsheet và token từng chi nhánh.
- [ ] Mỗi chi nhánh có mapping đúng.
- [ ] Dữ liệu `Employees` không có CCCD, ngày sinh, lương/phụ cấp.
- [ ] Mã nhân viên trùng trong cùng chi nhánh được cập nhật, không nhân bản.
- [ ] Chi nhánh khác không ghi đè bản ghi của nhau.
- [ ] `SyncLog` và `SyncRunLog` ghi nhận kết quả.
- [ ] Spreadsheet được dùng làm nguồn IMPORTRANGE chỉ chứa dữ liệu được phép chia sẻ.
