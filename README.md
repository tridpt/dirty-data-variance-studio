# 🧼📉 DataScrub & Variance Studio (Port 4195)

**Trạm Làm Sạch Dữ Liệu Bẩn Đặc Thù Việt Nam & Tự Động Giải Phẫu Biến Động Doanh Số / Chỉ Số (DA Practical Weapon)**

Giải quyết trọn vẹn 100% quy trình thực chiến hàng ngày của một Data Analyst tại doanh nghiệp (Ngân hàng, Logistics, Sản xuất, Nhân sự, Bán lẻ):
- **70% thời gian đầu — Lọc dữ liệu bẩn từ các phòng ban**:
  - **Chuẩn hóa Ngày tháng lệch định dạng**: Tự động gộp `dd/mm/yyyy` trộn `mm/dd/yyyy`, tự suy luận ngữ cảnh tháng để cứu các dòng ngày chuẩn Mỹ nhập nhằng (`10-04-2026` nằm giữa báo cáo tháng 10), đọc mã số Excel Serial (`46308`) & ngày chữ tiếng Việt.
  - **Chuẩn hóa SĐT Viễn thông Việt Nam**: Xử lý `+84`, khôi phục số `0` đầu bị mất do Excel, chuyển đổi thuê bao **11 số cũ sang 10 số chuẩn** (`016x ➔ 03x`, `012x ➔ 07x/08x`).
  - **Đọc Tiền tệ & Số liệu kiểu Việt Nam**: `"12.500.000 đ"`, `"8,4 tr"`, `"1,25 tỷ"`, `"(1.250.000)"`, khử lỗi công thức `#DIV/0!`, `#REF!`.
  - **Gom nhóm Tên viết sai / viết tắt (Fuzzy Matching + Từ điển VN)**: Gom `"TP.HCM"`, `"Sài Gòn"`, `"SG"` về `"TP. Hồ Chí Minh"`, chuẩn hóa Unicode Tổ hợp (NFD) về Dựng sẵn (NFC), có chốt chặn chống gộp nhầm tên riêng người Việt (`Nguyễn Văn Hùng` vs `Nguyễn Văn Dũng`) và khác biệt kích cỡ/tải trọng (`5T` vs `10T`).
  - **Bắt dòng Trùng lặp ngầm & Giá trị Ngoại lai theo Ngữ cảnh Mặt hàng**: Phát hiện bản ghi nhập đúp khác cách gõ chữ và kiểm tra chéo Đơn giá ngầm định (`Giá trị / Số lượng`) để bắt lỗi gõ dư số `0` mà không phạt nhầm đơn sỉ lớn.
- **30% thời gian quyết định — Trả lời câu hỏi của Sếp ("Sao kỳ này tăng/giảm vậy em?")**:
  - **Price–Volume–Mix (PVM) Waterfall Bridge**: Bóc tách chính xác tuyệt đối (`Sai số dư = 0`) biến động giữa 2 kỳ bất kỳ thành **Sản lượng (Volume)**, **Cơ cấu mặt hàng (Mix)**, **Đơn giá / Chiết khấu (Price)** và **Mới / Ngưng phát sinh**.
  - **Key-Driver Đa Chiều & Giải Phẫu Chéo 2 Tầng**: Chỉ đích danh Top 3 Chi nhánh / Nhân viên / Kênh bán kéo chỉ số đi xuống (hoặc gồng lên) nhiều nhất; bấm vào bất kỳ đơn vị nào để mổ xẻ PVM nội bộ của đơn vị đó.
  - **Kiểm toán Bẫy Báo Cáo Bẩn**: Chỉ rõ nếu chạy Pivot trực tiếp trên file bẩn gốc thì số liệu báo cáo sếp sẽ bị sai lệch bao nhiêu triệu đồng.
- **SQLite Analytical Workbench + Recipe 1-Chạm + Xuất Excel 4-Sheet (`.XLSX`)**:
  - Tích hợp động cơ `node:sqlite` với 4 bảng (`clean_data`, `pvm_summary`, `audit_log`, `raw_data`) hỗ trợ đầy đủ `WITH ... AS` và Window Functions (`LAG()`, `SUM() OVER`).
  - Lưu cấu hình thành **Recipe** để tuần/tháng sau chỉ cần kéo-thả file mới là có ngay Workbook `.XLSX` 4 Sheet gửi sếp sau 3 giây.

## 🚀 Khởi chạy nhanh

```bash
npm start
# Mở trình duyệt tại: http://localhost:4195
```

## 🧪 Chạy bộ kiểm thử tự động

```bash
npm test
```
