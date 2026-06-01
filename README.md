# 📄 PDF Invoice Extractor (Công Cụ Trích Xuất Hóa Đơn PDF)

![Python](https://shields.io)
![License](https://shields.io)

Ứng dụng chạy trên máy tính (PC Desktop App) giúp tự động quét, đọc và trích xuất các trường dữ liệu quan trọng từ các file hóa đơn thương mại dạng PDF (Commercial Invoice) sang định dạng cấu trúc dữ liệu mong muốn.

## ✨ Tính năng nổi bật
- 🔍 **Tìm kiếm thông minh**: Hỗ trợ chọn quét toàn bộ thư mục hoặc chọn đích danh 1 file PDF cụ thể tại cùng một giao diện.
- ⚙️ **Nhận diện chính xác**: Tự động lọc đúng trang chứa `Commercial Invoice` dựa trên dấu hiệu dòng đầu tiên.
- 📊 **Trích xuất dữ liệu**: Sử dụng các biểu thức chính quy (Regex) linh hoạt để bóc tách thông tin cần thiết.

## 📁 Cấu trúc thư mục dự án
```text
├── function/
│   └── exporter.py     # Xử lý chuyển hóa file sang định dạng XLSX
│   └── pdf_reader.py   # Đọc và phân loại trang rồi lấy thông tin cần thiết
├── ui/
│   └── main.py         # Giao diện chính của ứng dụng (Tkinter UI)
├── config.py           # Lưu trữ các mẫu Regex (FIELD_PATTERNS) và Marker nhận diện
├── main.py             # Khởi động ứng dụng
└── README.md           # Tài liệu hướng dẫn dự án (File này)
```

## 🚀 Hướng dẫn cài đặt & Chạy ứng dụng

### 1. Yêu cầu hệ thống
Máy tính đã cài đặt sẵn **Python 3.10** trở lên.

### 2. Các bước cài đặt
Cài đặt các thư viện phụ thuộc cần thiết bằng cách chạy lệnh sau tại thư mục dự án:
```bash
pip install -r requirements.txt
```
*(Nếu bạn sử dụng tính năng kéo thả nâng cao, cài đặt thêm: `pip install tkinterdnd2`)*

### 3. Khởi chạy ứng dụng
Chạy lệnh sau để mở giao diện phần mềm:
```bash
python main.py
```

## 🛠️ Cài đặt ứng dụng
Đang cập nhật!

## 📝 Giấy phép (License)
Dự án này được phân phối dưới giấy phép **MIT License**. Bạn có thể tự do chỉnh sửa và sử dụng cho mục đích cá nhân hoặc thương mại.
