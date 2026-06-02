# PDF Extract v3

Ứng dụng Python đọc file PDF Shipping Doc (Under Armour), trích xuất thông tin từ **Commercial Invoice**, và xuất kết quả ra file **CSV** và **Excel**.

---

## Tính năng

- Chọn nhiều file PDF cùng lúc qua dialog (Ctrl+Click / Shift+Click)
- Tìm file theo tên hoặc từ khóa một phần (nhập nhiều dòng cùng lúc)
- Hỏi người dùng khi thêm file bị trùng tên (Thay thế / Bỏ qua / Dừng)
- Bảng kết quả tích lũy — thêm và xử lý nhiều đợt trước khi xuất
- Tự động bỏ qua trang Packing List, chỉ đọc trang Commercial Invoice
- Xuất đồng thời file `.csv` và `.xlsx` với định dạng header màu

---

## Dữ liệu được trích xuất

| Trường | Ví dụ |
|---|---|
| File | ASN1034657.pdf |
| Invoice Date | 05/28/2026 |
| Invoice Number | 5VS12625066 |
| Shipment ID / ASN | ASN1034657 |

---

## Cấu trúc dự án

```
invoice_extractor_v3/
├── main.py                  ← Điểm khởi chạy duy nhất
├── config.py                ← Toàn bộ cấu hình (đường dẫn, màu, regex)
├── requirements.txt         ← Danh sách thư viện
├── build.bat                ← Script đóng gói thành .exe (Windows)
├── backend/
│   ├── pdf_reader.py        ← Đọc PDF, parse fields, tìm file theo tên
│   └── exporter.py          ← Xuất CSV và Excel
└── ui/
    └── app.py               ← Toàn bộ giao diện Tkinter
```

---

## Cài đặt & Chạy

### Yêu cầu

- Python 3.10 trở lên
- Windows (do đường dẫn mạng `\\server\...`)

### Bước 1 — Cài thư viện

```
pip install pdfplumber openpyxl
```

### Bước 2 — Chạy ứng dụng

```
python main.py
```

---

## Hướng dẫn sử dụng

### Giao diện gồm 3 khu vực chính

```
┌─────────────────────────────┬──────────────────────────┐
│  Panel trái                 │  Panel phải              │
│  Danh sách file chờ xử lý  │  Tìm file theo tên / ASN │
└─────────────────────────────┴──────────────────────────┘
┌────────────────── Bảng kết quả đã xử lý ───────────────┐
│  File | Invoice Date | Invoice Number | ASN | Status    │
└─────────────────────────────────────────────────────────┘
[ Xử lý queue ]  [ Xuất Excel ]          [ Lưu tại: ... ]
```

---

### Cách 1 — Chọn file thủ công

1. Nhấn **+ Chọn file** ở panel trái
2. Dialog mở tại thư mục mặc định — giữ `Ctrl` để chọn nhiều file
3. File được thêm vào danh sách queue
4. Nhấn **Xử lý file trong queue**
5. Kết quả hiện ra trong bảng phía dưới

### Cách 2 — Tìm theo tên / ASN

1. Nhập tên hoặc từ khóa vào ô nhập liệu bên phải — mỗi dòng một cái

   ```
   ASN1034657
   ASN1034658
   1034659
   ```

2. Nhấn **Tìm & thêm vào queue**

   - Khớp chính xác 1 file → tự động thêm
   - Khớp nhiều file → popup cho chọn
   - Không tìm thấy → hiển thị tên bị thiếu

3. Nhấn **Xử lý file trong queue**

---

### Xử lý file trùng tên

Khi thêm file đã có trong queue, app sẽ hỏi:

| Lựa chọn | Kết quả |
|---|---|
| **Yes** | Thay thế file cũ bằng file mới |
| **No** | Giữ file cũ, bỏ qua file mới |
| **Cancel** | Dừng toàn bộ thao tác thêm |

---

### Thêm nhiều đợt

Queue tự xóa sau mỗi lần xử lý, nhưng **bảng kết quả được giữ nguyên**. Có thể:

1. Thêm đợt 1 → Xử lý → xem kết quả
2. Thêm đợt 2 → Xử lý → kết quả cộng dồn vào bảng
3. Nhấn **Xuất Excel** để xuất toàn bộ một lần

---

### Xuất kết quả

Nhấn **Xuất Excel** — app tạo đồng thời 2 file tại thư mục đã chọn:

| File | Mô tả |
|---|---|
| `invoice_data.csv` | CSV UTF-8, mở được bằng Excel |
| `invoice_data.xlsx` | Excel có định dạng header màu, dòng xen kẽ |

> Chỉ các dòng có status **OK** mới được xuất. Dòng lỗi bị bỏ qua.

---

## Cấu hình

Toàn bộ cài đặt nằm trong **`config.py`** — không cần sửa code chính:

| Biến | Mô tả |
|---|---|
| `DEFAULT_PDF_FOLDER` | Đường dẫn thư mục PDF mặc định |
| `DEFAULT_OUTPUT_FOLDER` | Thư mục lưu kết quả mặc định |
| `FIELD_PATTERNS` | Regex để tìm các trường trong PDF |
| `OUTPUT_COLUMNS` | Thứ tự cột khi xuất |
| `EXCEL_HEADER_COLOR` | Màu header Excel (mã hex) |

Ví dụ thêm trường mới vào `config.py`:

```python
FIELD_PATTERNS: dict[str, str] = {
    "Invoice Date":      r"Invoice Date\s+([\d/]+)",
    "Invoice Number":    r"Invoice Number\s+(\S+)",
    "Shipment ID / ASN": r"Shipment ID\s*/\s*ASN\s+(\S+)",
    "Total Cartons":     r"Total Carton Quantity\s+([\d]+)",   # thêm mới
}

OUTPUT_COLUMNS = ["File", "Invoice Date", "Invoice Number", "Shipment ID / ASN", "Total Cartons"]
```

---

## Đóng gói thành file .exe

Để chạy trên máy không có Python:

### Cách 1 — Script tự động

```
Double-click build.bat
```

### Cách 2 — Tự chạy lệnh

```
pip install pyinstaller
pyinstaller --onefile --windowed --name "PDFExtract" --add-data "config.py;." main.py
```

File `.exe` xuất hiện trong thư mục `dist\PDFExtract.exe`.

> **Lưu ý:** Lần đầu chạy `.exe` có thể mất 5–10 giây để load.

---

## Yêu cầu thư viện

| Thư viện | Phiên bản tối thiểu | Mục đích |
|---|---|---|
| `pdfplumber` | 0.10.0 | Đọc và trích xuất text từ PDF |
| `openpyxl` | 3.1.0 | Tạo và định dạng file Excel |
| `tkinter` | (built-in) | Giao diện đồ họa |