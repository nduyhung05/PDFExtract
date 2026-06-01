from pathlib import Path

# ── Đường dẫn mặc định ──────────────────────────────────────
DEFAULT_PDF_FOLDER = r"\\stffnp03\SFDATA_2024\28. SHP\1. EXPORT\08. UA\SHIPPING DOC"
DEFAULT_OUTPUT_FOLDER = str(Path.home() / "Desktop")
OUTPUT_CSV_NAME  = "invoice_data.csv"
OUTPUT_XLSX_NAME = "invoice_data.xlsx"

# ── Nhận dạng trang ─────────────────────────────────────────
INVOICE_PAGE_MARKER  = "Commercial Invoice"   # trang cần đọc
SKIP_PAGE_MARKER     = "Packing List"         # trang bỏ qua

# ── Trường dữ liệu cần lấy (tên hiển thị → regex) ───────────
FIELD_PATTERNS: dict[str, str] = {
    "Invoice Date":       r"Invoice Date\s+([\d/]+)",
    "Invoice Number":     r"Invoice Number\s+(\S+)",
    "Shipment ID / ASN":  r"Shipment ID\s*/\s*ASN\s+(\S+)",
}

# ── Cột xuất ra (thứ tự cố định) ────────────────────────────
OUTPUT_COLUMNS = ["File", "Invoice Date", "Invoice Number", "Shipment ID / ASN"]

# ── Giao diện Excel ─────────────────────────────────────────
EXCEL_HEADER_COLOR = "1F4E79"   # màu nền header (hex, không có #)
EXCEL_ROW_EVEN     = "DCE6F1"   # màu dòng chẵn
EXCEL_ROW_ODD      = "FFFFFF"   # màu dòng lẻ
EXCEL_COL_WIDTHS   = [35, 18, 22, 22]   # độ rộng cột theo OUTPUT_COLUMNS

# ── Giao diện GUI ────────────────────────────────────────────
APP_TITLE       = "Invoice Extractor - UA Shipping Doc"
APP_GEOMETRY    = "800x600"
COLOR_PRIMARY   = "#1F4E79"
COLOR_BTN_BLUE  = "#2E75B6"
COLOR_BTN_GREEN = "#1F7A3E"
COLOR_BTN_ORANGE= "#C55A11"
COLOR_BTN_GRAY  = "#767171"
COLOR_BG        = "#F0F4F8"
COLOR_ROW_EVEN  = "#DCE6F1"
COLOR_ROW_ERROR = "#FFD7D7"