# ============================================================
#  config.py — Cấu hình toàn bộ ứng dụng
# ============================================================

from pathlib import Path

DEFAULT_PDF_FOLDER    = r"\\stffnp03\SFDATA_2024\28. SHP\1. EXPORT\08. UA\SHIPPING DOC"
DEFAULT_OUTPUT_FOLDER = str(Path.home() / "Desktop")
OUTPUT_CSV_NAME       = "invoice_data.csv"
OUTPUT_XLSX_NAME      = "invoice_data.xlsx"

INVOICE_PAGE_MARKER = "Commercial Invoice"
SKIP_PAGE_MARKER    = "Packing List"

FIELD_PATTERNS: dict[str, str] = {
    "Invoice Date":      r"Invoice Date\s+([\d/]+)",
    "Invoice Number":    r"Invoice Number\s+(\S+)",
    "Shipment ID / ASN": r"Shipment ID\s*/\s*ASN\s+(\S+)",
}

OUTPUT_COLUMNS = ["File", "Invoice Date", "Invoice Number", "Shipment ID / ASN"]

EXCEL_HEADER_COLOR = "1F4E79"
EXCEL_ROW_EVEN     = "DCE6F1"
EXCEL_ROW_ODD      = "FFFFFF"
EXCEL_COL_WIDTHS   = [35, 18, 22, 22]

APP_TITLE        = "Invoice Extractor - UA Shipping Doc"
APP_GEOMETRY     = "960x720"
COLOR_PRIMARY    = "#1F4E79"
COLOR_BTN_BLUE   = "#2E75B6"
COLOR_BTN_GREEN  = "#1F7A3E"
COLOR_BTN_ORANGE = "#C55A11"
COLOR_BTN_RED    = "#C0392B"
COLOR_BTN_GRAY   = "#767171"
COLOR_BG         = "#F0F4F8"
COLOR_ROW_EVEN   = "#DCE6F1"
COLOR_ROW_ERROR  = "#FFD7D7"
COLOR_PANEL_BG   = "#E8EEF5"