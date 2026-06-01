# ============================================================
#  function/pdf_reader.py — Đọc & parse file PDF
#  Không phụ thuộc vào bất kỳ thành phần UI nào
# ============================================================

import os
import re
from pathlib import Path

import pdfplumber

from config import INVOICE_PAGE_MARKER, FIELD_PATTERNS


def find_pdf_files(target: str) -> list[str]:
    """
    Tìm và trả về danh sách đường dẫn file PDF từ một nguồn duy nhất.
    - Nếu 'target' là một file .pdf cụ thể: trả về list chứa file đó.
    - Nếu 'target' là một thư mục: quét và trả về tất cả file .pdf bên trong.
    """
    target_path = Path(target)
    if not target_path.exists():
        return []

    # Trường hợp 1: Người dùng chọn đích danh một file PDF cụ thể
    if target_path.is_file():
        if target_path.suffix.lower() == ".pdf":
            return [str(target_path.resolve())]
        return []

    # Trường hợp 2: Người dùng chọn một thư mục chứa các file PDF
    if target_path.is_dir():
        return sorted(str(p.resolve()) for p in target_path.glob("*.pdf"))

    return []
def get_specific_pdf(file_path: str) -> list[str]:
    """
    HÀNH ĐỘNG 2: Chọn cụ thể một file.
    Kiểm tra file có tồn tại và đúng định dạng PDF không.
    Trả về list chứa 1 đường dẫn duy nhất để đồng bộ dữ liệu với hàm quét thư mục.
    """
    path = Path(file_path)
    if path.exists() and path.is_file() and path.suffix.lower() == ".pdf":
        return [str(path.resolve())]
    return []


def _is_invoice_page(text: str) -> bool:
    """
    Trả về True nếu đây là trang Commercial Invoice.
    Dựa vào dòng đầu tiên của text có bắt đầu bằng INVOICE_PAGE_MARKER.
    """
    if not text:
        return False
    first_line = text.strip().split("\n")[0].strip()
    return first_line.startswith(INVOICE_PAGE_MARKER)


def _parse_fields(text: str) -> dict[str, str]:
    """
    Áp dụng các regex trong FIELD_PATTERNS lên text,
    trả về dict {tên_trường: giá_trị}.
    """
    result: dict[str, str] = {}
    for field, pattern in FIELD_PATTERNS.items():
        match = re.search(pattern, text)
        result[field] = match.group(1).strip() if match else ""
    return result


def extract_from_pdf(pdf_path: str) -> dict | None:
    """
    Mở file PDF, duyệt từng trang, tìm trang Commercial Invoice đầu tiên,
    rồi trích xuất các trường dữ liệu.

    Trả về:
        dict  — nếu tìm được trang CI và parse thành công
                {"File": ..., "Invoice Date": ..., ...}
        dict  — nếu có lỗi đọc file
                {"File": ..., "error": "..."}
        None  — nếu không có trang CI trong file
    """
    filename = os.path.basename(pdf_path)
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                if _is_invoice_page(text):
                    fields = _parse_fields(text)
                    fields["File"] = filename
                    return fields
    except Exception as exc:
        return {"File": filename, "error": str(exc)}
    return None