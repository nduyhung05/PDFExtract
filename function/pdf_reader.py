# ============================================================
#  function/pdf_reader.py — Đọc & parse file PDF
# ============================================================

import os
import re
from pathlib import Path

import pdfplumber

from config import FIELD_PATTERNS, INVOICE_PAGE_MARKER


def find_pdf_files(folder: str) -> list[str]:
    """Tất cả file .pdf trong thư mục (không đệ quy)."""
    p = Path(folder)
    if not p.exists():
        return []
    return sorted(str(f) for f in p.glob("*.pdf"))


def search_pdf_by_names(folder: str, names: list[str]) -> dict[str, list[str]]:
    """
    Tìm file PDF trong folder khớp với danh sách tên/từ khóa.
    Hỗ trợ tên chính xác và từ khóa một phần (case-insensitive).
    Trả về { tên_tìm: [đường_dẫn_file, ...] }
    """
    all_pdfs = find_pdf_files(folder)
    results: dict[str, list[str]] = {}
    for name in names:
        name = name.strip()
        if not name:
            continue
        keyword = name.lower().removesuffix(".pdf")
        matched = [
            p for p in all_pdfs
            if keyword in os.path.basename(p).lower().removesuffix(".pdf")
        ]
        results[name] = matched
    return results


def _is_invoice_page(text: str) -> bool:
    if not text:
        return False
    return text.strip().split("\n")[0].strip().startswith(INVOICE_PAGE_MARKER)


def _parse_fields(text: str) -> dict[str, str]:
    result: dict[str, str] = {}
    for field, pattern in FIELD_PATTERNS.items():
        m = re.search(pattern, text)
        result[field] = m.group(1).strip() if m else ""
    return result


def extract_from_pdf(pdf_path: str) -> dict | None:
    """
    Trả về dict các trường từ trang Commercial Invoice,
    {"File": ..., "error": ...} nếu lỗi, hoặc None nếu không có trang CI.
    """
    filename = os.path.basename(pdf_path)
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                if _is_invoice_page(text):
                    fields = _parse_fields(text)
                    fields["File"] = filename
                    fields["_path"] = pdf_path
                    return fields
    except Exception as exc:
        return {"File": filename, "error": str(exc), "_path": pdf_path}
    return None