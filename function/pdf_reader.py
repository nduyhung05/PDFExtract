# ============================================================
#  backend/pdf_reader.py — Đọc & parse file PDF
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


def _priority_score(path: str) -> tuple[int, int]:
    """
    Tính điểm ưu tiên để chọn file khi có nhiều kết quả trùng.

    Quy tắc (điểm thấp hơn = ưu tiên hơn):
      - Ưu tiên 1: tên chứa 'CI' hoặc 'IMP' (score 0, ngược lại 1)
      - Ưu tiên 2: bắt đầu bằng số thứ tự "N) " ví dụ "1) IMP..."
                   → lấy số N đó (nhỏ hơn = ưu tiên hơn)
                   → không có số → coi như 9999
    """
    name = os.path.basename(path).upper()

    # Ưu tiên 1: có CI hoặc IMP trong tên không?
    has_keyword = 1 if ("CI" in name or "IMP" in name) else 0
    priority_1  = 0 if has_keyword else 1   # 0 = ưu tiên cao

    # Ưu tiên 2: số thứ tự đứng đầu tên "N) ..."
    m = re.match(r"^(\d+)\s*[).]", name)
    priority_2 = int(m.group(1)) if m else 9999

    return (priority_1, priority_2)


def _auto_pick_best(matches: list[str]) -> str:
    """
    Tự động chọn 1 file tốt nhất từ danh sách nhiều kết quả
    dựa theo _priority_score.
    """
    return min(matches, key=_priority_score)


def search_pdf_by_names(folder: str, names: list[str]) -> dict[str, list[str]]:
    """
    Tìm file PDF khớp với danh sách tên/từ khóa.
    Trả về { tên_tìm: [đường_dẫn, ...] } — đã sắp xếp theo độ ưu tiên.
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
        # Sắp xếp theo ưu tiên để file tốt nhất lên đầu
        matched.sort(key=_priority_score)
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
                    fields["File"]  = filename
                    fields["_path"] = pdf_path
                    return fields
    except Exception as exc:
        return {"File": filename, "error": str(exc), "_path": pdf_path}
    return None