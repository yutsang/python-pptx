#!/usr/bin/env python3
"""Shared PDF/image → compressed vision payload helpers for contract tools.

The Workbench gateway rejects request bodies over ~4MB. Base64 + JSON wrapper
inflates payloads, so pages are rasterized to JPEG and downscaled until each
image fits a safe byte budget.
"""
from __future__ import annotations

import base64
import io
from pathlib import Path
from typing import List, Optional, Sequence, Tuple

try:
    import pypdfium2 as pdfium
except ImportError:
    pdfium = None

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

# Gateway limit observed in production: 4194304 bytes for the whole request body.
# Leave headroom for prompt text + JSON envelope when multiple images are sent.
MAX_REQUEST_BODY_BYTES = 4_194_304
SAFE_SINGLE_IMAGE_BYTES = 2_800_000
SAFE_MULTI_IMAGE_BYTES = 450_000
# Data URLs expand JPEG bytes by roughly 4/3 due to base64. Keep the raw-image
# budget below 2.5MB so the encoded images plus prompt stay under the 4MB cap.
_MULTI_IMAGE_BODY_BUDGET = 2_450_000


def multi_image_byte_budget(n_images: int) -> int:
    """Per-image JPEG budget so n_images fit under the gateway body cap."""
    n = max(1, int(n_images))
    return max(180_000, min(SAFE_MULTI_IMAGE_BYTES, _MULTI_IMAGE_BODY_BUDGET // n))

_IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".gif", ".webp"}
_MIN_TEXT_CHARS_PER_PAGE = 20


def is_image_file(path: Path) -> bool:
    return path.suffix.lower() in _IMAGE_EXTS


def pdf_page_count(pdf_path: Path) -> int:
    if pdfium is None:
        raise RuntimeError("pypdfium2 not installed -- run `pip install pypdfium2`")
    return len(pdfium.PdfDocument(str(pdf_path)))


def pdf_is_digital(pdf_path: Path) -> bool:
    """True when every page has a real extractable text layer."""
    if PdfReader is None:
        return False
    try:
        reader = PdfReader(str(pdf_path))
    except Exception:
        return False
    if not reader.pages:
        return False
    for page in reader.pages:
        try:
            text = page.extract_text() or ""
        except Exception:
            text = ""
        if len(text.strip()) < _MIN_TEXT_CHARS_PER_PAGE:
            return False
    return True


def extract_pdf_text(pdf_path: Path, max_chars: int = 60000) -> str:
    if PdfReader is None:
        raise RuntimeError("pypdf not installed -- run `pip install pypdf`")
    reader = PdfReader(str(pdf_path))
    parts: List[str] = []
    for i, page in enumerate(reader.pages, 1):
        try:
            text = page.extract_text() or ""
        except Exception:
            text = ""
        parts.append(f"--- page {i} ---\n{text.strip()}")
    joined = "\n\n".join(parts).strip()
    if len(joined) > max_chars:
        head = joined[: max_chars // 2]
        tail = joined[-(max_chars // 2) :]
        return head + "\n\n...[truncated]...\n\n" + tail
    return joined


def select_pages(page_count: int, max_pages: int = 10) -> List[int]:
    """1-indexed page picks for lease contracts.

    Commercial terms (term/rent/deposit/free-rent) are rarely only on the
    cover + signature page of a 30+ page scan. Bias heavily to the front,
    sprinkle a few mid-doc pages, and always keep the last page.
    """
    if page_count <= 0:
        return []
    if page_count <= max_pages:
        return list(range(1, page_count + 1))

    front_n = max(3, int(round(max_pages * 0.6)))
    front_n = min(front_n, max_pages - 1, page_count)
    pages: List[int] = list(range(1, front_n + 1))

    remaining_slots = max_pages - len(pages) - 1  # reserve last page
    if remaining_slots > 0 and page_count > front_n + 1:
        start = front_n + 1
        end = page_count - 1
        if end >= start:
            if remaining_slots == 1:
                mids = [(start + end) // 2]
            else:
                step = (end - start) / (remaining_slots - 1)
                mids = [int(round(start + i * step)) for i in range(remaining_slots)]
            for m in mids:
                if m not in pages and start <= m <= end:
                    pages.append(m)

    if page_count not in pages:
        pages.append(page_count)

    # Dedupe, preserve order; if over budget keep front + last.
    seen = set()
    out: List[int] = []
    for p in pages:
        if p in seen or p < 1 or p > page_count:
            continue
        seen.add(p)
        out.append(p)
    if len(out) > max_pages:
        out = out[: max_pages - 1] + [page_count]
        seen = set()
        deduped: List[int] = []
        for p in out:
            if p not in seen:
                seen.add(p)
                deduped.append(p)
        out = deduped
    return out


def _pil_to_jpeg_bytes(pil_image, quality: int, max_edge: int) -> bytes:
    img = pil_image.convert("RGB")
    w, h = img.size
    scale = min(1.0, float(max_edge) / float(max(w, h)))
    if scale < 1.0:
        img = img.resize((max(1, int(w * scale)), max(1, int(h * scale))))
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=quality, optimize=True)
    return buf.getvalue()


def rasterize_page_jpeg(
    pdf_path: Path,
    page_num: int,
    *,
    max_bytes: int = SAFE_SINGLE_IMAGE_BYTES,
    dpi_start: int = 140,
    max_edge_start: int = 1800,
) -> bytes:
    """Rasterize one PDF page to JPEG under max_bytes."""
    if pdfium is None:
        raise RuntimeError("pypdfium2 not installed -- run `pip install pypdfium2`")
    pdf = pdfium.PdfDocument(str(pdf_path))
    if page_num < 1 or page_num > len(pdf):
        raise ValueError(f"Page {page_num} out of range ({len(pdf)} page(s)).")
    page = pdf[page_num - 1]

    attempts = [
        (dpi_start, 88, max_edge_start),
        (dpi_start, 82, min(max_edge_start, 2800)),
        (160, 76, min(max_edge_start, 2400)),
        (120, 70, 1400),
        (100, 60, 1200),
        (85, 50, 1000),
        (72, 40, 800),
    ]
    last: Optional[bytes] = None
    for dpi, quality, max_edge in attempts:
        bitmap = page.render(scale=dpi / 72.0)
        raw = _pil_to_jpeg_bytes(bitmap.to_pil(), quality=quality, max_edge=max_edge)
        last = raw
        if len(raw) <= max_bytes:
            return raw
    assert last is not None
    return last


def rasterize_page_tile_jpegs(
    pdf_path: Path,
    page_num: int,
    *,
    tile_count: int = 3,
    dpi: int = 300,
    max_bytes_each: int = 2_450_000,
) -> List[Tuple[int, bytes]]:
    """Render one page as overlapping horizontal strips for dense tables."""
    if pdfium is None:
        raise RuntimeError("pypdfium2 not installed -- run `pip install pypdfium2`")
    pdf = pdfium.PdfDocument(str(pdf_path))
    if page_num < 1 or page_num > len(pdf):
        raise ValueError(f"Page {page_num} out of range ({len(pdf)} page(s)).")
    page = pdf[page_num - 1]
    image = page.render(scale=dpi / 72.0).to_pil().convert("RGB")
    width, height = image.size
    # Remove scanner margins, then split vertically with enough overlap to
    # avoid losing a table row exactly at a tile boundary.
    page_crop = image.crop((
        int(width * 0.015),
        int(height * 0.015),
        int(width * 0.985),
        int(height * 0.985),
    ))
    width, height = page_crop.size
    count = max(2, int(tile_count))
    overlap = max(30, int(height * 0.025))
    out: List[Tuple[int, bytes]] = []
    for index in range(count):
        top = max(0, int(index * height / count) - overlap)
        bottom = min(height, int((index + 1) * height / count) + overlap)
        tile = page_crop.crop((0, top, width, bottom))
        last: Optional[bytes] = None
        for quality, max_edge in [(90, 3400), (84, 3000), (76, 2600), (68, 2200)]:
            raw = _pil_to_jpeg_bytes(tile, quality=quality, max_edge=max_edge)
            last = raw
            if len(raw) <= max_bytes_each:
                break
        assert last is not None
        out.append((index + 1, last))
    return out


def image_file_to_jpeg_bytes(path: Path, max_bytes: int = SAFE_SINGLE_IMAGE_BYTES) -> bytes:
    from PIL import Image

    img = Image.open(path)
    attempts = [(80, 1800), (70, 1400), (60, 1200), (50, 1000), (40, 800)]
    last: Optional[bytes] = None
    for quality, max_edge in attempts:
        raw = _pil_to_jpeg_bytes(img, quality=quality, max_edge=max_edge)
        last = raw
        if len(raw) <= max_bytes:
            return raw
    assert last is not None
    return last


def to_data_url(image_bytes: bytes, mime: str = "image/jpeg") -> str:
    return f"data:{mime};base64,{base64.b64encode(image_bytes).decode('ascii')}"


def build_page_data_urls(
    pdf_path: Path,
    pages: Sequence[int],
    *,
    max_bytes_each: int = SAFE_MULTI_IMAGE_BYTES,
) -> List[Tuple[int, str]]:
    out: List[Tuple[int, str]] = []
    for page_num in pages:
        raw = rasterize_page_jpeg(pdf_path, page_num, max_bytes=max_bytes_each)
        out.append((page_num, to_data_url(raw)))
    return out
