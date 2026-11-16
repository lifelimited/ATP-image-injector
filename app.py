"""
Unit Test Image Injector v1.5.0 — Portable GUI App
====================================================
Inserts device photos from SN folders into the
Unit Test Word document, matching by Serial Number.

Supports: .jpg .jpeg .png .bmp .tiff .heic .heif
Auto-converts HEIC/HEIF (iPhone photos) to JPG.

Changelog:
  v1.5.0 — Starting table selection, arbitrary table layout picker, auto device model detection & filtering
  v1.4.0 — OCR engine upgraded to OpenThai2.0-OCR-Text (Thai+English, CPU-only)
  v1.3.1 — Fix image corruption in Word caused by missing DPI headers
  v1.3.0 — Smart Auto-Rotate using EasyOCR (GPU accelerated)
  v1.2.1 — Auto-Rotate images (EXIF based)
  v1.2.0 — Dynamic interactive table cell selection
  v1.1.1 — Bug fixes and optimizations
  v1.1.0 — Customizable folder keyword, Stop button, version display
  v1.0.0 — Initial release

Usage: python app.py
"""

__version__ = "1.5.0"
__app_name__ = "Unit Test Image Injector"

import sys
import os
import re
import glob
import tempfile
import atexit
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

# ─────────────────────────────────────────────
# Core Engine
# ─────────────────────────────────────────────
VALID_IMG_EXTS = {'.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.heic', '.heif'}
_temp_files = []


_ocr_init_failed = False

def get_best_orientation(img_path: str, log_cb=None) -> int:
    """Find the PIL rotation angle (0, 90, 180, 270, counter-clockwise) that makes
    the text upright, using the shared OCR engine (OpenThai2.0-OCR-Text by default,
    see ocr_engine.py). Returns 0 when OCR is unavailable or finds too little text."""
    global _ocr_init_failed
    if _ocr_init_failed:
        return 0
    try:
        from PIL import Image as PILImage
        from ocr_engine import get_engine, detect_orientation
        engine = get_engine(log_cb=log_cb)
    except Exception as e:
        _ocr_init_failed = True
        if log_cb: log_cb(f"      [OCR Disabled] Could not start OCR engine: {e}")
        return 0

    with PILImage.open(img_path) as img:
        best_angle, best_score = detect_orientation(img, engine)

    # Too little readable text (score < 15) is probably noise. Trust EXIF.
    if best_score < 15:
        if log_cb: log_cb(f"      [OCR] Score too low ({best_score:.0f} < 15), trusting EXIF.")
        return 0
    return best_angle




def setup_temp_dir():
    for f in _temp_files:
        try:
            os.remove(f)
        except OSError:
            pass


def cleanup_temp_files():
    for f in _temp_files:
        try:
            os.remove(f)
        except OSError:
            pass


atexit.register(cleanup_temp_files)


def convert_heic_to_jpg(heic_path: str) -> str:
    """Converts a HEIC/HEIF image to a temporary JPEG and returns the new path."""
    try:
        from PIL import Image as PILImage, ImageOps
        import pillow_heif
        pillow_heif.register_heif_opener()
    except ImportError:
        pass
    
    img = PILImage.open(heic_path)
    
    # HEIC images have EXIF rotation metadata. If we just save as JPEG, Pillow strips the EXIF!
    # So we MUST forcefully apply the EXIF rotation here before converting.
    try:
        from PIL import ImageOps
        img_rotated = ImageOps.exif_transpose(img)
        if img_rotated is not None:
            img = img_rotated
    except Exception:
        pass
    
    # Save to temp
    fd, tmp_path = tempfile.mkstemp(suffix='.jpg', prefix='heic_')
    os.close(fd)
    
    img.convert('RGB').save(tmp_path, 'JPEG', quality=90, dpi=(72, 72))
    _temp_files.append(tmp_path)
    return tmp_path


class InjectionResult(dict):
    """
    Result dictionary with backward-compatible 4-tuple unpacking,
    property access, and integer indexing.
    """
    def __init__(self, injected: int = 0, img_count: int = 0, no_folder: int = 0, errors: list = None):
        errs = list(errors) if errors is not None else []
        super().__init__({
            "injected": int(injected),
            "img_count": int(img_count),
            "no_folder": int(no_folder),
            "errors": errs,
        })

    @property
    def injected(self) -> int:
        return self["injected"]

    @property
    def img_count(self) -> int:
        return self["img_count"]

    @property
    def no_folder(self) -> int:
        return self["no_folder"]

    @property
    def errors(self) -> list:
        return self["errors"]

    def __iter__(self):
        # Enables 4-tuple unpacking: injected, img_count, no_folder, errors = res
        return iter((self["injected"], self["img_count"], self["no_folder"], self["errors"]))

    def __getitem__(self, item):
        if isinstance(item, int):
            return (self["injected"], self["img_count"], self["no_folder"], self["errors"])[item]
        return super().__getitem__(item)


# Regex patterns for model detection and matching
CAMERA_SUBDIR_REGEX = re.compile(
    r'^(?:dcim|camera|photos?|images?|\d{3}(?:apple|media|canon|nikon|fuji|pana|olymp|pentx|andro|[a-z0-9_-]{2,}))(?:\b|_|$)',
    re.IGNORECASE
)
MODEL_CODE_REGEX = re.compile(r'\bCP-[0-9A-Z]+(?:-[0-9A-Z]+)*(?:=|\b)', re.IGNORECASE)
MODEL_LABEL_REGEX = re.compile(r'(?:Name|P/N|Part\s*Number|Model)\s*:\s*([A-Za-z0-9\-=_ ]+?)(?=\s*(?:Name|P/N|Part\s*Number|S/N|Serial\s*Number|Model)\s*:|\s*$)', re.IGNORECASE)
MODEL_NUM_REGEX = re.compile(r'\b(?:78\d{2}|88\d{2}|69\d{2}|79\d{2})\b')
PART_NUM_LABEL_CELL_REGEX = re.compile(r'^(?:P/N|Part\s*Number|Model)\s*:?$', re.IGNORECASE)


def model_matches(candidate: str, target: str) -> bool:
    """
    Check if candidate model string matches target model string.
    Supports exact match, Cisco prefix/suffix normalization, and word-boundary matching.
    e.g. '7821' matches 'CP-7821-K9=', 'CP-7821' matches '7821', '8841' matches 'Cisco IP Phone 8841'.
    """
    c = str(candidate).strip().upper() if candidate else ""
    t = str(target).strip().upper() if target else ""
    if not c or not t:
        return False
    if c == t:
        return True

    # Strip Cisco prefix 'CP-' and suffix '-K9=', '-K9', '='
    c_core = re.sub(r'^CP-', '', c)
    c_core = re.sub(r'(-K9)?=?$', '', c_core).strip()
    t_core = re.sub(r'^CP-', '', t)
    t_core = re.sub(r'(-K9)?=?$', '', t_core).strip()

    if c_core and t_core and c_core == t_core:
        return True

    # Substring / word boundary check
    if c_core and re.search(r'\b' + re.escape(c_core) + r'\b', t):
        return True
    if t_core and re.search(r'\b' + re.escape(t_core) + r'\b', c):
        return True
    return False


def extract_doc_part_descriptions(doc) -> dict[str, str]:
    """
    Scan multi-column summary tables (containing 'Part Number' and 'Description' columns)
    and return a mapping of {part_number: primary_description} (e.g. {'02355UEL': 'S6750-H48Y8C'}).
    """
    pn_to_desc: dict[str, str] = {}
    if not doc or not hasattr(doc, "tables"):
        return pn_to_desc
    for table in doc.tables:
        if len(table.rows) < 2 or len(table.columns) < 3:
            continue
        headers = [c.text.strip().lower() for c in table.rows[0].cells]
        pn_col = next((i for i, h in enumerate(headers) if "part number" in h or h in ("p/n", "pn")), None)
        desc_col = next((i for i, h in enumerate(headers) if "description" in h or "model" in h or "รายการ" in h), None)
        if pn_col is None or desc_col is None:
            continue
        for row in table.rows[1:]:
            cells = row.cells
            if max(pn_col, desc_col) >= len(cells):
                continue
            pn = cells[pn_col].text.strip()
            desc_raw = cells[desc_col].text.strip().replace("\n", " ")
            if not pn or not desc_raw:
                continue
            # Take primary model name before comma or parenthesis if long
            primary = desc_raw.split(",")[0].strip()
            if primary and pn not in pn_to_desc:
                pn_to_desc[pn] = primary
    return pn_to_desc


def extract_table_models(table, sn_row_col: tuple = (5, 3), pn_to_desc: dict[str, str] | None = None) -> list[str]:
    """Extract all device models from a single table using prioritized heuristics."""
    sn_r, sn_c = sn_row_col
    found = []

    def add_model(m_str: str):
        if not m_str:
            return
        s = m_str.strip()
        if s and s not in found:
            found.append(s)

    # Heuristic 1: Check Row 2 (Name : CP-XXXX / P/N : CP-XXXX / Part Number: XXXX)
    if len(table.rows) > 2:
        r2_text = " ".join(cell.text for cell in table.rows[2].cells)
        for m in MODEL_LABEL_REGEX.finditer(r2_text):
            val = m.group(1).strip()
            if pn_to_desc and val in pn_to_desc:
                add_model(f"{pn_to_desc[val]} ({val})")
            add_model(val)
        for m in MODEL_CODE_REGEX.finditer(r2_text):
            add_model(m.group(0))

    # Heuristic 2: Check Row sn_row_col[0] (Part Number at sn_c - 1, Detail at sn_c - 2)
    if len(table.rows) > sn_r:
        row_cells = table.rows[sn_r].cells
        if 0 <= sn_c - 1 < len(row_cells):
            pn_text = row_cells[sn_c - 1].text.strip()
            for m in MODEL_CODE_REGEX.finditer(pn_text):
                add_model(m.group(0))
        if 0 <= sn_c - 2 < len(row_cells):
            detail_text = row_cells[sn_c - 2].text.strip()
            for m in MODEL_CODE_REGEX.finditer(detail_text):
                add_model(m.group(0))

    # Heuristic 3: Key-value table rows where cell 0 is 'Part Number:' / 'P/N:' / 'Model:' and cell 1 is the code
    for row in table.rows:
        cells = row.cells
        if len(cells) >= 2 and PART_NUM_LABEL_CELL_REGEX.match(cells[0].text.strip()):
            val = cells[1].text.strip()
            if val:
                if pn_to_desc and val in pn_to_desc:
                    add_model(f"{pn_to_desc[val]} ({val})")
                add_model(val)
        row_text = " ".join(cell.text for cell in cells)
        for m in MODEL_CODE_REGEX.finditer(row_text):
            add_model(m.group(0))

    # If no full model codes found, look for 4-digit model numbers
    if not found:
        if len(table.rows) > 2:
            r2_text = " ".join(cell.text for cell in table.rows[2].cells)
            for m in MODEL_NUM_REGEX.finditer(r2_text):
                add_model(m.group(0))
        if len(table.rows) > sn_r:
            row_cells = table.rows[sn_r].cells
            if 0 <= sn_c - 1 < len(row_cells):
                for m in MODEL_NUM_REGEX.finditer(row_cells[sn_c - 1].text):
                    add_model(m.group(0))
            if 0 <= sn_c - 2 < len(row_cells):
                for m in MODEL_NUM_REGEX.finditer(row_cells[sn_c - 2].text):
                    add_model(m.group(0))
        for row in table.rows:
            row_text = " ".join(cell.text for cell in row.cells)
            for m in MODEL_NUM_REGEX.finditer(row_text):
                add_model(m.group(0))

    # Normalize/deduplicate: drop bare 4-digit numbers if a full model code containing that number is already present
    cleaned = []
    for item in found:
        if MODEL_NUM_REGEX.fullmatch(item) and any(item in other for other in found if other != item):
            continue
        if item not in cleaned:
            cleaned.append(item)

    return cleaned


def extract_table_model(table, sn_row_col: tuple = (5, 3)) -> str | None:
    """Extract primary device model from a single table using prioritized heuristics."""
    models = extract_table_models(table, sn_row_col=sn_row_col)
    return models[0] if models else None


def detect_docx_models(docx_path: str, sn_row_col: tuple = (5, 3)) -> list[str]:
    """
    Scan docx tables and return sorted unique list of detected device models
    (e.g. ['CP-7821-K9=', 'CP-8841-K9=']).
    """
    if not docx_path or not os.path.isfile(docx_path):
        return []
    try:
        from docx import Document
        doc = Document(docx_path)
    except Exception:
        return []

    pn_to_desc = extract_doc_part_descriptions(doc)
    detected = set()
    for table in doc.tables:
        table_models = extract_table_models(table, sn_row_col=sn_row_col, pn_to_desc=pn_to_desc)
        for model in table_models:
            # Skip bare Part Numbers if their descriptive 'Model (PN)' form is also present in table_models
            if pn_to_desc and model in pn_to_desc and f"{pn_to_desc[model]} ({model})" in table_models:
                continue
            detected.add(model)

    return sorted(detected)


def detect_folder_models(sn_root: str) -> list[str]:
    """
    Scan folder hierarchy and return sorted unique list of detected device models from directory names.
    Handles both nested model root (e.g. SVB containing 7821/, 8841/) and single-model root
    (e.g. SVB/7821/ containing direct S/N folders). Subfolders without valid image files are ignored.
    """
    if not sn_root or not os.path.isdir(sn_root):
        return []

    def dir_has_images(path: str) -> bool:
        for _, _, files in os.walk(path):
            if any(os.path.splitext(f)[1].lower() in VALID_IMG_EXTS for f in files):
                return True
        return False

    detected = set()
    try:
        subdirs = [d for d in os.listdir(sn_root)
                   if os.path.isdir(os.path.join(sn_root, d)) and not d.startswith(('.', '__'))]
    except OSError:
        return []

    for d in subdirs:
        d_path = os.path.join(sn_root, d)
        if MODEL_CODE_REGEX.search(d) or MODEL_NUM_REGEX.search(d):
            if dir_has_images(d_path):
                detected.add(d)
        else:
            # S/N folders containing camera subfolders (e.g. 100APPLE, DCIM, CAMERA)
            # must NOT be detected as models.
            try:
                inner_dirs = [sd for sd in os.listdir(d_path)
                              if os.path.isdir(os.path.join(d_path, sd)) and not sd.startswith(('.', '__'))]
            except OSError:
                inner_dirs = []
            if inner_dirs:
                valid_inner = [sd for sd in inner_dirs if not CAMERA_SUBDIR_REGEX.search(sd)]
                if valid_inner and (MODEL_CODE_REGEX.search(d) or MODEL_NUM_REGEX.search(d)):
                    for test_sd in valid_inner[:3]:
                        test_sd_path = os.path.join(d_path, test_sd)
                        if dir_has_images(test_sd_path):
                            detected.add(d)
                            break

    if not detected:
        base = os.path.basename(os.path.abspath(sn_root))
        if (MODEL_CODE_REGEX.search(base) or MODEL_NUM_REGEX.search(base)) and dir_has_images(sn_root):
            detected.add(base)

    return sorted(detected)


class SnFolderMap(dict):
    """Dict of {sn: [image_paths]} with an attached sn_to_model mapping."""
    def __init__(self, *args, sn_to_model=None, **kwargs):
        super().__init__(*args, **kwargs)
        self.sn_to_model = sn_to_model if sn_to_model is not None else {}


def find_sn_folders(root: str, keyword: str = "", return_model_map: bool = False):
    """
    Scan root directory for S/N folders containing images.
    Supports both 1-level (root/SN/imgs) and 2-level (root/model/SN/imgs).
    Returns SnFolderMap ({sn: [image_paths]}), which has an attribute .sn_to_model ({sn: model_name}).
    If return_model_map=True, returns (sn_map, sn_to_model).
    """
    sn_map = {}
    sn_to_model = {}
    if not root or not os.path.isdir(root):
        res = SnFolderMap(sn_map, sn_to_model=sn_to_model)
        return (res, sn_to_model) if return_model_map else res

    kw_upper = keyword.strip().upper() if keyword else ""

    def get_images(path):
        imgs = []
        for dirpath, _, filenames in os.walk(path):
            for fname in sorted(filenames):
                if '_converted' in fname:
                    continue
                if os.path.splitext(fname)[1].lower() in VALID_IMG_EXTS:
                    imgs.append(os.path.join(dirpath, fname))
        return imgs

    try:
        entries = sorted(os.listdir(root))
    except OSError:
        entries = []

    for entry in entries:
        if entry.startswith(('.', '__')):
            continue
        entry_path = os.path.join(root, entry)
        if not os.path.isdir(entry_path):
            continue

        # Check if entry is a model folder containing S/N subdirs
        try:
            child_dirs = [d for d in os.listdir(entry_path)
                          if os.path.isdir(os.path.join(entry_path, d)) and not d.startswith(('.', '__'))]
        except OSError:
            child_dirs = []

        is_model_dir = False
        if child_dirs:
            for test_cd in child_dirs[:3]:
                if get_images(os.path.join(entry_path, test_cd)):
                    is_model_dir = True
                    break

        if is_model_dir:
            model_name = entry
            for sd in child_dirs:
                if kw_upper and not sd.upper().startswith(kw_upper):
                    continue
                sd_path = os.path.join(entry_path, sd)
                imgs = get_images(sd_path)
                if imgs:
                    sn_map[sd] = imgs
                    sn_to_model[sd] = model_name
        else:
            if not kw_upper or entry.upper().startswith(kw_upper):
                imgs = get_images(entry_path)
                if imgs:
                    sn_map[entry] = imgs
                    sn_to_model[entry] = os.path.basename(os.path.abspath(root))

    res = SnFolderMap(sn_map, sn_to_model=sn_to_model)
    if return_model_map:
        return res, sn_to_model
    return res


def run_injection(docx_path: str, sn_root: str, output_path: str,
                  keyword: str = "WZP",
                  sn_row_col: tuple = (5, 3), img_row_col: tuple = (12, 0),
                  start_table: int = 1,
                  selected_models: list[str] | None = None,
                  auto_rotate: bool = True, smart_rotate: bool = False,
                  log_callback=None, progress_callback=None,
                  status_callback=None,
                  cancel_event=None) -> InjectionResult:
    """
    Main injection logic. Returns InjectionResult (supports 4-tuple unpacking and dict access).
    - start_table: 1-based table index to begin injection. Tables before this index are untouched.
    - selected_models: list of model strings to inject. Non-matching tables/folders are untouched.
    - status_callback: alias for log_callback.
    - cancel_event: threading.Event — set to stop processing early.
    """
    from docx import Document
    from docx.shared import Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from PIL import Image as PILImage

    status_cb = log_callback or status_callback
    def log(msg):
        if status_cb:
            status_cb(msg)

    def is_cancelled():
        return cancel_event and cancel_event.is_set()

    # ── Step 1: Find SN folders ──
    # If selected_models is active, do not restrict folder scanning by keyword prefix
    model_filtering_active = selected_models is not None
    active_models = [str(m).strip() for m in selected_models if m and str(m).strip()] if model_filtering_active else None
    scan_kw = "" if model_filtering_active else (keyword or "")
    log(f"Scanning folders matching '{scan_kw}*'...")
    sn_folders = find_sn_folders(sn_root, scan_kw)
    sn_to_model = getattr(sn_folders, 'sn_to_model', {})
    log(f"  Found {len(sn_folders)} folder(s) with images")

    if not sn_folders:
        log(f"ERROR: No matching folders found in the selected directory.")
        return InjectionResult(0, 0, 0, [f"No matching folders found"])

    # Show summary (avoid flooding log with 600+ lines)
    sn_sorted = sorted(sn_folders.keys())
    if len(sn_sorted) <= 10:
        for sn in sn_sorted:
            imgs = sn_folders[sn]
            heic_count = sum(1 for i in imgs if os.path.splitext(i)[1].lower() in ('.heic', '.heif'))
            log(f"    {sn}: {len(imgs)} file(s)" + (f" ({heic_count} HEIC)" if heic_count else ""))
    else:
        for sn in sn_sorted[:3]:
            imgs = sn_folders[sn]
            log(f"    {sn}: {len(imgs)} file(s)")
        log(f"    ... ({len(sn_sorted) - 6} more) ...")
        for sn in sn_sorted[-3:]:
            imgs = sn_folders[sn]
            log(f"    {sn}: {len(imgs)} file(s)")

    if is_cancelled():
        log("\n[CANCELLED] Stopped by user.")
        return InjectionResult(0, 0, 0, ["Cancelled"])

    # ── Step 2: Open docx ──
    log(f"\nOpening docx: {os.path.basename(docx_path)}")
    doc = Document(docx_path)
    total_tables = len(doc.tables)
    log(f"  Tables found: {total_tables}")

    # Validate and normalize start_table (1-based)
    try:
        start_table = int(start_table)
    except (TypeError, ValueError, OverflowError):
        start_table = 1
    if start_table < 1:
        start_table = 1

    log(f"  Starting from table: {start_table} (1-based)")
    if start_table > total_tables:
        log(f"  [WARN] start_table ({start_table}) > total tables ({total_tables}). No tables will be processed.")

    keyword_upper = keyword.strip().upper() if keyword else ""

    pn_to_desc = extract_doc_part_descriptions(doc) if model_filtering_active else None

    # ── Step 3: Collect eligible candidate tables in sequential document order ──
    candidates = []
    for t_idx, table in enumerate(doc.tables, start=1):
        if t_idx < start_table:
            # Skip tables prior to start_table without touching DOM
            continue

        try:
            sn_cell = table.rows[sn_row_col[0]].cells[sn_row_col[1]]
            sn = sn_cell.text.strip()
        except (IndexError, AttributeError):
            continue

        if not sn:
            continue

        if model_filtering_active:
            if not active_models:
                continue
            table_models = extract_table_models(table, sn_row_col=sn_row_col, pn_to_desc=pn_to_desc)
            folder_model = sn_to_model.get(sn)
            matched = any(
                any(model_matches(sel, tm) for tm in table_models) or
                (folder_model and model_matches(sel, folder_model))
                for sel in active_models
            )
            if not matched:
                continue
        elif keyword_upper:
            if not sn.upper().startswith(keyword_upper):
                continue

        candidates.append((t_idx, table, sn))

    log(f"  Eligible tables to evaluate: {len(candidates)}")

    if is_cancelled():
        log("\n[CANCELLED] Stopped by user.")
        return InjectionResult(0, 0, 0, ["Cancelled"])

    # ── Step 4: Inject ──
    log(f"\n{'─' * 50}")
    log("Injecting images...")
    log(f"{'─' * 50}")

    injected = 0
    img_count = 0
    no_folder = 0
    errors = []
    total = len(candidates)
    current = 0

    for t_idx, table, sn in candidates:
        if is_cancelled():
            log(f"\n[CANCELLED] Stopped by user at {current}/{total}.")
            break

        current += 1
        if progress_callback:
            progress_callback(current, total)

        if sn not in sn_folders:
            no_folder += 1
            log(f"  [SKIP] {sn} (table {t_idx}): no photo folder found")
            continue

        images = sn_folders[sn]

        try:
            img_cell = table.rows[img_row_col[0]].cells[img_row_col[1]]
        except IndexError:
            err_msg = f"{sn}: table {t_idx} has no row {img_row_col[0]}"
            log(f"  [ERR ] {err_msg}")
            errors.append(err_msg)
            continue

        # Clear cell ONLY immediately before injecting photos
        for para in img_cell.paragraphs:
            for run in para.runs:
                run._element.getparent().remove(run._element)
        while len(img_cell.paragraphs) > 1:
            p = img_cell.paragraphs[-1]._element
            p.getparent().remove(p)

        cell_img_count = 0
        for img_idx, img_path in enumerate(images):
            if is_cancelled():
                break
            try:
                ext = os.path.splitext(img_path)[1].lower()
                if ext in ('.heic', '.heif'):
                    orig = os.path.basename(img_path)
                    img_path = convert_heic_to_jpg(img_path)
                    log(f"      Converted {orig} -> JPG")

                if auto_rotate or smart_rotate:
                    from PIL import Image as PILImage, ImageOps
                    try:
                        with PILImage.open(img_path) as img:
                            modified = False
                            if auto_rotate:
                                img_rotated = ImageOps.exif_transpose(img)
                                if img_rotated is not None and img_rotated != img:
                                    img = img_rotated
                                    modified = True
                                    log(f"      Auto-rotated (EXIF) {os.path.basename(img_path)}")

                            if modified:
                                fd, tmp_path = tempfile.mkstemp(suffix='.jpg', prefix='rot_')
                                os.close(fd)
                                img.convert('RGB').save(tmp_path, 'JPEG', quality=90, dpi=(72, 72))
                                _temp_files.append(tmp_path)
                                img_path = tmp_path

                            if smart_rotate:
                                best_angle = get_best_orientation(img_path, log_cb=log)
                                if best_angle != 0:
                                    with PILImage.open(img_path) as img2:
                                        img_smart = img2.rotate(best_angle, expand=True)
                                        fd, tmp_path_smart = tempfile.mkstemp(suffix='.jpg', prefix='ocr_rot_')
                                        os.close(fd)
                                        img_smart.convert('RGB').save(tmp_path_smart, 'JPEG', quality=90, dpi=(72, 72))
                                        _temp_files.append(tmp_path_smart)
                                        img_path = tmp_path_smart
                                        log(f"      Smart-rotated (OCR) {best_angle}° for {os.path.basename(img_path)}")
                                        modified = True

                    except Exception as e:
                        log(f"      Rotation Error for {os.path.basename(img_path)}: {e}")

                width = Cm(14.0)
                if img_idx == 0:
                    para = img_cell.paragraphs[0]
                else:
                    para = img_cell.add_paragraph()
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = para.add_run()
                run.add_picture(img_path, width=width)
                cell_img_count += 1
            except Exception as e:
                err_msg = f"{sn}/{os.path.basename(img_path)}: {e}"
                log(f"      ERROR: {err_msg}")
                errors.append(err_msg)

        if cell_img_count > 0:
            injected += 1
            img_count += cell_img_count
            log(f"  [OK  ] {sn} (table {t_idx:>3}): {cell_img_count} image(s)")

    # ── Step 5: Save ──
    cancelled = is_cancelled()
    if injected > 0:
        log(f"\n{'═' * 50}")
        log(f"  Tables injected  : {injected}")
        log(f"  Images inserted  : {img_count}")
        log(f"  Tables no folder : {no_folder}")
        log(f"  Errors           : {len(errors)}")
        if cancelled:
            log(f"  Status           : PARTIAL (cancelled)")
        log(f"{'═' * 50}")

        doc.save(output_path)
        log(f"\nSaved -> {output_path}")
    else:
        log("\nNo images were injected. Output file not saved.")

    if cancelled:
        log("Process was cancelled by user.")
    else:
        log("Done!")

    return InjectionResult(injected=injected, img_count=img_count, no_folder=no_folder, errors=errors)


class TableSelectionDialog(tk.Toplevel):
    def __init__(self, parent, docx_path: str, current_sn_rc: tuple, current_img_rc: tuple, initial_table_idx: int = 0):
        super().__init__(parent)
        self.title("Configure Table Layout")
        self.geometry("960x640")
        self.minsize(800, 520)

        # Fallback theme colors to prevent KeyError in headless or partial environments
        default_colors = {
            "bg": "#1a1a2e", "card": "#16213e", "card_alt": "#0f3460",
            "accent": "#e94560", "accent_ok": "#00b894", "text": "#eaeaea",
            "text_dim": "#8892b0", "input_bg": "#0d1b2a", "input_fg": "#e0e0e0",
            "border": "#233554", "btn_run": "#00b894", "btn_run_fg": "#ffffff",
            "btn_stop": "#e94560", "btn_stop_fg": "#ffffff",
            "btn_pick": "#0f3460", "btn_pick_fg": "#e0e0e0"
        }
        parent_colors = getattr(parent, "COLORS", {})
        self.COLORS = {**default_colors, **parent_colors}

        self.configure(bg=self.COLORS["bg"])
        try:
            self.transient(parent)
            self.grab_set()
        except Exception:
            pass

        self.docx_path = docx_path
        self.result = None
        self.sn_rc = current_sn_rc
        self.img_rc = current_img_rc
        self.mode = tk.StringVar(value="sn")
        self.buttons = {}

        # 1. Load Word document once and determine total tables
        from docx import Document
        try:
            self.doc = Document(docx_path)
            self.total_tables = len(self.doc.tables) if self.doc else 0
        except Exception:
            self.doc = None
            self.total_tables = 0

        # Safely coerce initial table index to int
        try:
            init_idx = int(initial_table_idx)
        except (TypeError, ValueError, OverflowError):
            init_idx = 0

        # Clamp initial table index safely
        if self.total_tables > 0:
            self.current_table_idx = max(0, min(init_idx, self.total_tables - 1))
        else:
            self.current_table_idx = 0

        # ── Header ──
        header = tk.Frame(self, bg=self.COLORS["bg"])
        header.pack(fill="x", padx=10, pady=(10, 5))

        tk.Label(
            header, text="Select Cell Locations",
            bg=self.COLORS["bg"], fg=self.COLORS["text"],
            font=("Segoe UI", 14, "bold")
        ).pack(side="left")

        instructions = tk.Label(
            header,
            text="1. Select 'S/N' mode and click Serial Number cell.  2. Select 'Image' mode and click Image cell.  3. Click Approve.",
            bg=self.COLORS["bg"], fg=self.COLORS["text_dim"],
            font=("Segoe UI", 9), justify="left"
        )
        instructions.pack(side="left", padx=15)

        # ── Action Controls (Mode Radios & Approve/Cancel) ──
        controls = tk.Frame(self, bg=self.COLORS["card"])
        controls.pack(fill="x", padx=10, pady=5)

        tk.Radiobutton(
            controls, text="Select S/N Cell (Blue)", variable=self.mode, value="sn",
            bg=self.COLORS["card"], fg="#00a8ff", selectcolor=self.COLORS["bg"],
            font=("Segoe UI", 10, "bold")
        ).pack(side="left", padx=10, pady=5)

        tk.Radiobutton(
            controls, text="Select Image Cell (Green)", variable=self.mode, value="img",
            bg=self.COLORS["card"], fg="#00b894", selectcolor=self.COLORS["bg"],
            font=("Segoe UI", 10, "bold")
        ).pack(side="left", padx=10, pady=5)

        tk.Button(
            controls, text="Approve",
            bg=self.COLORS["btn_run"], fg=self.COLORS["btn_run_fg"],
            font=("Segoe UI", 10, "bold"), command=self._approve
        ).pack(side="right", padx=10, pady=5)

        tk.Button(
            controls, text="Cancel",
            bg=self.COLORS["btn_stop"], fg=self.COLORS["btn_stop_fg"],
            font=("Segoe UI", 10, "bold"), command=self.destroy
        ).pack(side="right")

        # ── Table Navigation Bar ──
        nav_bar = tk.Frame(self, bg=self.COLORS["card"])
        nav_bar.pack(fill="x", padx=10, pady=(0, 5))

        self.prev_btn = tk.Button(
            nav_bar, text="◀ Previous",
            bg=self.COLORS.get("btn_pick", "#0f3460"), fg=self.COLORS.get("btn_pick_fg", "#e0e0e0"),
            font=("Segoe UI", 9, "bold"), activebackground=self.COLORS.get("card_alt", "#0f3460"),
            command=self._prev_table
        )
        self.prev_btn.pack(side="left", padx=(10, 5), pady=5)

        tk.Label(
            nav_bar, text="Table:",
            bg=self.COLORS["card"], fg=self.COLORS["text"],
            font=("Segoe UI", 9, "bold")
        ).pack(side="left", padx=(5, 3))

        self.table_var = tk.StringVar()
        combo_values = [f"Table {i+1} of {self.total_tables}" for i in range(self.total_tables)] if self.total_tables > 0 else ["No tables"]
        self.table_combo = ttk.Combobox(
            nav_bar, textvariable=self.table_var,
            values=combo_values, width=16,
            state="readonly" if self.total_tables > 0 else "disabled"
        )
        self.table_combo.pack(side="left", padx=3)
        self.table_combo.bind("<<ComboboxSelected>>", self._on_table_combo_change)
        self.table_combo.bind("<Return>", self._on_table_combo_change)
        self.table_spin = self.table_combo  # Attribute alias for spinbox compatibility

        self.next_btn = tk.Button(
            nav_bar, text="Next ▶",
            bg=self.COLORS.get("btn_pick", "#0f3460"), fg=self.COLORS.get("btn_pick_fg", "#e0e0e0"),
            font=("Segoe UI", 9, "bold"), activebackground=self.COLORS.get("card_alt", "#0f3460"),
            command=self._next_table
        )
        self.next_btn.pack(side="left", padx=(5, 10), pady=5)

        self.table_info_label = tk.Label(
            nav_bar, text="",
            bg=self.COLORS["card"], fg=self.COLORS["text_dim"],
            font=("Segoe UI", 9)
        )
        self.table_info_label.pack(side="left", padx=5)

        self.coords_label = tk.Label(
            nav_bar,
            text=f"S/N: R{self.sn_rc[0]+1}C{self.sn_rc[1]+1}  |  Image: R{self.img_rc[0]+1}C{self.img_rc[1]+1}",
            bg=self.COLORS["card"], fg=self.COLORS.get("accent_ok", "#00b894"),
            font=("Segoe UI", 9, "bold")
        )
        self.coords_label.pack(side="right", padx=10)

        # ── Table Grid Canvas & Scrollbars ──
        container = tk.Frame(self, bg=self.COLORS["bg"])
        container.pack(fill="both", expand=True, padx=10, pady=5)

        self.canvas = tk.Canvas(container, bg=self.COLORS["input_bg"], highlightthickness=0)
        vsb = ttk.Scrollbar(container, orient="vertical", command=self.canvas.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=self.canvas.xview)
        self.grid_frame = tk.Frame(self.canvas, bg=self.COLORS["input_bg"])

        self.grid_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.grid_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

        # Initial render of the active table
        self._render_table(self.current_table_idx)

    def _render_table(self, table_idx: int):
        """
        Clears existing button grid and renders table at table_idx (0-based).
        Updates navigation controls, coordinate highlighting, and scroll offset.
        """
        for widget in self.grid_frame.winfo_children():
            widget.destroy()
        self.buttons.clear()

        if not self.doc or self.total_tables == 0:
            lbl = tk.Label(
                self.grid_frame, text="No tables found in the document.",
                fg=self.COLORS.get("accent", "#e94560"), bg=self.COLORS.get("input_bg", "#0d1b2a"),
                font=("Segoe UI", 11, "italic")
            )
            lbl.pack(padx=20, pady=20)
            self.current_table_idx = 0
            self._update_nav_controls(0, 0)
            return

        clamped_idx = max(0, min(table_idx, self.total_tables - 1))
        self.current_table_idx = clamped_idx

        table = self.doc.tables[self.current_table_idx]
        num_rows = len(table.rows)
        num_cols = len(table.columns) if num_rows > 0 else 0

        for r_idx, row in enumerate(table.rows):
            for c_idx, cell in enumerate(row.cells):
                raw_text = cell.text.strip()
                text = raw_text[:30] if raw_text else "(Empty)"

                btn = tk.Button(
                    self.grid_frame, text=text, width=15, height=2,
                    bg=self.COLORS.get("card", "#16213e"), fg=self.COLORS.get("text", "#eaeaea"),
                    activebackground=self.COLORS.get("card_alt", "#0f3460"),
                    command=lambda r=r_idx, c=c_idx: self._on_cell_click(r, c)
                )
                btn.grid(row=r_idx, column=c_idx, padx=1, pady=1, sticky="nsew")
                self.buttons[(r_idx, c_idx)] = btn

        if hasattr(self, "canvas") and self.canvas:
            self.canvas.xview_moveto(0.0)
            self.canvas.yview_moveto(0.0)

        self._update_colors()
        self._update_nav_controls(num_rows, num_cols)

    def _update_nav_controls(self, num_rows: int = 0, num_cols: int = 0):
        if not hasattr(self, "prev_btn") or not hasattr(self, "next_btn"):
            return

        if self.total_tables <= 0:
            self.prev_btn.configure(state="disabled")
            self.next_btn.configure(state="disabled")
            if hasattr(self, "table_combo"):
                self.table_combo.configure(state="disabled")
            if hasattr(self, "table_info_label"):
                self.table_info_label.configure(text="No tables available")
            return

        self.prev_btn.configure(state="normal" if self.current_table_idx > 0 else "disabled")
        self.next_btn.configure(state="normal" if self.current_table_idx < self.total_tables - 1 else "disabled")

        if hasattr(self, "table_combo"):
            self.table_combo.configure(state="readonly")
            self.table_var.set(f"Table {self.current_table_idx + 1} of {self.total_tables}")

        if hasattr(self, "table_info_label"):
            self.table_info_label.configure(text=f"Rows: {num_rows} | Cols: {num_cols}")

    def _prev_table(self):
        """Navigate to previous table, clamped at index 0."""
        if self.total_tables <= 0:
            return
        target_idx = max(0, self.current_table_idx - 1)
        self._render_table(target_idx)

    def _next_table(self):
        """Navigate to next table, clamped at index total_tables - 1."""
        if self.total_tables <= 0:
            return
        target_idx = min(self.total_tables - 1, self.current_table_idx + 1)
        self._render_table(target_idx)

    def _on_table_combo_change(self, event=None):
        """Handler for combobox selection or manual input."""
        if self.total_tables <= 0:
            return
        val = ""
        if hasattr(self, "table_combo") and self.table_combo:
            val = str(self.table_combo.get()).strip()
        elif hasattr(self, "table_var"):
            val = str(self.table_var.get()).strip()

        match = re.search(r'\d+', val)
        if match:
            try:
                table_num = int(match.group(0))
                target_idx = max(0, min(table_num - 1, self.total_tables - 1))
                self._render_table(target_idx)
            except Exception:
                pass

    _on_table_change = _on_table_combo_change

    def _load_table(self, docx_path: str = None, table_idx: int = 0):
        """Backwards compatibility helper."""
        if docx_path and (self.doc is None or docx_path != self.docx_path):
            from docx import Document
            try:
                self.docx_path = docx_path
                self.doc = Document(docx_path)
                self.total_tables = len(self.doc.tables) if self.doc else 0
            except Exception:
                self.doc = None
                self.total_tables = 0
        self._render_table(table_idx)

    def _on_cell_click(self, r, c):
        if self.mode.get() == "sn":
            self.sn_rc = (r, c)
        else:
            self.img_rc = (r, c)
        self._update_colors()
        if hasattr(self, "coords_label"):
            self.coords_label.configure(
                text=f"S/N: R{self.sn_rc[0]+1}C{self.sn_rc[1]+1}  |  Image: R{self.img_rc[0]+1}C{self.img_rc[1]+1}"
            )

    def _update_colors(self):
        for (r, c), btn in self.buttons.items():
            if (r, c) == self.sn_rc:
                btn.configure(bg="#00a8ff", fg="white")
            elif (r, c) == self.img_rc:
                btn.configure(bg="#00b894", fg="white")
            else:
                btn.configure(bg=self.COLORS.get("card", "#16213e"), fg=self.COLORS.get("text", "#eaeaea"))

    def _approve(self):
        self.result = (self.sn_rc, self.img_rc)
        self.destroy()


# ─────────────────────────────────────────────
# GUI Application
# ─────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title(f"{__app_name__} v{__version__}")
        self.geometry("850x720")
        self.minsize(720, 580)
        self.configure(bg="#1a1a2e")

        # ── Colors ──
        self.COLORS = {
            "bg":         "#1a1a2e",
            "card":       "#16213e",
            "card_alt":   "#0f3460",
            "accent":     "#e94560",
            "accent_ok":  "#00b894",
            "text":       "#eaeaea",
            "text_dim":   "#8892b0",
            "input_bg":   "#0d1b2a",
            "input_fg":   "#e0e0e0",
            "border":     "#233554",
            "btn_run":    "#00b894",
            "btn_run_fg": "#ffffff",
            "btn_stop":   "#e94560",
            "btn_stop_fg":"#ffffff",
            "btn_pick":   "#0f3460",
            "btn_pick_fg":"#e0e0e0",
        }

        style = ttk.Style(self)
        style.theme_use('clam')

        style.configure("Card.TFrame",  background=self.COLORS["card"])
        style.configure("BG.TFrame",    background=self.COLORS["bg"])
        style.configure("Title.TLabel",
                         background=self.COLORS["bg"],
                         foreground=self.COLORS["text"],
                         font=("Segoe UI", 18, "bold"))
        style.configure("Subtitle.TLabel",
                         background=self.COLORS["bg"],
                         foreground=self.COLORS["text_dim"],
                         font=("Segoe UI", 10))
        style.configure("Version.TLabel",
                         background=self.COLORS["bg"],
                         foreground=self.COLORS["card_alt"],
                         font=("Segoe UI", 9))
        style.configure("Status.TLabel",
                         background=self.COLORS["bg"],
                         foreground=self.COLORS["accent_ok"],
                         font=("Segoe UI", 10))
        style.configure("green.Horizontal.TProgressbar",
                         troughcolor=self.COLORS["input_bg"],
                         background=self.COLORS["accent_ok"],
                         thickness=8)

        self._running = False
        self._cancel_event = threading.Event()
        self.sn_row_col = (5, 3)
        self.img_row_col = (12, 0)
        self.auto_rotate_var = tk.BooleanVar(value=True)
        self.smart_rotate_var = tk.BooleanVar(value=True)

        # ── Milestone 2 State Variables ──
        self.start_table_var = tk.StringVar(value="1")
        self._total_tables = 0
        self.model_vars: dict[str, tk.BooleanVar] = {}
        self.detected_models: list[str] = []

        self._build_ui()

        # Variable traces for programmatic updates & automated tests
        self.docx_var.trace_add("write", lambda *_: self._on_docx_var_changed())
        self.sn_root_var.trace_add("write", lambda *_: self._on_sn_root_var_changed())

    def _build_ui(self):
        C = self.COLORS

        # ── Header ──
        header = ttk.Frame(self, style="BG.TFrame")
        header.pack(fill="x", padx=20, pady=(18, 5))

        title_row = tk.Frame(header, bg=C["bg"])
        title_row.pack(fill="x")
        ttk.Label(title_row, text=__app_name__, style="Title.TLabel").pack(side="left")
        ttk.Label(title_row, text=f"v{__version__}", style="Version.TLabel").pack(side="left", padx=(8, 0), anchor="s")

        ttk.Label(header, text="Insert device photos from SN folders into Unit Test Word document",
                  style="Subtitle.TLabel").pack(anchor="w", pady=(2, 0))

        # ── Input Cards ──
        inputs_frame = ttk.Frame(self, style="BG.TFrame")
        inputs_frame.pack(fill="x", padx=20, pady=10)

        # Card 1: DOCX file
        self.docx_var = tk.StringVar()
        self._make_file_card(inputs_frame, "Word Document (.docx)",
                             self.docx_var, self._pick_docx, icon="📄", with_config=True)

        # Card 2: SN folder root + keyword
        self.sn_root_var = tk.StringVar()
        self.keyword_var = tk.StringVar(value="WZP")
        self._make_sn_card(inputs_frame)

        # Card 3: Output path
        self.output_var = tk.StringVar()
        self._make_file_card(inputs_frame, "Output File (.docx)",
                             self.output_var, self._pick_output, icon="💾")

        # ── Button Row ──
        btn_frame = ttk.Frame(self, style="BG.TFrame")
        btn_frame.pack(fill="x", padx=20, pady=(5, 5))

        self.run_btn = tk.Button(
            btn_frame,
            text="▶  Run Injection",
            font=("Segoe UI", 12, "bold"),
            bg=C["btn_run"], fg=C["btn_run_fg"],
            activebackground="#00a885", activeforeground="#fff",
            relief="flat", cursor="hand2", padx=24, pady=8,
            command=self._on_run
        )
        self.run_btn.pack(side="left")

        self.stop_btn = tk.Button(
            btn_frame,
            text="■  Stop",
            font=("Segoe UI", 12, "bold"),
            bg=C["btn_stop"], fg=C["btn_stop_fg"],
            activebackground="#c0392b", activeforeground="#fff",
            relief="flat", cursor="hand2", padx=24, pady=8,
            command=self._on_stop,
            state="disabled"
        )
        self.stop_btn.pack(side="left", padx=(10, 0))

        opts_frame = ttk.Frame(btn_frame, style="BG.TFrame")
        opts_frame.pack(side="left", padx=10)

        tk.Checkbutton(
            opts_frame, text="Auto-Rotate (EXIF)", variable=self.auto_rotate_var,
            bg=C["bg"], fg=C["text"], selectcolor=C["card"], activebackground=C["bg"],
            activeforeground=C["text"], font=("Segoe UI", 9)
        ).pack(anchor="w")

        tk.Checkbutton(
            opts_frame, text="Smart Rotate (OpenThai OCR · CPU)", variable=self.smart_rotate_var,
            bg=C["bg"], fg="#a29bfe", selectcolor=C["card"], activebackground=C["bg"],
            activeforeground="#a29bfe", font=("Segoe UI", 9, "bold")
        ).pack(anchor="w")

        self.status_label = ttk.Label(btn_frame, text="Ready", style="Status.TLabel")
        self.status_label.pack(side="left", padx=15)

        # ── Progress Bar ──
        self.progress = ttk.Progressbar(self, style="green.Horizontal.TProgressbar",
                                         orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", padx=20, pady=(0, 5))

        # ── Log Area ──
        log_frame = tk.Frame(self, bg=C["card"], bd=0, highlightthickness=1,
                             highlightbackground=C["border"])
        log_frame.pack(fill="both", expand=True, padx=20, pady=(0, 18))

        self.log_text = tk.Text(
            log_frame,
            bg=C["input_bg"], fg=C["input_fg"],
            font=("Cascadia Code", 9),
            insertbackground=C["text"],
            selectbackground=C["accent"],
            relief="flat", wrap="word",
            padx=10, pady=8,
            state="disabled"
        )
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True)

        self.log_text.tag_configure("ok",       foreground="#00b894")
        self.log_text.tag_configure("err",      foreground="#e94560")
        self.log_text.tag_configure("info",     foreground="#8892b0")
        self.log_text.tag_configure("title",    foreground="#e0e0e0", font=("Cascadia Code", 9, "bold"))
        self.log_text.tag_configure("cancel",   foreground="#f39c12")

        self._auto_detect_paths()

    def _make_file_card(self, parent, label_text, var, command, icon="📄", with_config=False):
        C = self.COLORS
        card = tk.Frame(parent, bg=C["card"], bd=0, highlightthickness=1,
                        highlightbackground=C["border"])
        card.pack(fill="x", pady=3)

        inner = tk.Frame(card, bg=C["card"])
        inner.pack(fill="x", padx=12, pady=8)

        top = tk.Frame(inner, bg=C["card"])
        top.pack(fill="x")
        tk.Label(top, text=f"{icon}  {label_text}",
                 bg=C["card"], fg=C["text"],
                 font=("Segoe UI", 10, "bold")).pack(side="left")

        btn = tk.Button(top, text="Browse",
                        font=("Segoe UI", 9),
                        bg=C["btn_pick"], fg=C["btn_pick_fg"],
                        activebackground=C["card_alt"],
                        relief="flat", cursor="hand2",
                        padx=12, pady=2,
                        command=command)
        btn.pack(side="right")
        
        if with_config:
            self.config_btn = tk.Button(top, text="⚙ Configure Layout",
                                        font=("Segoe UI", 9),
                                        bg=C["border"], fg=C["text"],
                                        activebackground=C["card_alt"],
                                        relief="flat", cursor="hand2",
                                        padx=12, pady=2,
                                        command=self._configure_layout)
            self.config_btn.pack(side="right", padx=(0, 10))

        path_label = tk.Label(inner, textvariable=var,
                              bg=C["card"], fg=C["text_dim"],
                              font=("Consolas", 9), anchor="w")
        path_label.pack(fill="x", pady=(4, 0))
        
        if with_config:
            bottom_row = tk.Frame(inner, bg=C["card"])
            bottom_row.pack(fill="x", pady=(4, 0))

            self.layout_label = tk.Label(
                bottom_row,
                text=f"Layout: S/N at R{self.sn_row_col[0]+1}C{self.sn_row_col[1]+1} | Image at R{self.img_row_col[0]+1}C{self.img_row_col[1]+1}",
                bg=C["card"], fg=C["accent_ok"], font=("Segoe UI", 9, "italic"), anchor="w"
            )
            self.layout_label.pack(side="left")

            # Start Table / Page Selector
            start_frame = tk.Frame(bottom_row, bg=C["card"])
            start_frame.pack(side="right")

            tk.Label(
                start_frame, text="Start Table:",
                bg=C["card"], fg=C["text_dim"], font=("Segoe UI", 9)
            ).pack(side="left", padx=(0, 5))

            self.start_table_spinbox = tk.Spinbox(
                start_frame, from_=1, to=1,
                textvariable=self.start_table_var,
                width=5,
                bg=C.get("input_bg", "#0d1b2a"), fg=C.get("accent_ok", "#00b894"),
                buttonbackground=C.get("btn_pick", "#0f3460"),
                insertbackground=C.get("text", "#eaeaea"),
                relief="flat",
                highlightthickness=1,
                highlightbackground=C.get("border", "#233554"),
                highlightcolor=C.get("accent_ok", "#00b894"),
                font=("Consolas", 9, "bold"),
                command=self._on_start_table_changed
            )
            self.start_table_spinbox.pack(side="left")
            self.start_table_spinbox.bind("<FocusOut>", lambda e: self._on_start_table_changed())
            self.start_table_spinbox.bind("<Return>", lambda e: self._on_start_table_changed())

            self.total_tables_label = tk.Label(
                start_frame, text=" (of 1)",
                bg=C["card"], fg=C["text_dim"], font=("Segoe UI", 9)
            )
            self.total_tables_label.pack(side="left", padx=(3, 0))

    def _make_sn_card(self, parent):
        """Build the SN Folders card with an editable keyword field."""
        C = self.COLORS
        card = tk.Frame(parent, bg=C["card"], bd=0, highlightthickness=1,
                        highlightbackground=C["border"])
        card.pack(fill="x", pady=3)

        inner = tk.Frame(card, bg=C["card"])
        inner.pack(fill="x", padx=12, pady=8)

        # Top row: label + keyword entry + browse button
        top = tk.Frame(inner, bg=C["card"])
        top.pack(fill="x")

        tk.Label(top, text="📁  Photos Folders Directory(SN. or etc.)",
                 bg=C["card"], fg=C["text"],
                 font=("Segoe UI", 10, "bold")).pack(side="left")

        btn = tk.Button(top, text="Browse",
                        font=("Segoe UI", 9),
                        bg=C["btn_pick"], fg=C["btn_pick_fg"],
                        activebackground=C["card_alt"],
                        relief="flat", cursor="hand2",
                        padx=12, pady=2,
                        command=self._pick_sn_root)
        btn.pack(side="right")

        # Keyword entry inline
        kw_frame = tk.Frame(top, bg=C["card"])
        kw_frame.pack(side="right", padx=(0, 10))

        tk.Label(kw_frame, text="Keyword:",
                 bg=C["card"], fg=C["text_dim"],
                 font=("Segoe UI", 9)).pack(side="left")

        kw_entry = tk.Entry(kw_frame,
                            textvariable=self.keyword_var,
                            bg=C["input_bg"], fg=C["accent_ok"],
                            font=("Consolas", 10, "bold"),
                            insertbackground=C["text"],
                            relief="flat", width=12,
                            highlightthickness=1,
                            highlightbackground=C["border"],
                            highlightcolor=C["accent_ok"])
        kw_entry.pack(side="left", padx=(5, 0))

        tk.Label(kw_frame, text="*",
                 bg=C["card"], fg=C["accent_ok"],
                 font=("Consolas", 10, "bold")).pack(side="left")

        # Path display
        path_label = tk.Label(inner, textvariable=self.sn_root_var,
                              bg=C["card"], fg=C["text_dim"],
                              font=("Consolas", 9), anchor="w")
        path_label.pack(fill="x", pady=(4, 0))

        # ── Device Models Multi-Select Filter ──
        models_header = tk.Frame(inner, bg=C["card"])
        models_header.pack(fill="x", pady=(8, 2))

        tk.Label(
            models_header, text="🏷  Device Models Filter:",
            bg=C["card"], fg=C["text"], font=("Segoe UI", 9, "bold")
        ).pack(side="left")

        self.btn_select_all = tk.Button(
            models_header, text="Select All",
            font=("Segoe UI", 8),
            bg=C.get("card_alt", "#0f3460"), fg=C.get("text", "#eaeaea"),
            activebackground=C.get("btn_pick", "#0f3460"), activeforeground=C.get("text", "#eaeaea"),
            relief="flat", cursor="hand2", padx=6, pady=1,
            command=self._select_all_models,
            state="disabled"
        )
        self.btn_select_all.pack(side="right", padx=(4, 0))

        self.btn_deselect_all = tk.Button(
            models_header, text="Deselect All",
            font=("Segoe UI", 8),
            bg=C.get("card_alt", "#0f3460"), fg=C.get("text", "#eaeaea"),
            activebackground=C.get("btn_pick", "#0f3460"), activeforeground=C.get("text", "#eaeaea"),
            relief="flat", cursor="hand2", padx=6, pady=1,
            command=self._deselect_all_models,
            state="disabled"
        )
        self.btn_deselect_all.pack(side="right")

        self.models_count_label = tk.Label(
            models_header, text="",
            bg=C["card"], fg=C["text_dim"], font=("Segoe UI", 8, "italic")
        )
        self.models_count_label.pack(side="right", padx=(0, 8))

        # Dynamic checklist container
        self.models_container = tk.Frame(inner, bg=C["card"])
        self.models_container.pack(fill="x", pady=(2, 0))
        self.models_frame = self.models_container  # Alias for compatibility

        self.models_placeholder = tk.Label(
            self.models_container,
            text="(No models detected yet — select document or photo folder to scan)",
            bg=C["card"], fg=C["text_dim"], font=("Segoe UI", 9, "italic"), anchor="w"
        )
        self.models_placeholder.pack(anchor="w")

    def _auto_detect_paths(self):
        """Leave paths empty — user always selects via Browse."""
        pass

    def _get_start_table(self) -> int:
        """Safely parse start_table_var to a valid 1-based integer."""
        val = self.start_table_var.get().strip() if hasattr(self, "start_table_var") else "1"
        try:
            num = int(val)
            if num < 1:
                return 1
            if getattr(self, "_total_tables", 0) > 0 and num > self._total_tables:
                return self._total_tables
            return num
        except (ValueError, TypeError, OverflowError):
            return 1

    def _on_start_table_changed(self):
        """Clamp start_table_var upon focus out, enter key, or spin action."""
        num = self._get_start_table()
        if hasattr(self, "start_table_var"):
            self.start_table_var.set(str(num))

    def _update_docx_info(self, docx_path: str):
        """Read total tables from docx and update start_table spinbox limits."""
        if not docx_path or not os.path.isfile(docx_path):
            return
        try:
            from docx import Document
            doc = Document(docx_path)
            total = len(doc.tables)
            curr_start = self._get_start_table()
            self._total_tables = total
            if total > 0:
                if hasattr(self, "start_table_spinbox"):
                    self.start_table_spinbox.configure(to=total)
                if hasattr(self, "total_tables_label"):
                    self.total_tables_label.configure(text=f" (of {total})")
                if curr_start > total:
                    self.start_table_var.set(str(total))
            else:
                if hasattr(self, "start_table_spinbox"):
                    self.start_table_spinbox.configure(to=1)
                if hasattr(self, "total_tables_label"):
                    self.total_tables_label.configure(text=" (0 tables)")
                if hasattr(self, "start_table_var"):
                    self.start_table_var.set("1")
        except Exception:
            if hasattr(self, "total_tables_label"):
                self.total_tables_label.configure(text=" (error)")

    def _on_docx_var_changed(self):
        """Triggered automatically when docx_var changes."""
        path = self.docx_var.get().strip() if hasattr(self, "docx_var") else ""
        if path and os.path.isfile(path):
            self._update_docx_info(path)
            self._refresh_detected_models()

    def _on_sn_root_var_changed(self):
        """Triggered automatically when sn_root_var changes."""
        path = self.sn_root_var.get().strip() if hasattr(self, "sn_root_var") else ""
        if path and os.path.isdir(path):
            self._refresh_detected_models()

    def _refresh_detected_models(self):
        """Scan docx and folder to discover device models and update checklist."""
        docx_path = self.docx_var.get().strip() if hasattr(self, "docx_var") else ""
        sn_root = self.sn_root_var.get().strip() if hasattr(self, "sn_root_var") else ""

        docx_models = []
        if docx_path and os.path.isfile(docx_path):
            try:
                docx_models = detect_docx_models(docx_path, sn_row_col=self.sn_row_col)
            except Exception as e:
                self._log(f"Warning: Failed to detect models from docx: {e}")

        folder_models = []
        if sn_root and os.path.isdir(sn_root):
            try:
                folder_models = detect_folder_models(sn_root)
            except Exception as e:
                self._log(f"Warning: Failed to detect models from folders: {e}")

        # Merge docx and folder models, deduplicating synonyms (e.g. 7821 vs CP-7821-K9=)
        merged = list(docx_models)
        for fm in folder_models:
            if not any(model_matches(fm, dm) for dm in merged):
                merged.append(fm)

        self.detected_models = sorted(merged)
        self._update_models_ui(self.detected_models)

    def _update_models_ui(self, models: list[str]):
        """Render checkboxes for detected models."""
        if not hasattr(self, "models_container"):
            return

        for widget in self.models_container.winfo_children():
            widget.destroy()

        if not models:
            self.models_placeholder = tk.Label(
                self.models_container,
                text="(No models detected yet — select document or photo folder to scan)",
                bg=self.COLORS.get("card", "#16213e"), fg=self.COLORS.get("text_dim", "#8892b0"),
                font=("Segoe UI", 9, "italic"), anchor="w"
            )
            self.models_placeholder.pack(anchor="w")
            if hasattr(self, "models_count_label"):
                self.models_count_label.configure(text="")
            if hasattr(self, "btn_select_all"):
                self.btn_select_all.configure(state="disabled")
            if hasattr(self, "btn_deselect_all"):
                self.btn_deselect_all.configure(state="disabled")
            self.model_vars.clear()
            return

        if hasattr(self, "btn_select_all"):
            self.btn_select_all.configure(state="normal")
        if hasattr(self, "btn_deselect_all"):
            self.btn_deselect_all.configure(state="normal")

        new_vars = {}
        for idx, model in enumerate(models):
            prev_val = self.model_vars[model].get() if model in self.model_vars else True
            var = tk.BooleanVar(value=prev_val)
            new_vars[model] = var

            cb = tk.Checkbutton(
                self.models_container,
                text=model,
                variable=var,
                bg=self.COLORS.get("card", "#16213e"),
                fg=self.COLORS.get("text", "#eaeaea"),
                selectcolor=self.COLORS.get("input_bg", "#0d1b2a"),
                activebackground=self.COLORS.get("card", "#16213e"),
                activeforeground=self.COLORS.get("accent_ok", "#00b894"),
                font=("Segoe UI", 9),
                command=self._on_model_toggle
            )
            row = idx // 4
            col = idx % 4
            cb.grid(row=row, column=col, sticky="w", padx=(0, 15), pady=2)

        self.model_vars = new_vars
        self._on_model_toggle()

    def _on_model_toggle(self):
        """Update selected count label when any checkbox is toggled."""
        if not hasattr(self, "models_count_label"):
            return
        total = len(self.model_vars)
        selected = sum(1 for v in self.model_vars.values() if v.get())
        self.models_count_label.configure(text=f"({selected}/{total} selected)")

    def _select_all_models(self):
        """Select all detected models."""
        for var in self.model_vars.values():
            var.set(True)
        self._on_model_toggle()

    def _deselect_all_models(self):
        """Deselect all detected models."""
        for var in self.model_vars.values():
            var.set(False)
        self._on_model_toggle()

    def _pick_docx(self):
        path = filedialog.askopenfilename(
            title="Select Unit Test Word Document",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )
        if path:
            self.docx_var.set(path)
            self._update_docx_info(path)
            self._refresh_detected_models()
            self._configure_layout()

    def _configure_layout(self):
        docx = self.docx_var.get().strip()
        if not docx or not os.path.isfile(docx):
            messagebox.showerror("Error", "Please select a valid Word document (.docx) first.")
            return

        initial_idx = max(0, self._get_start_table() - 1)
        dialog = TableSelectionDialog(self, docx, self.sn_row_col, self.img_row_col, initial_table_idx=initial_idx)
        self.wait_window(dialog)

        if dialog.result:
            self.sn_row_col, self.img_row_col = dialog.result
            self.layout_label.configure(text=f"Layout: S/N at R{self.sn_row_col[0]+1}C{self.sn_row_col[1]+1} | Image at R{self.img_row_col[0]+1}C{self.img_row_col[1]+1}")
            self._log(f"Table layout configured: S/N={self.sn_row_col}, Image={self.img_row_col}")
            self._refresh_detected_models()

    def _pick_sn_root(self):
        path = filedialog.askdirectory(
            title="Select folder containing serial number subfolders"
        )
        if path:
            self.sn_root_var.set(path)
            self._refresh_detected_models()

    def _pick_output(self):
        path = filedialog.asksaveasfilename(
            title="Save output as",
            filetypes=[("Word Documents", "*.docx")],
            defaultextension=".docx",
            initialfile="unit_test_with_images.docx"
        )
        if path:
            self.output_var.set(path)

    def _log(self, msg):
        def _append():
            self.log_text.configure(state="normal")
            tag = None
            if msg.startswith("  [OK"):
                tag = "ok"
            elif "ERROR" in msg or "[ERR" in msg or "CRASH" in msg:
                tag = "err"
            elif "CANCELLED" in msg or "cancelled" in msg:
                tag = "cancel"
            elif msg.startswith("  [") or msg.startswith("    "):
                tag = "info"
            elif "═" in msg or "─" in msg:
                tag = "title"

            if tag:
                self.log_text.insert("end", msg + "\n", tag)
            else:
                self.log_text.insert("end", msg + "\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        self.after(0, _append)

    def _update_progress(self, current, total):
        def _set():
            self.progress["maximum"] = total
            self.progress["value"] = current
            pct = int(100 * current / total) if total > 0 else 0
            self.status_label.configure(text=f"Processing... {current}/{total} ({pct}%)")
        self.after(0, _set)

    def _set_running(self, running: bool):
        def _update():
            self._running = running
            if running:
                self.run_btn.configure(state="disabled", bg="#555")
                self.stop_btn.configure(state="normal")
            else:
                self.run_btn.configure(state="normal", bg=self.COLORS["btn_run"])
                self.stop_btn.configure(state="disabled")
        self.after(0, _update)

    def _on_stop(self):
        if self._running:
            self._cancel_event.set()
            self.status_label.configure(text="Stopping...", foreground="#f39c12")
            self.stop_btn.configure(state="disabled")

    def _on_run(self):
        if self._running:
            return

        docx = self.docx_var.get().strip()
        sn_root = self.sn_root_var.get().strip()
        output = self.output_var.get().strip()
        keyword = self.keyword_var.get().strip()

        if not docx or not os.path.isfile(docx):
            messagebox.showerror("Error", "Please select a valid Word document (.docx)")
            return
        if not sn_root or not os.path.isdir(sn_root):
            messagebox.showerror("Error", "Please select a valid SN folders directory")
            return
        if not output:
            messagebox.showerror("Error", "Please specify an output file path")
            return
        if not keyword:
            messagebox.showerror("Error", "Please enter a folder keyword (e.g. WZP)")
            return

        # ── Start Table Validation ──
        start_table = self._get_start_table()

        # ── Model Filtering Extraction ──
        selected_models = None
        if self.model_vars:
            checked = [m for m, var in self.model_vars.items() if var.get()]
            if not checked:
                messagebox.showerror("Error", "Please select at least one device model to inject.")
                return
            if len(checked) == len(self.model_vars):
                selected_models = None
            else:
                selected_models = checked

        # Clear
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")
        self.progress["value"] = 0
        self._cancel_event.clear()
        self._set_running(True)
        self.status_label.configure(text="Starting...", foreground=self.COLORS["accent_ok"])

        def worker():
            try:
                import pillow_heif
                pillow_heif.register_heif_opener()

                injected, img_count, no_folder, errors = run_injection(
                    docx, sn_root, output,
                    keyword=keyword,
                    sn_row_col=self.sn_row_col, img_row_col=self.img_row_col,
                    auto_rotate=self.auto_rotate_var.get(),
                    smart_rotate=self.smart_rotate_var.get(),
                    start_table=start_table,
                    selected_models=selected_models,
                    log_callback=self._log,
                    progress_callback=self._update_progress,
                    cancel_event=self._cancel_event
                )

                def finish():
                    self._set_running(False)
                    cancelled = self._cancel_event.is_set()
                    if cancelled:
                        self.status_label.configure(
                            text=f"Cancelled: {injected} tables saved before stop",
                            foreground="#f39c12")
                    elif errors:
                        self.status_label.configure(
                            text=f"Done: {injected} tables, {img_count} images, {len(errors)} error(s)",
                            foreground=self.COLORS["accent"])
                    else:
                        self.status_label.configure(
                            text=f"Done: {injected} tables, {img_count} images inserted",
                            foreground=self.COLORS["accent_ok"])
                    self.progress["value"] = self.progress["maximum"]
                self.after(0, finish)

            except Exception as e:
                import traceback
                self._log(f"\nCRASH: {e}")
                self._log(traceback.format_exc())
                def err_finish():
                    self._set_running(False)
                    self.status_label.configure(text=f"Error: {e}", foreground=self.COLORS["accent"])
                self.after(0, err_finish)

        t = threading.Thread(target=worker, daemon=True)
        t.start()


# ─────────────────────────────────────────────
# Entry Point
# ─────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()
