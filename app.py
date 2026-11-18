"""
Unit Test Image Injector v1.7.0-beta — Portable GUI App
====================================================
Inserts device photos from SN folders into the
Unit Test Word document, matching by Serial Number.

Supports: .jpg .jpeg .png .bmp .tiff .heic .heif
Auto-converts HEIC/HEIF (iPhone photos) to JPG.

Changelog:
  v1.7.0-beta — Verify photos: OCR every S/N / set folder and move misplaced photos to the right table
  v1.6.0 — Auto mode: match S/N anywhere in tables, nested folders, OCR S/N from unsorted photos
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

__version__ = "1.7.0-beta"
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


def _patch_docx_image_sha1_cache():
    """
    python-docx checks every existing image's SHA-1 on each add_picture() (duplicate check) and
    recomputes the hash every time: O(n^2) hashing of all image bytes. Image parts never change,
    so cache the hash per part. (On a 683-photo run this alone took 249 s.)
    """
    try:
        from docx.parts.image import ImagePart
    except Exception:
        return
    if getattr(ImagePart, "_sha1_cache_patched", False):
        return
    original = ImagePart.sha1

    def sha1(self):
        cached = self.__dict__.get("_cached_sha1")
        if cached is None:
            cached = original.fget(self)
            self.__dict__["_cached_sha1"] = cached
        return cached

    ImagePart.sha1 = property(sha1)
    ImagePart._sha1_cache_patched = True


INSERT_MAX_PX = 2000  # long side of photos placed in Word (14 cm wide -> ~360 dpi); 0 = keep original
_thread_local = threading.local()


def _worker_ocr_engine():
    """One OCR engine per worker thread (the shared engine runs one inference at a time)."""
    eng = getattr(_thread_local, "ocr_engine", None)
    if eng is None:
        backend = (os.environ.get("OCR_BACKEND") or "openthai").lower()
        if backend in ("openthai", "openthai2", "iapp"):
            from ocr_engine import OpenThaiEngine
            eng = OpenThaiEngine(threads=2)
        else:
            from ocr_engine import get_engine
            eng = get_engine()
        _thread_local.ocr_engine = eng
    return eng


def prepare_photo(img_path: str, auto_rotate: bool = True, smart_rotate: bool = False,
                  max_px: int = INSERT_MAX_PX) -> tuple[str, list[str]]:
    """
    Make one photo ready for Word in a single decode: HEIC -> JPG, EXIF rotation,
    Smart Rotate (OCR), shrink to max_px on the long side, save once as JPEG (dpi 72).
    Safe to run in worker threads. Returns (path to insert, log notes).
    The original file is returned untouched when nothing needs to change.
    """
    from PIL import Image as PILImage, ImageOps
    notes: list[str] = []
    src = img_path
    if os.path.splitext(src)[1].lower() in ('.heic', '.heif'):
        src = convert_heic_to_jpg(src)
        notes.append(f"Converted {os.path.basename(img_path)} -> JPG")
    if not (auto_rotate or smart_rotate or max_px):
        return src, notes

    with PILImage.open(src) as im:
        orig_size = im.size
        if max_px and im.format == "JPEG":
            im.draft("RGB", (max_px, max_px))  # decode at 1/2, 1/4 ... size: much faster for big photos
        orientation = im.getexif().get(0x0112, 1) if auto_rotate else 1
        work = ImageOps.exif_transpose(im) if auto_rotate else None
        work = (work if work is not None else im.copy()).convert("RGB")
    exif_rotated = orientation not in (1, None)
    if exif_rotated:
        notes.append(f"Auto-rotated (EXIF) {os.path.basename(img_path)}")

    angle = 0
    if smart_rotate:
        global _ocr_init_failed
        if not _ocr_init_failed:
            try:
                from ocr_engine import detect_orientation
                angle, score = detect_orientation(work, _worker_ocr_engine())
                if score < 15:
                    angle = 0
            except Exception as e:
                _ocr_init_failed = True
                notes.append(f"[OCR Disabled] Could not run OCR: {e}")
                angle = 0
        if angle:
            work = work.rotate(angle, expand=True)
            notes.append(f"Smart-rotated (OCR) {angle}° for {os.path.basename(img_path)}")

    too_big = bool(max_px) and max(orig_size) > max_px
    if not (exif_rotated or angle or too_big or src != img_path):
        return src, notes
    if too_big and max(work.size) > max_px:
        work.thumbnail((max_px, max_px), PILImage.LANCZOS)
    fd, tmp_path = tempfile.mkstemp(suffix='.jpg', prefix='prep_')
    os.close(fd)
    work.save(tmp_path, 'JPEG', quality=88, dpi=(72, 72))
    _temp_files.append(tmp_path)
    return tmp_path, notes


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


# ─────────────────────────────────────────────
# Auto Mode: locate S/N + image cell anywhere in a table
# ─────────────────────────────────────────────
PHOTO_LABEL_REGEX = re.compile(r"(รูป|ภาพ|photo|picture|image)", re.IGNORECASE)
_SN_SPLIT_REGEX = re.compile(r"[\s,;:|()\[\]{}<>\"'=]+")


def normalize_sn(text) -> str:
    """Uppercase and strip everything except A-Z / 0-9 (so 'wzp-301 000' == 'WZP301000')."""
    return re.sub(r"[^0-9A-Z]", "", str(text or "").upper())


def _is_sn_like(key: str, min_len: int) -> bool:
    return len(key) >= min_len and any(ch.isdigit() for ch in key)


SEQ_FOLDER_REGEX = re.compile(r"^\s*\d{1,4}\s*[-_/]\s*\d{1,4}\s*$")  # e.g. '195-280' (ชุดที่ 195/280)
# Stricter than CAMERA_SUBDIR_REGEX so serials like '21599990000GS0000208' are not mistaken for '100APPLE'
STRICT_CAMERA_DIR_REGEX = re.compile(r"^(?:dcim|camera|\d{3}[a-z_][a-z0-9_]{1,4})$", re.IGNORECASE)
_SN_VALUE_REGEX = re.compile(r"^[A-Za-z0-9][A-Za-z0-9\-/_.]{5,}$")


def find_sn_leaf_folders(root: str, log_cb=None) -> "SnFolderMap":
    """
    Auto-mode folder scan at any depth: every folder that directly holds images is a
    candidate serial folder named after itself, e.g. root/Model/By_SN/<SN>/img.jpg.
    - Camera subfolders (DCIM, 100APPLE, ...) take the name of their parent folder.
    - Sequence-number folders like '195-280' are ignored (not serials).
    - If the same name appears twice, the first folder (sorted path order) wins.
    sn_to_model maps each serial to the top-level folder under root.
    """
    sn_map: dict[str, list[str]] = {}
    sn_to_model: dict[str, str] = {}
    if not root or not os.path.isdir(root):
        return SnFolderMap(sn_map, sn_to_model=sn_to_model)
    root_abs = os.path.abspath(root)
    duplicates = []

    for dirpath, dirnames, filenames in os.walk(root_abs):
        dirnames[:] = sorted(d for d in dirnames if not d.startswith(('.', '__')))
        imgs = [os.path.join(dirpath, f) for f in sorted(filenames)
                if '_converted' not in f and os.path.splitext(f)[1].lower() in VALID_IMG_EXTS]
        if not imgs:
            continue
        rel_parts = os.path.relpath(dirpath, root_abs).split(os.sep)
        if rel_parts == ['.']:
            # images loose in the selected folder itself (named after it; resolved by OCR if not a serial)
            rel_parts = [os.path.basename(root_abs) or root_abs]
        parts = list(rel_parts)
        while parts and STRICT_CAMERA_DIR_REGEX.match(parts[-1]):
            parts.pop()
        if not parts:
            continue
        name = parts[-1]
        if SEQ_FOLDER_REGEX.match(name):
            continue
        if name in sn_map:
            if os.path.dirname(sn_map[name][0]) != dirpath:
                duplicates.append(name)
            else:
                sn_map[name].extend(imgs)
            continue
        sn_map[name] = imgs
        sn_to_model[name] = parts[0] if len(parts) > 1 else os.path.basename(root_abs)

    if duplicates and log_cb:
        log_cb(f"  [WARN] {len(duplicates)} duplicate folder name(s), using first: "
               + ", ".join(sorted(set(duplicates))[:5]))
    return SnFolderMap(sn_map, sn_to_model=sn_to_model)


def build_sn_index(sn_names) -> dict[str, str]:
    """
    Map normalized serial keys -> original folder name.
    Full folder names are indexed first; then S/N-looking tokens inside folder names
    (e.g. 'AP01 10299X000012' -> '10299X000012'). Keys claimed by >1 folder are dropped.
    """
    index: dict[str, str] = {}
    ambiguous: set[str] = set()

    def add(key, sn):
        if key in index and index[key] != sn:
            ambiguous.add(key)
        else:
            index[key] = sn

    for sn in sn_names:
        key = normalize_sn(sn)
        if _is_sn_like(key, 5):
            add(key, sn)
    full_keys = set(index)
    for sn in sn_names:
        for tok in re.split(r"[\s_()\[\]]+", str(sn)):
            key = normalize_sn(tok)
            if key not in full_keys and _is_sn_like(key, 6):
                add(key, sn)
    for key in ambiguous:
        index.pop(key, None)
    return index


def match_sns_in_text(text: str, sn_index: dict[str, str]) -> list[str]:
    """Return folder names whose serial appears in a cell's text (whole cell or any token)."""
    if not text or not text.strip():
        return []
    whole = normalize_sn(text)
    if whole in sn_index:
        return [sn_index[whole]]
    found = []
    for tok in _SN_SPLIT_REGEX.split(text):
        key = normalize_sn(tok)
        if key in sn_index and sn_index[key] not in found:
            found.append(sn_index[key])
    return found


def _iter_unique_cells(table):
    """Yield (row_idx, col_idx, cell) once per physical cell (merged cells reported once)."""
    seen = []
    for r_idx, row in enumerate(table.rows):
        try:
            cells = row.cells
        except Exception:
            continue
        for c_idx, cell in enumerate(cells):
            tc = cell._tc
            if any(tc is s for s in seen):
                continue
            seen.append(tc)
            yield r_idx, c_idx, cell


def _cell_has_drawing(cell) -> bool:
    try:
        return bool(cell._tc.xpath('.//*[local-name()="drawing" or local-name()="pict"]'))
    except Exception:
        return False


def detect_image_cell(table, sn_row: int, fallback_rc: tuple | None = None):
    """
    Guess the photo cell of a table, searching below the S/N row:
      1. first text-empty row merged into ONE cell (or already holding a picture)
      2. first text-empty row right after a short 'รูป/ภาพ/Photo/Picture' label row
      3. fallback_rc, only if that cell exists and is text-empty
    Returns ((row, col), reason) or (None, reason).
    """
    rows = {}
    for r, c, cell in _iter_unique_cells(table):
        rows.setdefault(r, []).append((c, cell))

    def row_empty(r):
        return r in rows and all(not cell.text.strip() for _, cell in rows[r])

    later = sorted(r for r in rows if r > sn_row)

    for r in later:
        if row_empty(r) and (len(rows[r]) == 1 or any(_cell_has_drawing(cell) for _, cell in rows[r])):
            return (r, rows[r][0][0]), "empty merged row"

    for r in later:
        label = " ".join(cell.text.strip() for _, cell in rows[r]).strip()
        if label and len(label) <= 40 and PHOTO_LABEL_REGEX.search(label) and row_empty(r + 1):
            return (r + 1, rows[r + 1][0][0]), f"below label '{label}'"

    if fallback_rc is not None:
        try:
            cell = table.rows[fallback_rc[0]].cells[fallback_rc[1]]
            if not cell.text.strip():
                return tuple(fallback_rc), "configured layout"
        except (IndexError, AttributeError, TypeError):
            pass
    return None, "no empty photo cell found"


def auto_locate_in_table(table, sn_index: dict[str, str], fallback_img_rc: tuple | None = None) -> dict | None:
    """
    Scan every cell of a table for a known folder serial.
    Returns None if no serial is found, otherwise a dict with:
      status: 'ok' | 'ambiguous' | 'no_image_cell'
      sn, sn_rc, img_rc, reason (for 'ok'); sns (for 'ambiguous')
    """
    matches: dict[str, tuple] = {}
    for r, c, cell in _iter_unique_cells(table):
        for sn in match_sns_in_text(cell.text, sn_index):
            matches.setdefault(sn, (r, c))
    if not matches:
        return None
    if len(matches) > 1:
        return {"status": "ambiguous", "sns": sorted(matches)}
    sn, sn_rc = next(iter(matches.items()))
    # A list/summary table has many other serial-like values in the same column
    siblings = 0
    for r, c, cell in _iter_unique_cells(table):
        if c == sn_rc[1] and r != sn_rc[0]:
            txt = cell.text.strip()
            if _SN_VALUE_REGEX.match(txt) and any(ch.isdigit() for ch in txt):
                siblings += 1
    if siblings >= 3:
        return {"status": "ambiguous", "sns": [sn], "list_rows": siblings + 1}
    img_rc, reason = detect_image_cell(table, sn_rc[0], fallback_img_rc)
    if img_rc is None:
        return {"status": "no_image_cell", "sn": sn, "sn_rc": sn_rc, "reason": reason}
    return {"status": "ok", "sn": sn, "sn_rc": sn_rc, "img_rc": img_rc, "reason": reason}


# ─────────────────────────────────────────────
# Auto Mode + OCR: photos not sorted into S/N folders
# ─────────────────────────────────────────────
SERIAL_LABEL_REGEX = re.compile(r"^\s*(?:serial\s*(?:no\.?|number)?|s\s*/\s*n)\s*[:.]?\s*$", re.IGNORECASE)
SERIAL_INLINE_REGEX = re.compile(r"^\s*(?:serial\s*(?:no\.?|number)?|s\s*/\s*n)\s*[:.]\s*(\S+)\s*$", re.IGNORECASE)
# 'S/N' as printed, plus common OCR misreads of the S ('5/N', '3/N', '$N'), only at a word start
_OCR_SN_PREFIX_REGEX = re.compile(r"(?<![A-Za-z0-9])[S5$3]\s*[/|]?\s*N\W*(.+)$", re.IGNORECASE)


def extract_doc_serials(doc) -> dict[str, str]:
    """
    Serial numbers written in the document, found by their label:
      'Serial Number:' | <value>   (label cell, value in the next cell)
      'S/N: <value>'               (label and value in one cell)
      header 'Serial Number' column -> values in the rows below
    Returns {normalized: original text}.
    """
    out: dict[str, str] = {}

    def add(val):
        key = normalize_sn(val)
        if _is_sn_like(key, 6):
            out.setdefault(key, val.strip())

    for table in doc.tables:
        rows: dict[int, list] = {}
        for r, c, cell in _iter_unique_cells(table):
            rows.setdefault(r, []).append((c, cell.text.strip()))
        header_cols = set()
        for r in sorted(rows):
            cells = rows[r]
            for i, (c, txt) in enumerate(cells):
                if SERIAL_LABEL_REGEX.match(txt):
                    header_cols.add(c)
                    if i + 1 < len(cells) and cells[i + 1][1]:
                        add(cells[i + 1][1])
                    continue
                m = SERIAL_INLINE_REGEX.match(txt)
                if m:
                    add(m.group(1))
                elif c in header_cols and txt and _SN_VALUE_REGEX.match(txt):
                    add(txt)
    return out


# Characters OCR typically confuses on labels; swapping within a group costs 1/3 of an edit.
_OCR_CONFUSABLE_GROUPS = ("O0DQ", "I1L7T", "S5", "B8", "Z2", "G6", "UV")
_OCR_CONFUSABLE = {(a, b) for g in _OCR_CONFUSABLE_GROUPS for a in g for b in g if a != b}
_CONFUSABLE_COST = 1 / 3


def _levenshtein(a: str, b: str, max_dist: float = 2.0) -> float:
    """Edit distance where OCR look-alike substitutions (O/0, 7/1, 8/B, ...) cost 1/3."""
    if abs(len(a) - len(b)) > max_dist:
        return max_dist + 1
    prev = [float(j) for j in range(len(b) + 1)]
    for i, ca in enumerate(a, 1):
        cur = [float(i)]
        for j, cb in enumerate(b, 1):
            if ca == cb:
                sub = 0.0
            elif (ca, cb) in _OCR_CONFUSABLE:
                sub = _CONFUSABLE_COST
            else:
                sub = 1.0
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + sub))
        if min(cur) > max_dist:
            return max_dist + 1
        prev = cur
    return prev[-1]


_SERIAL_FOLDER_TOKEN_REGEX = re.compile(r"^[A-Za-z0-9][A-Za-z0-9\-_.]*$")


def _folder_named_like_other_serial(name: str, doc_serials) -> bool:
    """
    True when a folder name is itself a serial number of a device that is NOT in this document
    (e.g. a folder of every device of the project, used with a document of only some of them).
    OCR-ing such folders cannot find a document serial and costs seconds per photo, so they are skipped.
    A name close to a document serial (typo, distance <= 2) is not skipped: OCR may still rescue it.
    """
    raw = os.path.basename(str(name)).strip()
    key = normalize_sn(raw)
    if not doc_serials or key in doc_serials or not _SERIAL_FOLDER_TOKEN_REGEX.match(raw):
        return False
    if SEQ_FOLDER_REGEX.match(raw) or STRICT_CAMERA_DIR_REGEX.match(raw):
        return False
    letters = sum(ch.isalpha() for ch in key)
    digits = sum(ch.isdigit() for ch in key)
    if letters < 1 or digits < 3:
        return False
    lens = [len(k) for k in doc_serials]
    if not (min(lens) - 2 <= len(key) <= max(lens) + 2):
        return False
    return all(_levenshtein(key, k) > 2 for k in doc_serials)


def find_serial_in_ocr_lines(lines: list[str], doc_serials: dict[str, str]) -> tuple[str | None, str]:
    """
    Pick the document serial printed in a photo from its OCR lines.
    Exact match first (text after 'S/N' preferred), then a unique near match
    (distance <= 2, look-alike characters cheap, serials of 10+ chars) to absorb OCR mistakes.
    Returns (normalized serial or None, 'exact' | 'fuzzyN.N' | '').
    """
    cands = []
    for text in lines:
        m = _OCR_SN_PREFIX_REGEX.search(text)
        if m:
            cands.insert(0, normalize_sn(m.group(1)))
        cands.append(normalize_sn(text))
        cands += [normalize_sn(t) for t in _SN_SPLIT_REGEX.split(text)]
    cands = [c for c in dict.fromkeys(cands) if len(c) >= 6]
    for c in cands:
        if c in doc_serials:
            return c, "exact"
    best: dict[str, float] = {}
    for c in cands:
        if len(c) < 10:
            continue
        for s in doc_serials:
            if len(s) < 10:
                continue
            limit = 2.5 if len(s) >= 12 else 2.0
            d = _levenshtein(c, s, max_dist=3.0)
            if d <= limit and d < best.get(s, 99):
                best[s] = d
    if best:
        ranked = sorted(best.items(), key=lambda kv: kv[1])
        # must clearly beat the next candidate (sibling serials like ...042 / ...073 stay apart)
        margin = 0.5 if ranked[0][1] > 2.0 else 1e-9
        if len(ranked) == 1 or ranked[0][1] < ranked[1][1] - margin:
            return ranked[0][0], f"fuzzy{ranked[0][1]:.1f}"
    return None, ""


def _load_photo_for_ocr(img_path: str):
    from PIL import Image as PILImage, ImageOps
    with PILImage.open(img_path) as img:
        img = ImageOps.exif_transpose(img).convert("RGB")
    img.thumbnail((4096, 4096))
    return img


def _ocr_image_lines(img_path: str, engine) -> list[str]:
    return engine.read(_load_photo_for_ocr(img_path)).texts(include_dropped=True)


def _serial_like_line(text: str) -> bool:
    if _OCR_SN_PREFIX_REGEX.search(text):
        return True
    return any(len(normalize_sn(t)) >= 8 and any(ch.isdigit() for ch in t) for t in _SN_SPLIT_REGEX.split(text))


_STICKER_REGEX = re.compile(r"(?<![\d/])(\d{1,4})\s*/\s*(\d{1,4})(?![\d/])")
# look-alikes seen when OCR reads big serif sticker digits (Thai ด/ต/ฉ for a wide 0)
_STICKER_FIX = str.maketrans({"O": "0", "o": "0", "D": "0", "Q": "0", "ด": "0", "ต": "0", "ฉ": "0",
                              "I": "1", "l": "1", "|": "1"})


def build_set_sticker_lookup(entries: list[dict]) -> dict[int, list[dict]]:
    """
    For set-number stickers on boxes ('31/100'): {total: [ {n: doc entry}, ... ]}.
    A device group is listed under its whole total (numbered across sections, e.g. 1..280)
    and, when it is split into sections, under each section's own total (e.g. 1..51).
    """
    groups: dict[str, dict] = {}
    for e in entries:
        g = groups.setdefault(e["group"], {"total": e["group_total"], "by_index": {}, "by_section": {}})
        g["by_index"][e["index"]] = e
        g["by_section"].setdefault(e["total"], {}).setdefault(e["n"], []).append(e)
    lookup: dict[int, list[dict]] = {}
    for g in groups.values():
        lookup.setdefault(g["total"], []).append(g["by_index"])
        for sec_total, by_n in g["by_section"].items():
            if sec_total != g["total"] and len(by_n) == sec_total and all(len(v) == 1 for v in by_n.values()):
                lookup.setdefault(sec_total, []).append({n: v[0] for n, v in by_n.items()})
    return lookup


def serial_from_set_sticker(lines: list[str], lookup: dict[int, list[dict]],
                            doc_serials: dict[str, str]) -> tuple[str | None, str]:
    """
    Read a set-number sticker like '31/100' (a line that is just 'N/T') and return the serial of
    the document table 'ชุดที่ N/T'. Only answers when exactly one device fits (e.g. '1/2' with two
    2-unit device groups is ambiguous -> None).
    Returns (normalized serial or None, 'N/T').
    """
    hits: dict[str, str] = {}
    for text in lines:
        fixed = text.strip().translate(_STICKER_FIX)
        for m in _STICKER_REGEX.finditer(fixed):
            if len(fixed.replace(" ", "")) > len(m.group(0).replace(" ", "")) + 3:
                continue  # part of a longer text, not a sticker
            n, total = int(m.group(1)), int(m.group(2))
            if not 1 <= n <= total:
                continue
            for mapping in lookup.get(total, []):
                e = mapping.get(n)
                key = normalize_sn(e.get("sn")) if e else ""
                if key and key in doc_serials:
                    hits.setdefault(key, f"{n}/{total}")
    if len(hits) == 1:
        return next(iter(hits.items()))
    return None, ""


def read_serial_from_photo(img_path: str, engine, doc_serials: dict[str, str],
                           max_zoom_lines: int = 6,
                           sticker_lookup: dict[int, list[dict]] | None = None) -> tuple[str | None, str]:
    """
    OCR a photo for a document serial:
      1. the S/N label (whole photo; if none, each 'S/N'/serial-looking line is cropped and
         enlarged to ~48/72 px text and read again),
      2. else a set-number sticker on the box ('31/100' -> document 'ชุดที่ 31/100' -> its S/N).
    When both are readable and disagree, the S/N wins and 'how' notes the sticker mismatch.
    Returns (normalized serial or None, how): 'exact' / 'fuzzyN.N' (+'+zoom') / 'sticker N/T'.
    """
    from PIL import Image as PILImage
    img = _load_photo_for_ocr(img_path)
    res = engine.read(img)
    texts = res.texts(include_dropped=True)
    sticker_key, sticker_label = (_read_foreground_sticker(img, res, engine, sticker_lookup, doc_serials,
                                                           allow_zoom=False)
                                  if sticker_lookup else (None, ""))
    key, how = find_serial_in_ocr_lines(texts, doc_serials)

    if not key:
        lines = [ln for ln in getattr(res, "lines", []) if getattr(ln, "box", None) and _serial_like_line(ln.text)]
        lines.sort(key=lambda ln: 0 if _OCR_SN_PREFIX_REGEX.search(ln.text) else 1)
        for ln in lines[:max_zoom_lines]:
            xs = [pt[0] for pt in ln.box]
            ys = [pt[1] for pt in ln.box]
            h = max(ys) - min(ys)
            if h <= 0:
                continue
            crop = img.crop((int(max(0, min(xs) - 2 * h)), int(max(0, min(ys) - 1.5 * h)),
                             int(min(img.width, max(xs) + 2 * h)), int(min(img.height, max(ys) + 1.5 * h))))
            tried = set()
            for target_px in (48.0, 72.0):
                scale = round(min(8.0, max(2.0, target_px / h)), 1)
                if scale in tried:
                    continue
                tried.add(scale)
                big = crop.resize((max(1, int(crop.width * scale)), max(1, int(crop.height * scale))),
                                  PILImage.LANCZOS)
                key, how = find_serial_in_ocr_lines(engine.read(big).texts(include_dropped=True), doc_serials)
                if key:
                    how += "+zoom"
                    break
            if key:
                break

    if key:
        if sticker_key and sticker_key != key:
            how += f" [sticker {sticker_label} says {doc_serials[sticker_key]}]"
        return key, how
    if sticker_key:
        return sticker_key, f"sticker {sticker_label}"
    if sticker_lookup:
        sticker_key, sticker_label = _read_foreground_sticker(img, res, engine, sticker_lookup, doc_serials)
        if sticker_key:
            return sticker_key, f"sticker {sticker_label}"
    return None, ""


def _foreground_sticker_line(img, res):
    """
    The set-number sticker of the box being photographed: among short OCR lines containing '/',
    the one with the best size x closeness-to-centre score. Boxes stacked in the background carry
    stickers near the photo edges ('271/280' behind '71/280'); they must never decide.
    """
    best, best_score = None, 0.0
    cx, cy = img.width / 2, img.height / 2
    for ln in getattr(res, "lines", []):
        box = getattr(ln, "box", None)
        text = ln.text.strip()
        # digits not required: big serif stickers are sometimes read as e.g. '/Y Y'
        if not box or "/" not in text or len(text) > 12:
            continue
        xs = [p[0] for p in box]
        ys = [p[1] for p in box]
        h = max(ys) - min(ys)
        mx, my = (min(xs) + max(xs)) / 2, (min(ys) + max(ys)) / 2
        dist = (((mx - cx) / cx) ** 2 + ((my - cy) / cy) ** 2) ** 0.5   # 0 centre .. ~1.4 corner
        touches_edge = min(xs) < 0.01 * img.width or max(xs) > 0.99 * img.width
        score = h * max(0.05, 1.0 - 0.8 * dist) * (0.3 if touches_edge else 1.0)
        if score > best_score:
            best, best_score = ln, score
    return best


def _read_foreground_sticker(img, res, engine, sticker_lookup, doc_serials, allow_zoom=True):
    """
    Read only the foreground sticker: its OCR text first; if that is not a usable 'N/T', crop it
    and read 9 variants (text height 40/48/56 px x width squeezed to 50/60/70%) - big bold serif
    digits are hard for the OCR model as photographed. A zoomed result needs >= 2 agreeing
    variants and no tie. Without boxes (simple engines) the plain text lines are used.
    Returns (normalized serial or None, 'N/T' or 'N/T+zoom').
    """
    from PIL import Image as PILImage, ImageOps
    from collections import Counter
    ln = _foreground_sticker_line(img, res)
    if ln is None:
        if getattr(res, "lines", None):
            return None, ""
        return serial_from_set_sticker(res.texts(include_dropped=True), sticker_lookup, doc_serials)
    key, label = serial_from_set_sticker([ln.text], sticker_lookup, doc_serials)
    if key or not allow_zoom:
        return key, label

    xs = [p[0] for p in ln.box]
    ys = [p[1] for p in ln.box]
    h = max(ys) - min(ys)
    if h < 6:
        return None, ""
    crop = img.crop((int(max(0, min(xs) - 1.5 * h)), int(max(0, min(ys) - 0.5 * h)),
                     int(min(img.width, max(xs) + 1.5 * h)), int(min(img.height, max(ys) + 0.5 * h))))
    gray = ImageOps.autocontrast(crop.convert("L"))
    votes: Counter = Counter()
    labels = {}
    for target in (40.0, 48.0, 56.0):
        k = target / h
        for squeeze in (0.5, 0.6, 0.7):
            v = gray.resize((max(1, int(gray.width * k * squeeze)), max(1, int(gray.height * k))),
                            PILImage.LANCZOS)
            v = ImageOps.expand(v.convert("RGB"), border=24, fill="white")
            vk, vl = serial_from_set_sticker(engine.read(v).texts(include_dropped=True),
                                             sticker_lookup, doc_serials)
            if vk:
                votes[vk] += 1
                labels[vk] = vl
    if votes:
        ranked = votes.most_common(2)
        if ranked[0][1] >= 2 and (len(ranked) == 1 or ranked[0][1] > ranked[1][1]):
            return ranked[0][0], labels[ranked[0][0]] + "+zoom"
    return None, ""


def photo_capture_time(img_path: str) -> float | None:
    """Capture time from EXIF (DateTimeOriginal + sub-seconds, else DateTime), or None."""
    from datetime import datetime
    from PIL import Image as PILImage
    try:
        with PILImage.open(img_path) as img:
            exif = img.getexif()
            sub = exif.get_ifd(0x8769)
            raw = sub.get(36867) or exif.get(306)
            subsec = sub.get(37521) or sub.get(37520) or ""
        if not raw:
            return None
        t = datetime.strptime(str(raw).strip("\x00 ")[:19], "%Y:%m:%d %H:%M:%S").timestamp()
        digits = "".join(ch for ch in str(subsec) if ch.isdigit())
        return t + (float("0." + digits) if digits else 0.0)
    except Exception:
        return None


UNREAD_MAX_GAP_S = 90.0  # an unread photo must be shot within this many seconds of its device's chain


def assign_unread_photos(seq: list[tuple[str, str | None]], times: list[float] | None = None,
                         max_gap: float = UNREAD_MAX_GAP_S) -> tuple[dict, str]:
    """
    seq: photos in shooting order as (path, serial-or-None). Each run of unread photos
    lying between serial A and serial B is split between them:
      - with capture times: at the largest time gap (a device's photos are taken close together),
        so it works whether the label photo was shot before or after the device photos.
        A photo is only attached through a chain of gaps <= max_gap; photos shot long before/after
        any readable label (e.g. another device whose label can't be read) stay unassigned;
      - without times: each serial is filled up to the usual photo count (A first, then B);
        photos beyond that stay unassigned.
    Returns ({serial: [paths in shooting order], None: [unassigned paths]}, method).
    """
    n = len(seq)
    owner: list = [k for _, k in seq]
    read_idx = [i for i, k in enumerate(owner) if k]
    if not read_idx:
        return ({None: [p for p, _ in seq]} if seq else {}), "none"
    typical = max(1, round(n / len({owner[i] for i in read_idx})))
    method = "time gap" if times else "photo count"

    runs = []
    i = 0
    while i < n:
        if owner[i] is None:
            j = i
            while j < n and owner[j] is None:
                j += 1
            runs.append((i, j))
            i = j
        else:
            i += 1

    for i, j in runs:
        a = owner[i - 1] if i > 0 else None
        b = owner[j] if j < n else None
        k = j - i
        if a is None:
            split = 0
        elif b is None or b == a:
            split = k
        elif times:
            gaps = [times[x + 1] - times[x] for x in range(i - 1, j)]  # k+1 gaps
            split = max(range(k + 1), key=lambda g: (gaps[g], -g))
        else:
            split = k
        left, right = list(range(i, i + split)), list(range(i + split, j))

        if times:
            keep_left, prev = [], i - 1
            for x in left:
                if times[x] - times[prev] > max_gap:
                    break
                keep_left.append(x)
                prev = x
            keep_right, nxt = [], j
            for x in reversed(right):
                if times[nxt] - times[x] > max_gap:
                    break
                keep_right.append(x)
                nxt = x
        else:
            if a is not None and b is not None and b != a:
                have_a = sum(1 for x in range(i) if owner[x] == a)
                split = min(k, max(0, typical - have_a))
                left, right = list(range(i, i + split)), list(range(i + split, j))
            elif a is not None:
                have_a = sum(1 for x in range(i) if owner[x] == a)
                if b != a:
                    left = left[:max(0, typical - have_a)]
            keep_left = left
            keep_right = []
            if b is not None and b != a:
                b_reads = 0
                x = j
                while x < n and owner[x] == b:
                    b_reads += 1
                    x += 1
                cap = max(0, typical - b_reads)
                keep_right = right[len(right) - cap:] if cap else []

        for x in keep_left:
            owner[x] = a
        for x in keep_right:
            owner[x] = b

    groups: dict = {}
    for (path, _), k in zip(seq, owner):
        groups.setdefault(k, []).append(path)
    return groups, method


def group_photos_by_ocr(images: list[str], doc_serials: dict[str, str], engine=None,
                        log_cb=None, cancel_event=None, time_fn=None,
                        sticker_lookup: dict[int, list[dict]] | None = None
                        ) -> tuple[dict[str, list[str]], list[str]]:
    """
    OCR each photo for a document serial, then group the photos per serial.
    Photos are put in shooting order (EXIF capture time when every photo has one, else
    file-name order); photos with no readable serial (e.g. wide shots of the device) are
    given to the serial shot closest to them (see assign_unread_photos).
    Returns ({original serial: [image paths]}, [unassigned image paths]).
    """
    def log(msg):
        if log_cb:
            log_cb(msg)

    if engine is None:
        from ocr_engine import get_engine
        engine = get_engine(log_cb=log_cb)
    time_fn = time_fn or photo_capture_time

    reads: list[tuple[str, str | None]] = []
    sticker_reads: dict[str, str] = {}  # path -> 'N/T' for photos identified only by a set sticker
    total = len(images)
    for i, path in enumerate(images, start=1):
        if cancel_event is not None and cancel_event.is_set():
            break
        try:
            key, how = read_serial_from_photo(path, engine, doc_serials, sticker_lookup=sticker_lookup)
        except Exception as e:
            key, how = None, ""
            log(f"      [OCR ] {os.path.basename(path)}: cannot read ({e})")
        reads.append((path, key))
        if key and how.startswith("sticker"):
            sticker_reads[path] = how.split()[1].replace("+zoom", "")
        if key and "[sticker" in how:
            log(f"    [WARN] {os.path.basename(path)}: label S/N {doc_serials[key]} but {how.split('[', 1)[1].rstrip(']')}"
                " - check the sticker (S/N label used)")
        elif key and how != "exact":
            log(f"      [OCR ] {os.path.basename(path)} -> {doc_serials[key]} ({how})")
        if i % 20 == 0 or i == total:
            log(f"    OCR {i}/{total} photo(s)...")

    # Shooting order
    stamps = [time_fn(p) for p, _ in reads]
    times = None
    if reads and all(t is not None for t in stamps):
        order = sorted(range(len(reads)), key=lambda i: (stamps[i], os.path.basename(reads[i][0])))
        reads = [reads[i] for i in order]
        times = [stamps[i] for i in order]
        log("    Shooting order: EXIF capture time")
    else:
        missing = sum(1 for t in stamps if t is None)
        if reads:
            log(f"    Shooting order: file name ({missing} photo(s) without EXIF time)")

    # A sticker read is weaker than an S/N label: if photos shot close by (within the 90 s chain)
    # have readable S/N labels, the sticker must agree with one of them, else it is ignored.
    if sticker_reads:
        keys = [k for _, k in reads]
        dropped = 0
        for i, (path, k) in enumerate(reads):
            if path not in sticker_reads or not k:
                continue
            strong = set()
            for step in (-1, 1):
                j = i
                while 0 <= j + step < len(reads):
                    nj = j + step
                    if times is not None and abs(times[nj] - times[j]) > UNREAD_MAX_GAP_S:
                        break
                    if keys[nj] and reads[nj][0] not in sticker_reads:
                        strong.add(keys[nj])
                        break
                    j = nj
            if strong and k not in strong:
                reads[i] = (path, None)
                dropped += 1
                log(f"      [OCR ] {os.path.basename(path)}: sticker {sticker_reads[path]} ignored "
                    f"(nearby label S/N {', '.join(doc_serials[s] for s in sorted(strong))})")
        used = len(sticker_reads) - dropped
        log(f"    {used} photo(s) identified by set-number sticker"
            + (f", {dropped} sticker read(s) ignored (disagree with nearby S/N labels)" if dropped else ""))

    groups, method = assign_unread_photos(reads, times)
    unassigned = groups.pop(None, [])
    n_unread = sum(1 for _, k in reads if not k)
    if n_unread - len(unassigned) > 0:
        log(f"    {n_unread - len(unassigned)} photo(s) without readable S/N assigned by {method}")

    # Same serial read in two separate places of the sequence -> warn (photos still grouped)
    seen_runs = []
    last = None
    for _, key in reads:
        if key and key != last:
            seen_runs.append(key)
            last = key
    repeated = sorted({k for k in seen_runs if seen_runs.count(k) > 1})
    if repeated:
        log(f"    [WARN] serial read in non-adjacent photos, please check: "
            + ", ".join(doc_serials[k] for k in repeated[:5]))

    return {doc_serials[k]: v for k, v in groups.items()}, unassigned


# ─────────────────────────────────────────────
# Verify photos: OCR every S/N / set folder and catch photos of another device
# ─────────────────────────────────────────────
VERIFY_MAX_FUZZY = 1.0  # a label read may move a photo only when exact or this close (fuzzy distance)


def _strong_label_read(how: str, max_fuzzy: float = VERIFY_MAX_FUZZY) -> bool:
    """True for an S/N label read trusted enough to move a photo: 'exact' or 'fuzzyN.N' with N.N <= max_fuzzy.
    Sticker reads ('sticker N/T') never count."""
    if not how or how.startswith("sticker"):
        return False
    m = re.match(r"^(exact|fuzzy(\d+(?:\.\d+)?))", how)
    if not m:
        return False
    return m.group(1) == "exact" or float(m.group(2)) <= max_fuzzy + 1e-9


VERIFY_OWN_MISREAD = 1.5   # a read this close to the folder's own serial is taken as a misread of its own label
VERIFY_LOOKALIKE = 1.0     # another document serial closer than this (look-alikes only) makes a read ambiguous


def _move_read_doubt(key: str, own: str, doc_serials: dict[str, str], how: str = "fuzzy") -> str:
    """Why a label read of another device must not move a photo, or '' if it may.
    Real cases: the folder's own 'NJ6' read as 'NI6' -> fuzzy match 'NTG'; 'P6H' read exactly as 'PGH',
    also a real serial. An exact read differing from the folder's serial by a real character is trusted."""
    if own and not how.startswith("exact") and _levenshtein(key, own, max_dist=VERIFY_OWN_MISREAD) <= VERIFY_OWN_MISREAD:
        return f"close to the folder's own S/N {doc_serials.get(own, own)}"
    for s in doc_serials:
        if s != key and abs(len(s) - len(key)) <= 1 and _levenshtein(key, s, max_dist=1.0) < VERIFY_LOOKALIKE:
            return f"look-alike of {doc_serials[s]}"
    return ""


def _lookalike_rivals(key: str, doc_serials: dict[str, str]) -> set:
    """Other document serials that differ from key only by OCR look-alikes (distance < VERIFY_LOOKALIKE)."""
    return {s for s in doc_serials if s != key and abs(len(s) - len(key)) <= 1
            and _levenshtein(key, s, max_dist=1.0) < VERIFY_LOOKALIKE}


CONFIRM_HEIGHTS = (40.0, 56.0, 72.0)   # text heights (px) of the zoomed re-reads
CONFIRM_MIN_AGREE = 5                  # of 6 variants (3 heights x plain / contrast+sharpen)


def _label_tokens(lines: list[str]) -> set:
    toks = set()
    for text in lines:
        m = _OCR_SN_PREFIX_REGEX.search(text)
        if m:
            toks.add(normalize_sn(m.group(1)))
        toks.add(normalize_sn(text))
        toks.update(normalize_sn(t) for t in _SN_SPLIT_REGEX.split(text))
    return {t for t in toks if len(t) >= 6}


def _confirm_label_read(img_path: str, engine, key: str, rivals: set, doc_serials: dict[str, str]) -> tuple[bool, str]:
    """
    Re-read a label whose serial has look-alike rivals in the document (e.g. 'PGH' vs 'P6H'):
    the label line is cropped and read in 6 variants (text height 40/56/72 px x plain / autocontrast+sharpen).
    Confirmed only when >= CONFIRM_MIN_AGREE variants read exactly key and no variant reads a rival,
    another serial, or a token at a rival's position (a single '6' vetoes 'G').
    Returns (confirmed, detail).
    """
    from PIL import Image as PILImage, ImageOps, ImageFilter
    img = _load_photo_for_ocr(img_path)
    res = engine.read(img)
    lines = [ln for ln in getattr(res, "lines", []) if getattr(ln, "box", None)]
    target = [ln for ln in lines if find_serial_in_ocr_lines([ln.text], doc_serials)[0] == key]
    if not target:
        target = sorted((ln for ln in lines if _serial_like_line(ln.text)),
                        key=lambda ln: 0 if _OCR_SN_PREFIX_REGEX.search(ln.text) else 1)[:1]
    if target:
        ln = target[0]
        xs = [pt[0] for pt in ln.box]
        ys = [pt[1] for pt in ln.box]
        h = max(ys) - min(ys)
        if h <= 0:
            return False, "label line has no height"
        x0, y0 = int(max(0, min(xs) - 2 * h)), int(max(0, min(ys) - 1.5 * h))
        x1, y1 = int(min(img.width, max(xs) + 2 * h)), int(min(img.height, max(ys) + 1.5 * h))
        if x1 <= x0 or y1 <= y0:
            return False, "label line outside the photo"
        base = img.crop((x0, y0, x1, y1))
    else:   # engines without line boxes: re-read the whole photo
        base, h = img, None
    agree, seen = 0, []
    for target_px in CONFIRM_HEIGHTS:
        scale = min(8.0, max(1.0, target_px / h)) if h else target_px / CONFIRM_HEIGHTS[0]
        big = base.resize((max(1, int(base.width * scale)), max(1, int(base.height * scale))), PILImage.LANCZOS)
        for variant in (big, ImageOps.autocontrast(big).filter(ImageFilter.SHARPEN)):
            texts = engine.read(variant).texts(include_dropped=True)
            k2, how2 = find_serial_in_ocr_lines(texts, doc_serials)
            toks = _label_tokens(texts)
            if toks & rivals or (k2 and k2 in rivals):
                return False, f"a re-read shows {doc_serials[(toks & rivals or {k2}).pop()]}"
            if k2 and k2 != key:
                return False, f"a re-read shows {doc_serials[k2]}"
            if k2 == key and how2 == "exact":
                agree += 1
            seen.append(how2 or "-")
    ok = agree >= CONFIRM_MIN_AGREE
    return ok, f"{agree}/{len(seen)} zoomed re-reads exact"


def _verify_ocr_reads(paths: list[str], engine, doc_serials: dict[str, str], sticker_lookup, log,
                      cancel_event=None) -> dict:
    """
    read_serial_from_photo() for every photo -> {path: (key or None, how)}.
    With the real OpenThai engine the photos are read in parallel, one engine per worker thread
    (the shared engine runs one inference at a time); other engines (tests, EasyOCR) read serially.
    Stops early when cancel_event is set (missing paths = not read).
    """
    from concurrent.futures import ThreadPoolExecutor, as_completed
    try:
        from ocr_engine import OpenThaiEngine
        parallel = isinstance(engine, OpenThaiEngine)
    except Exception:
        parallel = False
    workers = max(1, min(4, (os.cpu_count() or 2) // 2)) if parallel else 1
    total = len(paths)
    results: dict = {}

    def cancelled():
        return cancel_event is not None and cancel_event.is_set()

    def one(path):
        eng = _worker_ocr_engine() if parallel else engine
        return read_serial_from_photo(path, eng, doc_serials, sticker_lookup=sticker_lookup)

    def record(path, outcome):
        results[path] = outcome
        if len(results) % 50 == 0 or len(results) == total:
            log(f"    Verify OCR {len(results)}/{total} photo(s)...")

    if workers == 1:
        for path in paths:
            if cancelled():
                break
            try:
                record(path, one(path))
            except Exception as e:
                log(f"      [OCR ] {os.path.basename(path)}: cannot read ({e})")
                record(path, (None, ""))
        return results

    log(f"    Verify OCR with {workers} worker(s)")
    pool = ThreadPoolExecutor(max_workers=workers, thread_name_prefix="verify")
    try:
        futures = {pool.submit(one, p): p for p in paths}
        for fut in as_completed(futures):
            if cancelled():
                break
            path = futures[fut]
            try:
                record(path, fut.result())
            except Exception as e:
                log(f"      [OCR ] {os.path.basename(path)}: cannot read ({e})")
                record(path, (None, ""))
    finally:
        pool.shutdown(wait=True, cancel_futures=True)
    return results


def verify_folder_photos(folders: list[dict], doc_serials: dict[str, str], engine=None,
                         sticker_lookup: dict[int, list[dict]] | None = None, time_fn=None,
                         log_cb=None, cancel_event=None, max_fuzzy: float = VERIFY_MAX_FUZZY) -> list[dict]:
    """
    OCR every photo of folders whose device is known (S/N folders, set-number folders) and propose
    moving photos that show another device's S/N label.
      folders: [{'name': shown name, 'expected': normalized serial of the folder, 'imgs': [paths]}]
    Rules (a wrong move is worse than no move):
      - only an exact / near-exact S/N label read (see _strong_label_read) can move a photo;
        a set-number sticker may confirm the folder's own device but never moves a photo;
      - every strong read names one other device B and none the folder's own -> the whole folder is B
        (unread photos included);
      - otherwise (mixed) photos with a strong read of B move, and unread photos follow the nearest label
        through the time chain (<= UNREAD_MAX_GAP_S); without capture times only the label photos move.
    Returns proposals [{'folder': index in folders, 'path', 'to': normalized serial, 'reason', 'whole': bool}].
    """
    def log(msg):
        if log_cb:
            log_cb(msg)

    if engine is None:
        from ocr_engine import get_engine
        engine = get_engine(log_cb=log_cb)
    time_fn = time_fn or photo_capture_time
    cache = _verify_ocr_reads([p for f in folders for p in f["imgs"]], engine, doc_serials,
                              sticker_lookup, log, cancel_event)
    proposals: list[dict] = []
    for f_idx, f in enumerate(folders):
        if cancel_event is not None and cancel_event.is_set():
            break
        own = f["expected"]
        if any(p not in cache for p in f["imgs"]):
            break           # cancelled before all photos were read
        reads = [(p, *cache[p]) for p in f["imgs"]]   # (path, key or None, how)

        strong = {p: k for p, k, h in reads if k and _strong_label_read(h, max_fuzzy)}
        hows = {p: h for p, _k, h in reads}
        pending: dict[str, list] = {}          # look-alike ambiguous reads: serial -> [paths]
        for p, k in list(strong.items()):
            why = _move_read_doubt(k, own, doc_serials, hows[p]) if k != own else ""
            if why:
                del strong[p]
                if why.startswith("look-alike"):
                    pending.setdefault(k, []).append(p)
                log(f"      [VERF] {f['name']}/{os.path.basename(p)}: S/N read {doc_serials[k]} not trusted "
                    f"for a move ({why})")
        doubted = {p for paths in pending.values() for p in paths} | {
            p for p, k, h in reads if k and k != own and _strong_label_read(h, max_fuzzy) and p not in strong}
        # A weaker (fuzzy) label read of the same serial may be the second photo, if it passes the re-read too
        for p, k, h in reads:
            if k in pending and p not in pending[k] and h and not h.startswith("sticker"):
                pending[k].append(p)
        # A look-alike ambiguous serial is trusted again only with >= 2 photos that each pass the zoomed
        # re-read, and no photo of the folder reading a look-alike rival or the folder's own serial.
        # Exactly one such photo is kept for the swap check (see _swap_pair_proposals).
        single_confirmed = None
        for k, paths in pending.items():
            rivals = _lookalike_rivals(k, doc_serials) | ({own} if own else set())
            if any(rk in rivals for _p, rk, _h in reads if rk):
                log(f"      [VERF] {f['name']}: {doc_serials[k]} not confirmed - another photo reads a look-alike "
                    f"or the folder's own S/N")
                continue
            results = []
            for p in paths:
                try:
                    results.append(_confirm_label_read(p, engine, k, rivals, doc_serials))
                except Exception as e:
                    results.append((False, f"re-read failed ({e})"))
            good = [p for p, (ok, _d) in zip(paths, results) if ok]
            details = "; ".join(d for _ok, d in results)
            if len(good) >= 2:
                for p in good:
                    strong[p] = k
                log(f"      [VERF] {f['name']}: {doc_serials[k]} confirmed by {len(good)} photo(s) ({details})")
            else:
                if len(good) == 1 and len(pending) == 1:
                    single_confirmed = k
                log(f"      [VERF] {f['name']}: {doc_serials[k]} not confirmed ({details})")
        f["verify_info"] = {
            "own_read": any(k == own for _p, k, _h in reads),
            "label_keys": {k for _p, k, h in reads if k and h and not h.startswith("sticker")},
            "trusted_label_keys": {k for _p, k, h in reads if k and _strong_label_read(h, max_fuzzy)},
            "single_confirmed": single_confirmed,
        }
        foreign = {p: k for p, k in strong.items() if k != own}
        for p, k, h in reads:
            if k and k != own and h.startswith("sticker"):
                log(f"      [VERF] {f['name']}/{os.path.basename(p)}: sticker {h.split()[1]} says "
                    f"{doc_serials[k]} - sticker alone never moves a photo")
            elif k and k != own and p not in strong and p not in doubted:
                log(f"      [VERF] {f['name']}/{os.path.basename(p)}: weak S/N read {doc_serials[k]} ({h}) "
                    f"- not trusted for a move")
        if not foreign:
            continue

        others = set(foreign.values())
        if len(others) == 1 and not any(k == own for _p, k, _h in reads):
            target = next(iter(others))
            n_lab = len(foreign)
            # With a single label photo the unread photos may still be the folder's own device (e.g. a stray
            # label photo next to an unreadable wide shot): they are marked 'weak' and only move when the
            # target's folder in turn shows this folder's device (a swap, see resolve_photo_moves).
            weak = n_lab < 2
            for p, _k, _h in reads:
                why = (f"S/N label read {doc_serials[target]}" if p in foreign
                       else f"whole folder shows {doc_serials[target]} ({n_lab} label photo(s))")
                proposals.append({"folder": f_idx, "path": p, "to": target, "reason": why, "whole": True,
                                  "weak": weak and p not in foreign})
            continue

        # Mixed folder: labels decide; unread photos follow a label shot close by (time chain)
        stamps = [time_fn(p) for p, _k, _h in reads]
        if all(t is not None for t in stamps):
            order = sorted(range(len(reads)), key=lambda i: (stamps[i], os.path.basename(reads[i][0])))
            times = [stamps[i] for i in order]
        else:
            order, times = list(range(len(reads))), None
        seq = []
        for i in order:
            p, k, h = reads[i]
            if p in strong:
                seq.append((p, strong[p]))
            elif k == own:          # a weaker read (sticker / fuzzy) may confirm the folder's own device
                seq.append((p, own))
            else:
                seq.append((p, None))
        if times is not None:
            groups, _ = assign_unread_photos(seq, times)
        else:
            groups = {}
            for p, k in seq:
                groups.setdefault(k, []).append(p)
        names = ", ".join(doc_serials[k] for k in sorted(others))
        log(f"    [WARN] folder {f['name']}: mixed photos - labels of {names} besides "
            f"{doc_serials.get(own, own)}; split by S/N label"
            + ("" if times else " (no capture times: only label photos move)"))
        for target, paths in groups.items():
            if target is None or target == own:
                continue
            for p in paths:
                why = (f"S/N label read {doc_serials[target]}" if p in foreign
                       else f"no label; shot <= {UNREAD_MAX_GAP_S:.0f} s from a {doc_serials[target]} label photo")
                proposals.append({"folder": f_idx, "path": p, "to": target, "reason": why, "whole": False})
    return proposals


def _swap_pair_proposals(folders: list[dict], proposals: list[dict], holders: dict, doc_serials: dict,
                         log) -> list[dict]:
    """
    Two folders whose names were swapped: folder A (serial a) moves whole to b on trusted reads, and folder B
    (serial b) was not moved because its label read of a was not trusted alone. B also moves whole to a when:
      - B has no read of b at all (label or sticker),
      - every S/N label read in B names a, at least one of them exact or near-exact,
      - a and b are not close (distance > VERIFY_OWN_MISREAD), so B's read cannot be a misread of its own label.
    Returns the extra proposals for B (logged as [SWAP]).
    """
    extra = []
    moving = {pr["folder"] for pr in proposals}
    whole_to: dict[int, str] = {}
    for pr in proposals:
        if pr["whole"]:
            whole_to[pr["folder"]] = pr["to"]
    for a_idx, b in whole_to.items():
        a = folders[a_idx].get("expected")
        if not a:
            continue
        for b_idx in holders.get(b, []):
            info = folders[b_idx].get("verify_info")
            if b_idx in moving or not info or info["own_read"]:
                continue
            if info["label_keys"] != {a} or a not in info["trusted_label_keys"]:
                continue
            if _levenshtein(a, b, max_dist=VERIFY_OWN_MISREAD) <= VERIFY_OWN_MISREAD:
                continue
            fb = folders[b_idx]
            log(f"    [SWAP] folders {folders[a_idx]['name']} and {fb['name']} hold each other's photos - "
                f"both moved")
            for p in fb["imgs"]:
                extra.append({"folder": b_idx, "path": p, "to": a, "whole": True,
                              "reason": f"swapped with folder {folders[a_idx]['name']} (labels read "
                                        f"{doc_serials.get(a, a)})"})
            moving.add(b_idx)

    # Both halves blocked as look-alikes: each folder has exactly one label photo that passed the zoomed
    # re-read, naming the other folder's serial (two independent confirmed photos, as in the 2-photo rule).
    by_serial = {f.get("expected"): i for i, f in enumerate(folders) if f.get("verify_info")}
    for a_idx, fa in enumerate(folders):
        ia = fa.get("verify_info")
        if a_idx in moving or not ia or ia["own_read"] or not ia.get("single_confirmed"):
            continue
        a, b = fa["expected"], ia["single_confirmed"]
        b_idx = by_serial.get(b)
        if b_idx is None or b_idx in moving or b_idx == a_idx:
            continue
        fb = folders[b_idx]
        ib = fb["verify_info"]
        if ib["own_read"] or ib.get("single_confirmed") != a:
            continue
        if not ia["label_keys"] <= {b} or not ib["label_keys"] <= {a}:
            continue
        if _levenshtein(a, b, max_dist=VERIFY_OWN_MISREAD) <= VERIFY_OWN_MISREAD:
            continue
        log(f"    [SWAP] folders {fa['name']} and {fb['name']} hold each other's photos "
            f"(each confirmed by a zoomed re-read) - both moved")
        for src_idx, f_src, to, other in ((a_idx, fa, b, fb), (b_idx, fb, a, fa)):
            for p in f_src["imgs"]:
                extra.append({"folder": src_idx, "path": p, "to": to, "whole": True,
                              "reason": f"swapped with folder {other['name']} (label confirmed by zoomed re-reads)"})
        moving.update({a_idx, b_idx})
    return extra


def resolve_photo_moves(folders: list[dict], proposals: list[dict], fixed_holders: set | None = None,
                        doc_serials: dict[str, str] | None = None, log_cb=None) -> list[dict]:
    """
    Accept the proposals that do not put photos into a table that already has its own photos.
    A target serial T is free only when every folder of T (its S/N folder, its set folder) ends up empty
    after its own accepted moves, T is not in fixed_holders (photos not verified, e.g. OCR-grouped), and
    all photos for T come from a single folder. Rejections are repeated until nothing changes (a rejected
    move keeps its photos in place, which may block another move). Two swapped folders both move.
    Returns the accepted proposals; rejected ones are logged as [WARN].
    """
    def log(msg):
        if log_cb:
            log_cb(msg)

    doc_serials = doc_serials or {}
    fixed_holders = fixed_holders or set()
    holders: dict[str, list[int]] = {}
    for i, f in enumerate(folders):
        if f.get("expected"):
            holders.setdefault(f["expected"], []).append(i)
    accepted = list(proposals) + _swap_pair_proposals(folders, proposals, holders, doc_serials, log)

    # Unread photos of a folder with only one foreign label photo ('weak') move only in a swap:
    # the target's own folder moves whole to this folder's device.
    whole_to: dict[int, str] = {}
    for pr in accepted:
        if pr.get("whole"):
            whole_to.setdefault(pr["folder"], pr["to"])
    mutual = {fi for fi, to in whole_to.items()
              if any(whole_to.get(h) == folders[fi].get("expected") for h in holders.get(to, []))}
    kept, dropped = [], {}
    for pr in accepted:
        if pr.get("weak") and pr["folder"] not in mutual:
            dropped.setdefault(pr["folder"], []).append(pr)
            continue
        kept.append(pr)
    for fi, prs in dropped.items():
        log(f"    [VERF] folder {folders[fi]['name']}: only 1 label photo shows {doc_serials.get(prs[0]['to'], prs[0]['to'])}"
            f" - that photo is moved, {len(prs)} unread photo(s) stay in the folder")
    accepted = kept
    proposals = accepted[:]
    rejected_why: dict[str, str] = {}
    while True:
        out_count: dict[int, int] = {}
        for pr in accepted:
            out_count[pr["folder"]] = out_count.get(pr["folder"], 0) + 1
        sources: dict[str, set] = {}
        for pr in accepted:
            sources.setdefault(pr["to"], set()).add(pr["folder"])
        bad = {}
        for target, srcs in sources.items():
            if target in fixed_holders:
                bad[target] = "its table already has photos (OCR-grouped)"
            elif len(srcs) > 1:
                bad[target] = "photos of it found in several folders (" + ", ".join(
                    folders[i]["name"] for i in sorted(srcs)) + ")"
            else:
                for h in holders.get(target, []):
                    if len(folders[h]["imgs"]) > out_count.get(h, 0):
                        bad[target] = f"its table already has its own photos (folder {folders[h]['name']})"
                        break
        if not bad:
            break
        rejected_why.update(bad)
        accepted = [pr for pr in accepted if pr["to"] not in bad]
    for target, why in rejected_why.items():
        n = sum(1 for pr in proposals if pr["to"] == target)
        srcs = sorted({folders[pr["folder"]]["name"] for pr in proposals if pr["to"] == target})
        log(f"    [WARN] {n} photo(s) in {', '.join(srcs)} show S/N {doc_serials.get(target, target)} "
            f"but were NOT moved: {why} - please check")
    return accepted


# ─────────────────────────────────────────────
# Auto Mode fallback: set-number folders ('ชุดที่ 12/280' <-> folder '12-280')
# ─────────────────────────────────────────────
SET_LABEL_REGEX = re.compile(r"^\s*(?:ชุด\s*ที่|set\s*(?:no\.?)?|unit\s*(?:no\.?)?)\s*[:.]?\s*$", re.IGNORECASE)
SET_VALUE_REGEX = re.compile(r"^\s*(\d{1,4})\s*/\s*(\d{1,4})\s*$")
PN_LABEL_REGEX = re.compile(r"^\s*(?:part\s*(?:number|no\.?)|p\s*/\s*n)\s*[:.]?\s*$", re.IGNORECASE)
ITEM_LABEL_REGEX = re.compile(r"^\s*(?:รายการ|item|description|model)\s*[:.]?\s*$", re.IGNORECASE)
GENERIC_FOLDER_REGEX = re.compile(r"^(?:by[\s_-]*(?:number|num|no|sn|serial)|numbers?|no|sn|serials?)$", re.IGNORECASE)


def extract_set_entries(doc) -> list[dict]:
    """
    One entry per detail table that has a 'ชุดที่ : N/T' row:
      {t_idx, n, total, pn, item, sn, last_row, group, index, group_total}
    Tables of the same Part Number (else same รายการ) form one device group. When a group is
    split over several sections (e.g. 1/229 ... 229/229 then 1/51 ... 51/51 of the same PN),
    'index' numbers it continuously in document order (1..280) and group_total is the sum.
    """
    entries = []
    for t_idx, table in enumerate(doc.tables, start=1):
        rows: dict[int, list] = {}
        for r, c, cell in _iter_unique_cells(table):
            rows.setdefault(r, []).append((c, cell.text.strip()))
        info: dict = {}
        for r in sorted(rows):
            cells = rows[r]
            for i in range(len(cells) - 1):
                label, value = cells[i][1], cells[i + 1][1]
                if not value:
                    continue
                m = SET_VALUE_REGEX.match(value)
                if SET_LABEL_REGEX.match(label) and m and "n" not in info:
                    info.update(n=int(m.group(1)), total=int(m.group(2)))
                elif PN_LABEL_REGEX.match(label):
                    info.setdefault("pn", value)
                elif ITEM_LABEL_REGEX.match(label):
                    info.setdefault("item", value)
                elif SERIAL_LABEL_REGEX.match(label):
                    info.setdefault("sn", value)
                else:
                    continue
                info["last_row"] = max(info.get("last_row", 0), r)
        if "n" in info and 1 <= info["n"] <= info["total"]:
            info["t_idx"] = t_idx
            entries.append(info)

    # Continuous numbering per device group
    segments: dict[str, list] = {}
    for e in entries:
        e["group"] = normalize_sn(e.get("pn")) or e.get("item", "").strip() or "?"
        seg = (e.get("item", ""), e["total"])
        segs = segments.setdefault(e["group"], [])
        if seg not in segs:
            segs.append(seg)
    for e in entries:
        segs = segments[e["group"]]
        seg_i = segs.index((e.get("item", ""), e["total"]))
        e["index"] = sum(t for _, t in segs[:seg_i]) + e["n"]
        e["group_total"] = sum(t for _, t in segs)
    return entries


def find_set_number_folders(root: str) -> list[dict]:
    """
    Folders named like '12-280' (set 12 of 280) that hold images, at any depth.
    model = nearest parent folder that is not a generic container ('By_Number', 'SN', ...).
    Returns [{model, n, total, name, imgs}].
    """
    out = []
    if not root or not os.path.isdir(root):
        return out
    root_abs = os.path.abspath(root)
    root_name = os.path.basename(root_abs)
    for dirpath, dirnames, filenames in os.walk(root_abs):
        dirnames[:] = sorted(d for d in dirnames if not d.startswith(('.', '__')))
        imgs = [os.path.join(dirpath, f) for f in sorted(filenames)
                if '_converted' not in f and os.path.splitext(f)[1].lower() in VALID_IMG_EXTS]
        if not imgs:
            continue
        rel = os.path.relpath(dirpath, root_abs)
        parts = [] if rel == "." else rel.split(os.sep)
        while parts and STRICT_CAMERA_DIR_REGEX.match(parts[-1]):
            parts.pop()
        if not parts:
            continue
        m = re.match(r"^\s*(\d{1,4})\s*[-_/]\s*(\d{1,4})\s*$", parts[-1])
        if not m:
            continue
        n, total = int(m.group(1)), int(m.group(2))
        if not (1 <= n <= total):
            continue
        ancestors = [p for p in reversed(parts[:-1]) if not GENERIC_FOLDER_REGEX.match(p)]
        model = ancestors[0] if ancestors else root_name
        out.append({"model": model, "n": n, "total": total, "name": parts[-1], "imgs": imgs})
    return out


def _name_similarity(a: str, b: str) -> float:
    from difflib import SequenceMatcher
    na, nb = normalize_sn(a), normalize_sn(b)
    if not na or not nb:
        return 0.0
    if model_matches(a, b):
        return 1.0
    return SequenceMatcher(None, na, nb).ratio()


def map_set_folders_to_tables(set_folders: list[dict], entries: list[dict],
                              pn_to_desc: dict[str, str] | None = None, log_cb=None) -> dict[int, dict]:
    """
    Decide which device group each model folder belongs to, then map every set folder to a table.
    A group fits a folder when its total matches (whole group, or one section of it) and its
    name (summary-table Description, else รายการ) resembles the folder's model name.
    If names give no answer, a total that fits exactly one group is accepted; anything else is skipped.
    Returns {t_idx: {folder entry + 'entry': doc entry, 'how': str}}.
    """
    def log(msg):
        if log_cb:
            log_cb(msg)

    pn_to_desc = pn_to_desc or {}
    groups: dict[str, dict] = {}
    for e in entries:
        g = groups.setdefault(e["group"], {"total": e["group_total"], "sections": set(),
                                           "names": set(), "by_index": {}, "by_section": {}})
        g["sections"].add(e["total"])
        if e.get("pn") and e["pn"] in pn_to_desc:
            g["names"].add(pn_to_desc[e["pn"]])
        if e.get("item"):
            g["names"].add(e["item"])
        g["by_index"][e["index"]] = e
        g["by_section"].setdefault(e["total"], {}).setdefault(e["n"], []).append(e)

    result: dict[int, dict] = {}
    by_model: dict[tuple, list] = {}
    for f in set_folders:
        by_model.setdefault((f["model"], f["total"]), []).append(f)

    for (model, total), folders in sorted(by_model.items()):
        fits = []
        for key, g in groups.items():
            if g["total"] == total:
                mode = "whole"
            elif total in g["sections"] and len(g["by_section"][total]) == total \
                    and all(len(v) == 1 for v in g["by_section"][total].values()):
                mode = "section"
            else:
                continue
            sim = max((_name_similarity(model, nm) for nm in g["names"]), default=0.0)
            fits.append((sim, key, mode))
        fits.sort(reverse=True)
        chosen, how = None, ""
        if fits and fits[0][0] >= 0.8 and (len(fits) == 1 or fits[1][0] < fits[0][0]):
            chosen, how = fits[0], f"name ~ {max(groups[fits[0][1]]['names'], key=lambda nm: _name_similarity(model, nm))}"
        elif len(fits) == 1:
            chosen, how = fits[0], "only group with this total"
        if not chosen:
            why = "no device group with this total" if not fits else f"{len(fits)} possible device groups"
            log(f"  [SKIP] set folders '{model}' (x/{total}, {len(folders)} folder(s)): {why}")
            continue
        _, key, mode = chosen
        g = groups[key]
        log(f"  [SET ] '{model}' x/{total} -> {sorted(g['names'], key=len)[0][:40]} "
            f"({'numbered across sections' if mode == 'whole' and len(g['sections']) > 1 else mode}; {how})")
        for f in folders:
            if mode == "whole":
                e = g["by_index"].get(f["n"])
            else:
                lst = g["by_section"][total].get(f["n"], [])
                e = lst[0] if len(lst) == 1 else None
            if e is not None:
                result[e["t_idx"]] = dict(f, entry=e, how=how)
    return result


# ─────────────────────────────────────────────
# Run summary report (<output>_summary.txt)
# ─────────────────────────────────────────────
def find_empty_folders(root: str) -> list[str]:
    """Names of folders under root that contain no photos at any depth (e.g. an S/N folder with 0 files)."""
    has_images: dict[str, bool] = {}
    if not root or not os.path.isdir(root):
        return []
    for dirpath, dirnames, filenames in os.walk(root, topdown=False):
        own = any(os.path.splitext(f)[1].lower() in VALID_IMG_EXTS for f in filenames)
        has_images[dirpath] = own or any(has_images.get(os.path.join(dirpath, d), False) for d in dirnames)
    root_abs = os.path.normpath(root)
    return sorted({os.path.basename(p) for p, h in has_images.items() if not h and os.path.normpath(p) != root_abs})


def summary_path_for(output_path: str) -> str:
    return os.path.splitext(output_path)[0] + "_summary.txt"


def build_injection_summary(doc, *, docx_path: str, photos_root: str, output_path: str, saved: bool,
                            injected_tables: dict, skip_reasons: dict, start_table: int = 1,
                            active_models: list | None = None, unused_folders: list | None = None,
                            unassigned_photos: list | None = None, warnings: list | None = None,
                            cancelled: bool = False, images_total: int = 0, errors_total: int = 0,
                            empty_folders: list | None = None, moved_photos: list | None = None,
                            moved_out: dict | None = None) -> str:
    """
    Plain-text report of a run, per device group of the document:
      units with photos, partial units (fewer photos than usual), units without photos + reason,
    then photo folders not in the document, photos whose S/N could not be read, warnings/errors.
    injected_tables: {t_idx: (label, n_photos)}; skip_reasons: {t_idx: reason}.
    """
    from collections import Counter
    from datetime import datetime

    pn_to_desc = extract_doc_part_descriptions(doc)
    entries = {e["t_idx"]: e for e in extract_set_entries(doc)}
    doc_serials = extract_doc_serials(doc)
    serial_index = build_sn_index(list(doc_serials.values()))

    # Device tables: tables with 'ชุดที่ N/T', plus tables holding exactly one document serial
    # that no 'ชุดที่' table already has (a one-row summary table must not count twice).
    entry_serials = {normalize_sn(e.get("sn")) for e in entries.values() if e.get("sn")}
    seen_other: set = set()
    devices = []
    for t_idx, table in enumerate(doc.tables, start=1):
        e = entries.get(t_idx)
        if e:
            sn = e.get("sn") or ""
            sn_rc = (e.get("last_row", 0), 1)
            name = pn_to_desc.get(e.get("pn", "")) or e.get("item", "") or "Devices"
            group = e["group"]
            pn = e.get("pn", "")
        else:
            found = auto_locate_in_table(table, serial_index)
            if not found or found["status"] not in ("ok", "no_image_cell"):
                continue
            sn, sn_rc = found["sn"], found["sn_rc"]
            key = normalize_sn(sn)
            if key in entry_serials or key in seen_other:
                continue
            seen_other.add(key)
            models = extract_table_models(table, sn_row_col=sn_rc, pn_to_desc=pn_to_desc)
            name = models[0] if models else "Devices"
            group, pn = name, ""
        devices.append({"t_idx": t_idx, "sn": sn, "entry": e, "name": name, "group": group, "pn": pn,
                        "table": table, "sn_rc": sn_rc})
    empty_keys = {normalize_sn(f) for f in (empty_folders or [])}

    def set_label(d):
        e = d["entry"]
        return f"{e['n']}/{e['total']}" if e else f"table {d['t_idx']}"

    def excluded_by_filter(d):
        if active_models is None:
            return False
        if not active_models:
            return True
        models = extract_table_models(d["table"], sn_row_col=d["sn_rc"], pn_to_desc=pn_to_desc) + [d["name"]]
        return not any(model_matches(sel, m) for sel in active_models for m in models)

    groups: dict = {}
    for d in devices:
        groups.setdefault(d["group"], []).append(d)

    W = 70
    out = ["=" * W, "SUMMARY OF PHOTO INJECTION", "=" * W,
           f"Document : {docx_path}",
           f"Photos   : {photos_root}",
           f"Output   : {output_path}" + ("" if saved else "   (NOT saved: no photos inserted)"),
           f"Date     : {datetime.now().strftime('%Y-%m-%d %H:%M')}"]
    with_photos_total = sum(1 for d in devices if d["t_idx"] in injected_tables)
    out.append(f"Result   : {with_photos_total} of {len(devices)} device tables have photos, "
               f"{images_total} images, {errors_total} error(s)"
               + ("   [CANCELLED - partial run]" if cancelled else ""))
    out.append("")

    sec = 0
    for key, devs in groups.items():
        sec += 1
        first = devs[0]
        title = first["name"] + (f" (Part Number {first['pn']})" if first["pn"] else "")
        out += ["", f"{sec}. {title} (Total: {len(devs)} Units)", "-" * W]

        counts = [injected_tables[d["t_idx"]][1] for d in devs if d["t_idx"] in injected_tables]
        typical = Counter(counts).most_common(1)[0][0] if counts else 0
        before = [d for d in devs if d["t_idx"] < start_table]
        filtered = [d for d in devs if d["t_idx"] >= start_table and d["t_idx"] not in injected_tables
                    and excluded_by_filter(d)]
        partial = [d for d in devs if d["t_idx"] in injected_tables and injected_tables[d["t_idx"]][1] < typical]
        missing = [d for d in devs if d["t_idx"] >= start_table and d["t_idx"] not in injected_tables
                   and d not in filtered]

        if counts and not partial and not missing and not before and not filtered:
            out.append(f"- All {len(devs)} units have photos ({typical} each)")
        else:
            out.append(f"- With photos: {len(counts)} unit(s)" + (f" (usually {typical} photos each)" if counts else ""))
        if partial:
            out += [f"- Partial (fewer photos than usual): {len(partial)} unit(s)",
                    f"  {'ชุดที่':<12}| {'Serial Number':<22}| Photos",
                    f"  {'-' * 12}|{'-' * 23}|{'-' * 7}"]
            out += [f"  {set_label(d):<12}| {d['sn']:<22}| {injected_tables[d['t_idx']][1]}" for d in partial]
        if missing:
            out += [f"- Missing photos: {len(missing)} unit(s)",
                    f"  {'ชุดที่':<12}| {'Serial Number':<22}| Reason",
                    f"  {'-' * 12}|{'-' * 23}|{'-' * 30}"]
            for d in missing:
                reason = skip_reasons.get(d["t_idx"])
                if not reason and moved_out and d["sn"]:
                    reason = moved_out.get(normalize_sn(d["sn"]))
                if not reason:
                    reason = ("photo folder is empty (0 files)" if d["sn"] and normalize_sn(d["sn"]) in empty_keys
                              else "no photos found (no S/N folder, photo or sticker)")
                out.append(f"  {set_label(d):<12}| {d['sn'] or '-':<22}| {reason}")
        if filtered:
            out.append(f"- Not selected in Device Models filter: {len(filtered)} unit(s)")
        if before:
            out.append(f"- Before Start Table ({start_table}): {len(before)} unit(s), not processed")

    serial_tables: dict = {}
    for d in devices:
        if d["sn"]:
            serial_tables.setdefault(normalize_sn(d["sn"]), []).append(d)
    dups = {k: v for k, v in serial_tables.items() if len(v) > 1}
    if dups:
        sec += 1
        out += ["", f"{sec}. Document check: same S/N in more than one table ({len(dups)})", "-" * W]
        for k, v in dups.items():
            out.append(f"  {v[0]['sn']:<22} in " + ", ".join(f"{d['name']} {set_label(d)} (table {d['t_idx']})"
                                                        for d in v))
    if unused_folders:
        sec += 1
        out += ["", f"{sec}. Photo folders not found in the document ({len(unused_folders)})", "-" * W]
        out += [f"  {f}" for f in unused_folders]
    if unassigned_photos:
        sec += 1
        out += ["", f"{sec}. Photos without a readable S/N - not inserted ({len(unassigned_photos)})", "-" * W]
        out += [f"  {p}" for p in unassigned_photos]
    if moved_photos:
        sec += 1
        by_serial = {}
        for d in devices:
            if d["sn"]:
                by_serial.setdefault(normalize_sn(d["sn"]), d)
        out += ["", f"{sec}. Photos moved (wrong folder) ({len(moved_photos)})", "-" * W,
                "  Word output only - the files on disk were not moved.",
                f"  {'Photo':<34}| {'From folder':<24}| {'To table (ชุดที่ / S/N)':<34}| Reason",
                f"  {'-' * 34}|{'-' * 25}|{'-' * 35}|{'-' * 30}"]
        for mv in moved_photos:
            d = by_serial.get(normalize_sn(mv["to"]))
            if d:
                e = d["entry"]
                to = (f"{e['n']}/{e['total']}, " if e else "") + f"table {d['t_idx']}: {d['sn']}"
                if d["t_idx"] not in injected_tables:
                    to += " [not inserted]"
            else:
                to = mv["to"]
            out.append(f"  {mv['photo']:<34}| {mv['from']:<24}| {to:<34}| {mv['reason']}")
    if warnings:
        sec += 1
        out += ["", f"{sec}. Warnings and errors ({len(warnings)})", "-" * W]
        out += [f"  {w.strip()}" for w in warnings]
    out.append("")
    return "\n".join(out)


def _verify_and_move_photos(doc, sn_root: str, sn_folders: dict, sn_to_model: dict, set_folders: list,
                            ocr_grouped: set, report_moved: list, moved_out: dict, log, cancel_event=None,
                            engine=None, time_fn=None) -> int:
    """
    Verify step of Auto mode: OCR the photos of S/N folders and set-number folders, then move the photos
    of another device into that device's entry (sn_folders / set_folders are changed in place; files on
    disk are untouched). Fills report_moved [{photo, from, to, reason}] and moved_out {serial: reason}.
    Returns the number of photos moved.
    """
    doc_serials = extract_doc_serials(doc)
    if not doc_serials:
        log("\nVerify photos: no serial numbers in the document - skipped")
        return 0
    entries = extract_set_entries(doc)
    sticker_lookup = build_set_sticker_lookup(entries)
    set_expected: dict[int, str] = {}
    if set_folders:
        set_map = map_set_folders_to_tables(set_folders, entries, extract_doc_part_descriptions(doc))
        for m in set_map.values():
            key = normalize_sn(m["entry"].get("sn"))
            if key in doc_serials:
                set_expected[id(m["imgs"])] = key

    def folder_name(imgs, fallback):
        try:
            return os.path.relpath(os.path.dirname(imgs[0]), sn_root)
        except (ValueError, IndexError):
            return fallback

    folders = []
    sn_keys = {normalize_sn(n) for n in sn_folders}
    for name in sorted(sn_folders):
        key = normalize_sn(name)
        if key in doc_serials and key not in ocr_grouped and sn_folders[name]:
            folders.append({"name": folder_name(sn_folders[name], name), "expected": key,
                            "imgs": list(sn_folders[name]), "kind": "sn", "ref": name, "verify": True})
    for f in set_folders:
        key = set_expected.get(id(f["imgs"]))
        if key:
            folders.append({"name": folder_name(f["imgs"], f["name"]), "expected": key, "imgs": list(f["imgs"]),
                            "kind": "set", "ref": f, "verify": key not in sn_keys})
    check_idx = [i for i, f in enumerate(folders) if f["verify"]]
    n_photos = sum(len(folders[i]["imgs"]) for i in check_idx)
    log(f"\nVerify photos: OCR of {n_photos} photo(s) in {len(check_idx)} S/N / set folder(s) "
        f"to catch photos in the wrong folder")
    if not check_idx:
        return 0
    try:
        props = verify_folder_photos([folders[i] for i in check_idx], doc_serials, engine=engine,
                                     sticker_lookup=sticker_lookup, time_fn=time_fn, log_cb=log,
                                     cancel_event=cancel_event)
    except Exception as e:
        log(f"  [ERR ] Verify photos: OCR unavailable: {e}")
        return 0
    if cancel_event is not None and cancel_event.is_set():
        return 0
    for pr in props:
        pr["folder"] = check_idx[pr["folder"]]
    accepted = resolve_photo_moves(folders, props, fixed_holders=set(ocr_grouped),
                                   doc_serials=doc_serials, log_cb=log)
    if not accepted:
        log("  Verify photos: no photo needs moving")
        return 0

    # Apply: first take photos out of their folders, then give them to their device
    by_src: dict[int, set] = {}
    by_target: dict[str, list] = {}
    for pr in accepted:
        by_src.setdefault(pr["folder"], set()).add(pr["path"])
        by_target.setdefault(pr["to"], []).append(pr)
    for i, moved in by_src.items():
        f = folders[i]
        keep = [p for p in f["imgs"] if p not in moved]
        if f["kind"] == "sn":
            if keep:
                sn_folders[f["ref"]] = keep
            else:
                sn_folders.pop(f["ref"], None)
        else:
            f["ref"]["imgs"][:] = keep
        if not keep:
            moved_out[f["expected"]] = f"all photos in folder {f['name']} showed another S/N (moved)"
    set_folders[:] = [f for f in set_folders if f["imgs"]]
    for target, prs in by_target.items():
        src = folders[prs[0]["folder"]]
        for name in [n for n in sn_folders if normalize_sn(n) == target]:
            sn_folders.pop(name)        # emptied by its own moves (resolve_photo_moves checked)
        new_name = doc_serials[target]
        sn_folders[new_name] = [pr["path"] for pr in prs]
        model = sn_to_model.get(src["ref"]) if src["kind"] == "sn" else src["ref"].get("model")
        if model and not sn_to_model.get(new_name):
            sn_to_model[new_name] = model
        moved_out.pop(target, None)
        emptied = all(p in by_src.get(prs[0]["folder"], set()) for p in src["imgs"])
        log(f"  [MOVE] {len(prs)} photo(s) from folder {src['name']} -> S/N {new_name}"
            + (" (whole folder)" if emptied else ""))
        for pr in prs:
            report_moved.append({"photo": os.path.relpath(pr["path"], sn_root), "from": src["name"],
                                 "to": new_name, "reason": pr["reason"]})
            log(f"      {os.path.basename(pr['path'])}: {pr['reason']}")
    log(f"  Verify photos: {len(accepted)} photo(s) moved to the right device (Word output only)")
    return len(accepted)


def run_injection(docx_path: str, sn_root: str, output_path: str,
                  keyword: str = "WZP",
                  sn_row_col: tuple = (5, 3), img_row_col: tuple = (12, 0),
                  start_table: int = 1,
                  selected_models: list[str] | None = None,
                  auto_match: bool = False,
                  photo_ocr: bool = False,
                  set_number_fallback: bool = False,
                  verify_photos: bool = False,
                  max_image_px: int = INSERT_MAX_PX,
                  workers: int | None = None,
                  write_summary: bool = True,
                  auto_rotate: bool = True, smart_rotate: bool = False,
                  log_callback=None, progress_callback=None,
                  status_callback=None,
                  cancel_event=None) -> InjectionResult:
    """
    Main injection logic. Returns InjectionResult (supports 4-tuple unpacking and dict access).
    - start_table: 1-based table index to begin injection. Tables before this index are untouched.
    - selected_models: list of model strings to inject. Non-matching tables/folders are untouched.
    - auto_match: Auto mode — ignore keyword and fixed S/N cell; scan every cell of each table for a
      serial that matches a photo folder name, then detect the photo cell (img_row_col is the fallback).
    - photo_ocr: (auto mode) photos in folders NOT named by a document serial are OCR'd for the S/N
      on their label; unreadable photos go to the serial shot closest in time.
    - set_number_fallback: (auto mode) tables still without photos are matched by 'ชุดที่ N/T' to
      folders named 'N-T' (model folder <-> device group by name and total). S/N always wins.
    - verify_photos: (auto mode) OCR every photo of S/N and set-number folders too; a photo whose S/N
      label clearly names another device is put into that device's table (Word output only, files on
      disk are never moved). Never fills a table that already has its own photos. See verify_folder_photos.
    - max_image_px: shrink photos to this long side before inserting (0 = embed originals).
    - workers: photos prepared in parallel (rotation/OCR/resize); None = automatic.
    - write_summary: write '<output>_summary.txt' (per device: photos / partial / missing + reasons).
    - status_callback: alias for log_callback.
    - cancel_event: threading.Event — set to stop processing early.
    """
    from docx import Document
    from docx.shared import Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from PIL import Image as PILImage

    status_cb = log_callback or status_callback
    report_warnings: list[str] = []
    skip_reasons: dict[int, str] = {}
    injected_tables: dict[int, tuple] = {}
    report_unused: list[str] = []
    report_unassigned: list[str] = []
    report_moved: list[dict] = []
    moved_out: dict[str, str] = {}
    ocr_grouped: set[str] = set()

    def log(msg):
        if "[WARN]" in msg or "[ERR" in msg or "ERROR" in msg:
            report_warnings.append(msg)
        if status_cb:
            status_cb(msg)

    def is_cancelled():
        return cancel_event and cancel_event.is_set()

    # ── Step 1: Find SN folders ──
    # If selected_models is active, do not restrict folder scanning by keyword prefix
    model_filtering_active = selected_models is not None
    active_models = [str(m).strip() for m in selected_models if m and str(m).strip()] if model_filtering_active else None
    scan_kw = "" if (model_filtering_active or auto_match) else (keyword or "")
    if auto_match:
        log("AUTO MODE: matching serial numbers anywhere in each table")
    log(f"Scanning folders matching '{scan_kw}*'...")
    doc = None
    if auto_match:
        sn_folders = find_sn_leaf_folders(sn_root, log_cb=log)
    else:
        sn_folders = find_sn_folders(sn_root, scan_kw)
    sn_to_model = getattr(sn_folders, 'sn_to_model', {})
    log(f"  Found {len(sn_folders)} folder(s) with images")

    if auto_match and photo_ocr and sn_folders:
        doc = Document(docx_path)
        doc_serials = extract_doc_serials(doc)
        unnamed = [name for name in sorted(sn_folders) if normalize_sn(name) not in doc_serials]
        other_serial = [n for n in unnamed if _folder_named_like_other_serial(n, doc_serials)]
        if other_serial:
            skip = set(other_serial)
            unnamed = [n for n in unnamed if n not in skip]
            log(f"\n  Skipping OCR in {len(other_serial)} folder(s) named by a serial that is not in this "
                f"document ({', '.join(other_serial[:3])}{' ...' if len(other_serial) > 3 else ''})")
        if unnamed and doc_serials:
            ocr_imgs = [img for name in unnamed for img in sn_folders[name]]
            log(f"\nOCR: reading S/N labels in {len(ocr_imgs)} photo(s) from {len(unnamed)} folder(s) "
                f"not named by serial ({', '.join(unnamed[:3])}{' ...' if len(unnamed) > 3 else ''})")
            log(f"  Serial numbers in document: {len(doc_serials)}")
            sticker_lookup = build_set_sticker_lookup(extract_set_entries(doc))
            if sticker_lookup:
                log("  Set-number stickers understood (N/total): "
                    + ", ".join(f"/{t}" for t in sorted(sticker_lookup)))
            try:
                groups, unassigned = group_photos_by_ocr(ocr_imgs, doc_serials, log_cb=log,
                                                         cancel_event=cancel_event,
                                                         sticker_lookup=sticker_lookup)
            except Exception as e:
                groups, unassigned = {}, []
                log(f"  [ERR ] OCR unavailable: {e}")
            src_model = next((sn_to_model.get(n) for n in unnamed if sn_to_model.get(n)), None)
            for name in unnamed:
                sn_folders.pop(name, None)
                sn_to_model.pop(name, None)
            existing = {normalize_sn(k) for k in sn_folders}
            used = {}
            for serial, imgs in groups.items():
                if normalize_sn(serial) in existing:
                    log(f"  [WARN] {serial}: already has an S/N folder; OCR photos ignored")
                    continue
                used[serial] = imgs
                sn_folders[serial] = imgs
                if src_model:
                    sn_to_model[serial] = src_model
            ocr_grouped = {normalize_sn(k) for k in used}
            counts = [len(v) for v in used.values()]
            if counts:
                typical = max(set(counts), key=counts.count)
                odd = sorted(k for k, v in used.items() if len(v) != typical)
                log(f"  OCR grouped {sum(counts)} photo(s) into {len(used)} serial(s) "
                    f"(usually {typical} per serial)")
                if odd:
                    log(f"  [WARN] {len(odd)} serial(s) with an unusual photo count, please check: "
                        + ", ".join(f"{k}={len(used[k])}" for k in odd[:8]))
            if unassigned:
                report_unassigned.extend(os.path.relpath(p, sn_root) for p in unassigned)
                log(f"  [WARN] {len(unassigned)} photo(s) without a readable S/N: "
                    + ", ".join(os.path.basename(p) for p in unassigned[:5]))
            if is_cancelled():
                log("\n[CANCELLED] Stopped by user.")
                return InjectionResult(0, 0, 0, ["Cancelled"])

    set_folders = find_set_number_folders(sn_root) if (auto_match and set_number_fallback) else []
    if set_folders:
        log(f"  Found {len(set_folders)} set-number folder(s) (e.g. '{set_folders[0]['name']}') "
            f"for the ชุดที่ fallback")

    if auto_match and verify_photos and (sn_folders or set_folders):
        if doc is None:
            doc = Document(docx_path)
        _verify_and_move_photos(doc, sn_root, sn_folders, sn_to_model, set_folders, ocr_grouped,
                                report_moved, moved_out, log, cancel_event)
        if is_cancelled():
            log("\n[CANCELLED] Stopped by user.")
            return InjectionResult(0, 0, 0, ["Cancelled"])

    if not sn_folders and not set_folders:
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
    if doc is None:
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
    sn_index = build_sn_index(sn_folders.keys()) if auto_match else None
    auto_skipped = 0

    # ── Step 3: Collect eligible candidate tables in sequential document order ──
    candidates = []
    for t_idx, table in enumerate(doc.tables, start=1):
        if t_idx < start_table:
            # Skip tables prior to start_table without touching DOM
            continue

        table_sn_rc, table_img_rc = sn_row_col, img_row_col
        if auto_match:
            found = auto_locate_in_table(table, sn_index, fallback_img_rc=img_row_col)
            if found is None:
                continue
            if found["status"] == "ambiguous":
                auto_skipped += 1
                shown = ", ".join(found["sns"][:4]) + (" ..." if len(found["sns"]) > 4 else "")
                n_serials = max(len(found["sns"]), found.get("list_rows", 0))
                log(f"  [SKIP] table {t_idx}: {n_serials} serials in one table "
                    f"(summary/list table?) — {shown}")
                continue
            if found["status"] == "no_image_cell":
                auto_skipped += 1
                skip_reasons[t_idx] = f"photo folder found, but {found['reason']} in the table"
                log(f"  [SKIP] {found['sn']} (table {t_idx}): S/N found at "
                    f"R{found['sn_rc'][0]+1}C{found['sn_rc'][1]+1} but {found['reason']}")
                continue
            sn = found["sn"]
            table_sn_rc, table_img_rc = found["sn_rc"], found["img_rc"]
        else:
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
            table_models = extract_table_models(table, sn_row_col=table_sn_rc, pn_to_desc=pn_to_desc)
            folder_model = sn_to_model.get(sn)
            matched = any(
                any(model_matches(sel, tm) for tm in table_models) or
                (folder_model and model_matches(sel, folder_model))
                for sel in active_models
            )
            if not matched:
                continue
        elif keyword_upper and not auto_match:
            if not sn.upper().startswith(keyword_upper):
                continue

        if auto_match:
            log(f"  [FIND] {sn} (table {t_idx}): S/N at R{table_sn_rc[0]+1}C{table_sn_rc[1]+1} "
                f"-> photo at R{table_img_rc[0]+1}C{table_img_rc[1]+1} ({found['reason']})")
        candidates.append((t_idx, table, sn, table_img_rc, None))

    # ── Step 3a: the same S/N in several tables (document error) -> photos only in the first one ──
    if auto_match:
        first_table: dict[str, int] = {}
        kept = []
        for c in candidates:
            key = normalize_sn(c[2])
            if key in first_table:
                skip_reasons[c[0]] = (f"S/N also in table {first_table[key]} (document error?) - "
                                      f"photos not repeated")
                log(f"  [WARN] {c[2]}: same S/N in tables {first_table[key]} and {c[0]} - photos used only in "
                    f"table {first_table[key]}; please check the document")
                continue
            first_table[key] = c[0]
            kept.append(c)
        candidates = kept

    # ── Step 3b: Fallback — match remaining tables by ชุดที่ (set number) ──
    if set_folders:
        log(f"\n  ชุดที่ fallback: matching {len(set_folders)} set-number folder(s)...")
        entries = extract_set_entries(doc)
        set_map = map_set_folders_to_tables(set_folders, entries,
                                            pn_to_desc or extract_doc_part_descriptions(doc), log_cb=log)
        covered = {c[0] for c in candidates}
        set_added = set_covered = 0
        for t_idx in sorted(set_map):
            m = set_map[t_idx]
            e = m["entry"]
            if t_idx < start_table:
                continue
            if t_idx in covered:
                set_covered += 1
                continue
            table = doc.tables[t_idx - 1]
            if model_filtering_active:
                table_models = extract_table_models(table, sn_row_col=(e["last_row"], 1),
                                                    pn_to_desc=pn_to_desc)
                if not active_models or not any(
                        any(model_matches(sel, tm) for tm in table_models) or model_matches(sel, m["model"])
                        for sel in active_models):
                    continue
            img_rc, reason = detect_image_cell(table, e["last_row"], img_row_col)
            label = e.get("sn") or f"ชุดที่ {e['n']}/{e['total']}"
            if img_rc is None:
                auto_skipped += 1
                skip_reasons[t_idx] = f"set folder {m['name']} found, but {reason} in the table"
                log(f"  [SKIP] {label} (table {t_idx}): set folder {m['model']}/{m['name']} but {reason}")
                continue
            log(f"  [SET ] {label} (table {t_idx}): ชุดที่ {e['n']}/{e['total']} <- folder "
                f"{m['model']}/{m['name']} -> photo at R{img_rc[0]+1}C{img_rc[1]+1} ({reason})")
            candidates.append((t_idx, table, label, img_rc, m["imgs"]))
            set_added += 1
        candidates.sort(key=lambda c: c[0])
        log(f"  ชุดที่ fallback: {set_added} table(s) added, {set_covered} already matched by S/N, "
            f"{len(set_folders) - len(set_map)} folder(s) not mapped")

    log(f"  Eligible tables to evaluate: {len(candidates)}")
    if auto_match:
        matched_sns = {c[2] for c in candidates if c[4] is None}
        unused = [sn for sn in sorted(sn_folders) if sn not in matched_sns]
        report_unused.extend(unused)
        log(f"  Auto mode: {len(matched_sns)}/{len(sn_folders)} folder(s) matched, "
            f"{auto_skipped} table(s) skipped")
        if unused:
            log(f"  Folders not found in document ({len(unused)}): "
                + ", ".join(unused[:10]) + (" ..." if len(unused) > 10 else ""))

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

    _patch_docx_image_sha1_cache()

    # Prepare photos (HEIC/EXIF/Smart Rotate/resize) in parallel, insert them in order below.
    from concurrent.futures import ThreadPoolExecutor
    cpus = os.cpu_count() or 2
    if workers is None:
        workers = max(1, min(4, cpus // 2)) if smart_rotate else max(1, min(6, cpus))
    workers = max(1, int(workers))
    pool = ThreadPoolExecutor(max_workers=workers, thread_name_prefix="prep")
    prep_futures = {}
    for t_idx, _t, sn, _rc, set_imgs in candidates:
        imgs = set_imgs if set_imgs is not None else sn_folders.get(sn)
        for img_idx, path in enumerate(imgs or []):
            prep_futures[(t_idx, img_idx)] = pool.submit(prepare_photo, path, auto_rotate, smart_rotate,
                                                         max_image_px)
    log(f"  Preparing {len(prep_futures)} photo(s) with {workers} worker(s)"
        + (f", max {max_image_px}px" if max_image_px else ", original size"))

    for t_idx, table, sn, table_img_rc, set_imgs in candidates:
        if is_cancelled():
            log(f"\n[CANCELLED] Stopped by user at {current}/{total}.")
            break

        current += 1
        if progress_callback:
            progress_callback(current, total)

        if set_imgs is not None:
            images = set_imgs
        elif sn not in sn_folders:
            no_folder += 1
            skip_reasons[t_idx] = "no photo folder found"
            log(f"  [SKIP] {sn} (table {t_idx}): no photo folder found")
            continue
        else:
            images = sn_folders[sn]

        try:
            img_cell = table.rows[table_img_rc[0]].cells[table_img_rc[1]]
        except IndexError:
            err_msg = f"{sn}: table {t_idx} has no row {table_img_rc[0]}"
            skip_reasons[t_idx] = "photo cell missing in the table"
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
                fut = prep_futures.get((t_idx, img_idx))
                try:
                    if fut is None:
                        raise RuntimeError("not prepared")
                    img_path, notes = fut.result()
                    for note in notes:
                        log(f"      {note}")
                except Exception as e:
                    if fut is not None:
                        log(f"      Rotation Error for {os.path.basename(img_path)}: {e}")
                    if os.path.splitext(img_path)[1].lower() in ('.heic', '.heif'):
                        img_path = convert_heic_to_jpg(img_path)

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
            injected_tables[t_idx] = (sn, cell_img_count)
            injected += 1
            img_count += cell_img_count
            log(f"  [OK  ] {sn} (table {t_idx:>3}): {cell_img_count} image(s)")

    pool.shutdown(wait=True, cancel_futures=True)

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

    summary_file = None
    if write_summary:
        try:
            text = build_injection_summary(
                doc, docx_path=docx_path, photos_root=sn_root, output_path=output_path,
                saved=injected > 0, injected_tables=injected_tables, skip_reasons=skip_reasons,
                start_table=start_table, active_models=active_models, unused_folders=report_unused,
                unassigned_photos=report_unassigned, warnings=report_warnings, cancelled=cancelled,
                moved_photos=report_moved, moved_out=moved_out,
                images_total=img_count, errors_total=len(errors),
                empty_folders=find_empty_folders(sn_root))
            summary_file = summary_path_for(output_path)
            with open(summary_file, "w", encoding="utf-8-sig", newline="\r\n") as fh:
                fh.write(text)
            log(f"Summary -> {summary_file}")
        except Exception as e:
            summary_file = None
            log(f"  [WARN] Could not write summary: {e}")

    if cancelled:
        log("Process was cancelled by user.")
    else:
        log("Done!")

    result = InjectionResult(injected=injected, img_count=img_count, no_folder=no_folder, errors=errors)
    result["summary_path"] = summary_file
    return result


class TableSelectionDialog(tk.Toplevel):
    def __init__(self, parent, docx_path: str, current_sn_rc: tuple, current_img_rc: tuple, initial_table_idx: int = 0):
        super().__init__(parent)
        self.title("Configure Table Layout")
        _k = ui_scale(self)
        self.geometry(f"{min(int(960 * _k), self.winfo_screenwidth() - 40)}"
                      f"x{min(int(640 * _k), self.winfo_screenheight() - 80)}")
        self.minsize(800, 520)

        # Fallback theme colors to prevent KeyError in headless or partial environments
        default_colors = dict(UI_COLORS)
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
            bg=self.COLORS["card"], fg=self.COLORS["select_sn"], selectcolor=self.COLORS["card"],
            activebackground=self.COLORS["card"], activeforeground=self.COLORS["select_sn"],
            font=("Segoe UI", 10, "bold"), cursor="hand2"
        ).pack(side="left", padx=10, pady=5)

        tk.Radiobutton(
            controls, text="Select Image Cell (Green)", variable=self.mode, value="img",
            bg=self.COLORS["card"], fg=self.COLORS["select_img"], selectcolor=self.COLORS["card"],
            activebackground=self.COLORS["card"], activeforeground=self.COLORS["select_img"],
            font=("Segoe UI", 10, "bold"), cursor="hand2"
        ).pack(side="left", padx=10, pady=5)

        tk.Button(
            controls, text="Approve",
            bg=self.COLORS["btn_run"], fg=self.COLORS["btn_run_fg"],
            activebackground=self.COLORS["btn_run_hover"], activeforeground=self.COLORS["btn_run_fg"],
            relief="flat", bd=0, padx=16, pady=4, cursor="hand2",
            font=("Segoe UI", 10, "bold"), command=self._approve
        ).pack(side="right", padx=10, pady=5)

        tk.Button(
            controls, text="Cancel",
            bg=self.COLORS["soft"], fg=self.COLORS["soft_fg"],
            activebackground=self.COLORS["soft_hover"], activeforeground=self.COLORS["soft_fg"],
            relief="flat", bd=0, padx=16, pady=4, cursor="hand2",
            font=("Segoe UI", 10, "bold"), command=self.destroy
        ).pack(side="right")

        # ── Table Navigation Bar ──
        nav_bar = tk.Frame(self, bg=self.COLORS["card"])
        nav_bar.pack(fill="x", padx=10, pady=(0, 5))

        self.prev_btn = tk.Button(
            nav_bar, text="◀ Previous",
            bg=self.COLORS["soft"], fg=self.COLORS["soft_fg"], disabledforeground=self.COLORS["disabled_fg"],
            relief="flat", bd=0, padx=10, pady=3, cursor="hand2",
            font=("Segoe UI", 9, "bold"), activebackground=self.COLORS["soft_hover"],
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
            bg=self.COLORS["soft"], fg=self.COLORS["soft_fg"], disabledforeground=self.COLORS["disabled_fg"],
            relief="flat", bd=0, padx=10, pady=3, cursor="hand2",
            font=("Segoe UI", 9, "bold"), activebackground=self.COLORS["soft_hover"],
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
            bg=self.COLORS["card"], fg=self.COLORS["text"],
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
                fg=self.COLORS["btn_stop"], bg=self.COLORS["input_bg"],
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
                    bg=self.COLORS["card"], fg=self.COLORS["text"],
                    activebackground=self.COLORS["soft"], relief="solid", bd=1, cursor="hand2",
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
                btn.configure(bg=self.COLORS["select_sn"], fg="white")
            elif (r, c) == self.img_rc:
                btn.configure(bg=self.COLORS["select_img"], fg="white")
            else:
                btn.configure(bg=self.COLORS["card"], fg=self.COLORS["text"])

    def _approve(self):
        self.result = (self.sn_rc, self.img_rc)
        self.destroy()


# ─────────────────────────────────────────────
# GUI Application
# ─────────────────────────────────────────────
def enable_windows_dpi_awareness():
    """Sharp text on high-DPI screens (otherwise Windows stretches the window and it looks blurry).
    Must run before the first Tk window is created."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # system DPI aware
    except Exception:
        try:
            import ctypes
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def ui_scale(widget) -> float:
    """1.0 at 100% Windows scaling (96 dpi), 1.5 at 150%, ... used to size windows."""
    try:
        return max(1.0, float(widget.tk.call("tk", "scaling")) / (96 / 72))
    except Exception:
        return 1.0


UI_COLORS = {
    "bg":            "#eef2f7",   # window background (soft grey-blue)
    "card":          "#ffffff",   # cards
    "card_alt":      "#f1f5f9",   # subtle fills
    "border":        "#d5dce6",
    "text":          "#1e293b",   # main text (dark slate)
    "text_dim":      "#64748b",   # secondary text
    "input_bg":      "#f8fafc",
    "input_fg":      "#0f172a",
    "accent":        "#2563eb",   # blue: selection / links
    "accent_ok":     "#15803d",   # green text (success / info that is OK)
    "warn":          "#b45309",   # amber text (warnings)
    "btn_run":       "#16a34a",   # big green Run button
    "btn_run_hover": "#15803d",
    "btn_run_fg":    "#ffffff",
    "btn_stop":      "#dc2626",   # red Stop button (only while it can be pressed)
    "btn_stop_hover":"#b91c1c",
    "btn_stop_fg":   "#ffffff",
    "btn_pick":      "#2563eb",   # blue Browse buttons
    "btn_pick_hover":"#1d4ed8",
    "btn_pick_fg":   "#ffffff",
    "soft":          "#e0e7ff",   # light-blue secondary buttons
    "soft_hover":    "#c7d2fe",
    "soft_fg":       "#1e40af",
    "disabled_bg":   "#e2e8f0",
    "disabled_fg":   "#94a3b8",
    "select_sn":     "#2563eb",   # Configure Layout: S/N cell
    "select_img":    "#16a34a",   # Configure Layout: image cell
    "log_bg":        "#0f172a",   # dark console log
    "log_fg":        "#e2e8f0",
}


class App(tk.Tk):
    def __init__(self):
        enable_windows_dpi_awareness()
        super().__init__()

        self.title(f"{__app_name__} v{__version__}")
        k = ui_scale(self)
        screen_w, screen_h = self.winfo_screenwidth(), self.winfo_screenheight()
        width = min(int(940 * k), screen_w - int(40 * k))
        height = min(int(920 * k), screen_h - int(90 * k))
        self.geometry(f"{width}x{height}")
        self.minsize(min(int(780 * k), width), min(int(620 * k), height))

        # ── Colors ──
        # Light theme, one meaning per colour: green = run/success, red = stop/error,
        # blue = buttons/selection, amber = warnings. The log keeps a dark console look.
        self.COLORS = UI_COLORS.copy()
        self.configure(bg=self.COLORS["bg"])

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
                         foreground=self.COLORS["text_dim"],
                         font=("Segoe UI", 9))
        style.configure("Status.TLabel",
                         background=self.COLORS["bg"],
                         foreground=self.COLORS["accent_ok"],
                         font=("Segoe UI", 10, "bold"))
        style.configure("green.Horizontal.TProgressbar",
                         troughcolor=self.COLORS["disabled_bg"],
                         background=self.COLORS["btn_run"],
                         bordercolor=self.COLORS["disabled_bg"],
                         lightcolor=self.COLORS["btn_run"],
                         darkcolor=self.COLORS["btn_run"],
                         thickness=10)
        style.configure("Vertical.TScrollbar", background=self.COLORS["card_alt"],
                         troughcolor=self.COLORS["card"], bordercolor=self.COLORS["border"],
                         arrowcolor=self.COLORS["text_dim"])

        self._running = False
        self._cancel_event = threading.Event()
        self.sn_row_col = (5, 3)
        self.img_row_col = (12, 0)
        self.auto_rotate_var = tk.BooleanVar(value=True)
        self.smart_rotate_var = tk.BooleanVar(value=True)
        self.shrink_photos_var = tk.BooleanVar(value=True)
        self.auto_match_var = tk.BooleanVar(value=True)
        self.photo_ocr_var = tk.BooleanVar(value=True)
        self.set_fallback_var = tk.BooleanVar(value=True)
        self.verify_photos_var = tk.BooleanVar(value=False)  # slow: OCR of every photo

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
        tk.Button(
            title_row, text="❓  วิธีใช้", font=(self._thai_font(), 11, "bold"),
            bg=C["soft"], fg=C["soft_fg"], activebackground=C["soft_hover"],
            activeforeground=C["soft_fg"], relief="flat", bd=0, cursor="hand2", padx=16, pady=5,
            command=self._show_help
        ).pack(side="right")

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

        # ── Options card (two groups) ──
        self._make_options_card(inputs_frame)

        # ── Button Row ──
        btn_frame = ttk.Frame(self, style="BG.TFrame")
        btn_frame.pack(fill="x", padx=20, pady=(6, 4))

        self.run_btn = tk.Button(
            btn_frame,
            text="▶  Run Injection",
            font=("Segoe UI", 13, "bold"),
            bg=C["btn_run"], fg=C["btn_run_fg"],
            activebackground=C["btn_run_hover"], activeforeground=C["btn_run_fg"],
            disabledforeground=C["disabled_fg"],
            relief="flat", bd=0, cursor="hand2", padx=28, pady=9,
            command=self._on_run
        )
        self.run_btn.pack(side="left")
        self._add_hover(self.run_btn, C["btn_run"], C["btn_run_hover"])

        self.stop_btn = tk.Button(
            btn_frame,
            text="■  Stop",
            font=("Segoe UI", 13, "bold"),
            bg=C["disabled_bg"], fg=C["btn_stop_fg"],
            activebackground=C["btn_stop_hover"], activeforeground=C["btn_stop_fg"],
            disabledforeground=C["disabled_fg"],
            relief="flat", bd=0, cursor="hand2", padx=24, pady=9,
            command=self._on_stop,
            state="disabled"
        )
        self.stop_btn.pack(side="left", padx=(10, 0))

        self._last_summary = None
        self.summary_btn = tk.Button(
            btn_frame, text="📄  Summary", font=("Segoe UI", 11, "bold"),
            bg=C["disabled_bg"], fg=C["soft_fg"], disabledforeground=C["disabled_fg"],
            activebackground=C["soft_hover"], activeforeground=C["soft_fg"],
            relief="flat", bd=0, cursor="hand2", padx=16, pady=9,
            command=self._open_summary, state="disabled"
        )
        self.summary_btn.pack(side="right")

        self.status_label = ttk.Label(btn_frame, text="Ready", style="Status.TLabel")
        self.status_label.pack(side="left", padx=16)

        # ── Progress Bar ──
        self.progress = ttk.Progressbar(self, style="green.Horizontal.TProgressbar",
                                         orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", padx=20, pady=(2, 8))

        # ── Log Area (dark console) ──
        log_frame = tk.Frame(self, bg=C["log_bg"], bd=0, highlightthickness=1,
                             highlightbackground=C["border"])
        log_frame.pack(fill="both", expand=True, padx=20, pady=(0, 18))

        self.log_text = tk.Text(
            log_frame,
            bg=C["log_bg"], fg=C["log_fg"],
            font=("Cascadia Code", 9),
            insertbackground=C["log_fg"],
            selectbackground="#334155",
            relief="flat", wrap="word", height=8,
            padx=12, pady=8,
            state="disabled"
        )
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True)

        self.log_text.tag_configure("ok",       foreground="#4ade80")
        self.log_text.tag_configure("err",      foreground="#f87171")
        self.log_text.tag_configure("warn",     foreground="#fbbf24")
        self.log_text.tag_configure("info",     foreground="#94a3b8")
        self.log_text.tag_configure("title",    foreground="#f1f5f9", font=("Cascadia Code", 9, "bold"))
        self.log_text.tag_configure("cancel",   foreground="#fb923c")

        self._auto_detect_paths()
        self.kw_entry.configure(state="disabled" if self.auto_match_var.get() else "normal")

    def _add_hover(self, btn, normal_bg, hover_bg):
        """Darken a flat button while the mouse is over it (only when it is enabled)."""
        def enter(_e):
            if str(btn["state"]) != "disabled":
                btn.configure(bg=hover_bg)

        def leave(_e):
            if str(btn["state"]) != "disabled":
                btn.configure(bg=normal_bg)
        btn.bind("<Enter>", enter)
        btn.bind("<Leave>", leave)

    def _make_path_label(self, parent, var):
        """Selected path, or a grey hint when nothing is selected yet."""
        C = self.COLORS
        lbl = tk.Label(parent, bg=C["card"], anchor="w")
        lbl.pack(fill="x", pady=(4, 0))

        def refresh(*_):
            value = var.get().strip()
            if value:
                lbl.configure(text=value, fg=C["text_dim"], font=("Consolas", 9))
            else:
                lbl.configure(text="ยังไม่ได้เลือก — กด Browse", fg=C["disabled_fg"],
                              font=(self._thai_font(), 9, "italic"))
        var.trace_add("write", refresh)
        refresh()
        return lbl

    def _option_check(self, parent, text, var, command=None, sub=False):
        C = self.COLORS
        cb = tk.Checkbutton(
            parent, text=text, variable=var, command=command,
            bg=C["card"], fg=C["text"] if not sub else C["text_dim"],
            selectcolor=C["card"], activebackground=C["card"], activeforeground=C["accent"],
            disabledforeground=C["disabled_fg"], font=("Segoe UI", 10 if not sub else 9),
            anchor="w", bd=0, highlightthickness=0, cursor="hand2"
        )
        cb.pack(anchor="w", padx=(22 if sub else 0, 0), pady=1)
        return cb

    def _make_options_card(self, parent):
        """Options in two labelled groups: photo processing and S/N matching."""
        C = self.COLORS
        card = tk.Frame(parent, bg=C["card"], bd=0, highlightthickness=1,
                        highlightbackground=C["border"])
        card.pack(fill="x", pady=4)
        inner = tk.Frame(card, bg=C["card"])
        inner.pack(fill="x", padx=14, pady=10)
        inner.columnconfigure(0, weight=1, uniform="opt")
        inner.columnconfigure(1, weight=1, uniform="opt")

        left = tk.Frame(inner, bg=C["card"])
        left.grid(row=0, column=0, sticky="nw")
        right = tk.Frame(inner, bg=C["card"])
        right.grid(row=0, column=1, sticky="nw", padx=(12, 0))

        tk.Label(left, text="📷  Photos", bg=C["card"], fg=C["text"],
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 4))
        self._option_check(left, "Auto-Rotate (EXIF)", self.auto_rotate_var)
        self._option_check(left, "Smart Rotate (OCR, makes text upright)", self.smart_rotate_var)
        self._option_check(left, f"Shrink photos to {INSERT_MAX_PX}px (faster, smaller)",
                           self.shrink_photos_var)

        tk.Label(right, text="🔎  Matching", bg=C["card"], fg=C["text"],
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 4))
        self._option_check(right, "Auto-match S/N in tables (recommended)",
                           self.auto_match_var, command=self._on_auto_match_toggle)
        self.photo_ocr_cb = self._option_check(
            right, "Read S/N from photos (OCR)",
            self.photo_ocr_var, sub=True)
        self.set_fallback_cb = self._option_check(
            right, "Fallback: ชุดที่ folders (12-280)", self.set_fallback_var, sub=True)
        self.verify_photos_cb = self._option_check(
            right, "Verify photos (OCR every folder, slower)", self.verify_photos_var, sub=True)

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
                        font=("Segoe UI", 9, "bold"),
                        bg=C["btn_pick"], fg=C["btn_pick_fg"],
                        activebackground=C["btn_pick_hover"], activeforeground=C["btn_pick_fg"],
                        relief="flat", bd=0, cursor="hand2",
                        padx=16, pady=4,
                        command=command)
        btn.pack(side="right")
        self._add_hover(btn, C["btn_pick"], C["btn_pick_hover"])
        
        if with_config:
            self.config_btn = tk.Button(top, text="⚙ Configure Layout",
                                        font=("Segoe UI", 9),
                                        bg=C["soft"], fg=C["soft_fg"],
                                        activebackground=C["soft_hover"], activeforeground=C["soft_fg"],
                                        relief="flat", bd=0, cursor="hand2",
                                        padx=12, pady=4,
                                        command=self._configure_layout)
            self.config_btn.pack(side="right", padx=(0, 8))
            self._add_hover(self.config_btn, C["soft"], C["soft_hover"])

        self._make_path_label(inner, var)
        
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
                bg=C["input_bg"], fg=C["text"],
                buttonbackground=C["card_alt"],
                insertbackground=C["text"],
                relief="flat",
                highlightthickness=1,
                highlightbackground=C["border"],
                highlightcolor=C["accent"],
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
                        font=("Segoe UI", 9, "bold"),
                        bg=C["btn_pick"], fg=C["btn_pick_fg"],
                        activebackground=C["btn_pick_hover"], activeforeground=C["btn_pick_fg"],
                        relief="flat", bd=0, cursor="hand2",
                        padx=16, pady=4,
                        command=self._pick_sn_root)
        btn.pack(side="right")
        self._add_hover(btn, C["btn_pick"], C["btn_pick_hover"])

        # Keyword entry inline
        kw_frame = tk.Frame(top, bg=C["card"])
        kw_frame.pack(side="right", padx=(0, 10))

        tk.Label(kw_frame, text="Keyword:",
                 bg=C["card"], fg=C["text_dim"],
                 font=("Segoe UI", 9)).pack(side="left")

        self.kw_entry = kw_entry = tk.Entry(kw_frame,
                            textvariable=self.keyword_var,
                            bg=C["input_bg"], fg=C["text"],
                            font=("Consolas", 10, "bold"),
                            insertbackground=C["text"],
                            relief="flat", width=12,
                            disabledbackground=C["disabled_bg"],
                            disabledforeground=C["disabled_fg"],
                            highlightthickness=1,
                            highlightbackground=C["border"],
                            highlightcolor=C["accent"])
        kw_entry.pack(side="left", padx=(5, 0))

        tk.Label(kw_frame, text="*",
                 bg=C["card"], fg=C["text_dim"],
                 font=("Consolas", 10, "bold")).pack(side="left")

        # Path display
        self._make_path_label(inner, self.sn_root_var)

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
            bg=C["soft"], fg=C["soft_fg"], disabledforeground=C["disabled_fg"],
            activebackground=C["soft_hover"], activeforeground=C["soft_fg"],
            relief="flat", bd=0, cursor="hand2", padx=8, pady=2,
            command=self._select_all_models,
            state="disabled"
        )
        self.btn_select_all.pack(side="right", padx=(4, 0))

        self.btn_deselect_all = tk.Button(
            models_header, text="Deselect All",
            font=("Segoe UI", 8),
            bg=C["soft"], fg=C["soft_fg"], disabledforeground=C["disabled_fg"],
            activebackground=C["soft_hover"], activeforeground=C["soft_fg"],
            relief="flat", bd=0, cursor="hand2", padx=8, pady=2,
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
                bg=self.COLORS["card"], fg=self.COLORS["text_dim"],
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
                bg=self.COLORS["card"],
                fg=self.COLORS["text"],
                selectcolor=self.COLORS["card"],
                activebackground=self.COLORS["card"],
                activeforeground=self.COLORS["accent"],
                font=("Segoe UI", 9), bd=0, highlightthickness=0, cursor="hand2",
                command=self._on_model_toggle
            )
            n_cols = 3 if max(len(m) for m in models) <= 40 else 2
            row = idx // n_cols
            col = idx % n_cols
            cb.grid(row=row, column=col, sticky="w", padx=(0, 12), pady=1)

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
            elif "[WARN]" in msg:
                tag = "warn"
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
            C = self.COLORS
            if running:
                self.run_btn.configure(state="disabled", bg=C["disabled_bg"])
                self.stop_btn.configure(state="normal", bg=C["btn_stop"])
                self._add_hover(self.stop_btn, C["btn_stop"], C["btn_stop_hover"])
            else:
                self.run_btn.configure(state="normal", bg=C["btn_run"])
                self.stop_btn.configure(state="disabled", bg=C["disabled_bg"])
        self.after(0, _update)

    def _open_summary(self):
        """Open the last run's '<output>_summary.txt' with the default text editor."""
        path = getattr(self, "_last_summary", None)
        if not path or not os.path.isfile(path):
            messagebox.showinfo("Summary", "No summary yet - run the injection first.")
            return
        try:
            if sys.platform == "win32":
                os.startfile(path)  # noqa: S606 (opens in Notepad / default editor)
            else:
                import subprocess
                subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", path])
        except Exception as e:
            messagebox.showerror("Summary", f"Could not open {path}:\n{e}")

    def _thai_font(self) -> str:
        """A font that renders Thai well on this OS."""
        cached = getattr(self, "_thai_font_name", None)
        if cached:
            return cached
        try:
            from tkinter import font as tkfont
            families = set(tkfont.families(self))
        except Exception:
            families = set()
        for name in ("Leelawadee UI", "Leelawadee", "Tahoma", "Thonburi", "Sukhumvit Set",
                     "Noto Sans Thai", "Segoe UI"):
            if name in families:
                self._thai_font_name = name
                return name
        self._thai_font_name = "TkDefaultFont"
        return self._thai_font_name

    def _show_help(self):
        """Open (or focus) the Thai usage guide window."""
        existing = getattr(self, "_help_win", None)
        if existing is not None and existing.winfo_exists():
            existing.deiconify()
            existing.lift()
            existing.focus_force()
            return
        from help_text import HELP_TITLE, HELP_SECTIONS
        C = self.COLORS
        f = self._thai_font()

        win = tk.Toplevel(self)
        win.title(HELP_TITLE)
        win.configure(bg=C["bg"])
        k = ui_scale(self)
        win.geometry(f"{min(int(860 * k), self.winfo_screenwidth() - 40)}"
                     f"x{min(int(740 * k), self.winfo_screenheight() - int(90 * k))}")
        win.minsize(int(560 * k), int(420 * k))
        self._help_win = win

        bottom = tk.Frame(win, bg=C["bg"])
        bottom.pack(side="bottom", fill="x", padx=14, pady=(0, 12))  # packed first so it never gets squeezed out
        frame = tk.Frame(win, bg=C["card"], highlightthickness=1, highlightbackground=C["border"])
        frame.pack(fill="both", expand=True, padx=14, pady=(14, 8))
        PADX = 22
        # wrap="none": Tk's own word-wrap mis-measures shaped Thai text and can split a consonant
        # from its vowel/tone mark (shown as ◌่). Lines are broken here instead, only at spaces,
        # using the real font width, and re-laid out when the window is resized.
        text = tk.Text(frame, bg=C["card"], fg=C["text"], wrap="none", relief="flat", height=10,
                       padx=PADX, pady=16, font=(f, 11), cursor="arrow", highlightthickness=0)
        sb = ttk.Scrollbar(frame, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        text.pack(side="left", fill="both", expand=True)

        from tkinter import font as tkfont
        # kind: (font, colour, first-line indent, continuation indent, space above, space below, bg)
        styles = {
            "h1":   ((f, 16, "bold"), C["text"], 0, 0, 14, 8, None),
            "h2":   ((f, 12, "bold"), C["accent"], 0, 0, 10, 4, None),
            "p":    ((f, 11), C["text"], 0, 0, 2, 4, None),
            "li":   ((f, 11), C["text"], 14, 30, 2, 4, None),
            "step": ((f, 12, "bold"), C["text"], 10, 36, 4, 4, None),
            "tip":  ((f, 11), C["accent_ok"], 10, 26, 6, 4, None),
            "warn": ((f, 11), C["warn"], 10, 32, 6, 4, None),
            "code": (("Consolas", 10), "#334155", 24, 24, 4, 6, C["card_alt"]),
            "gap":  ((f, 6), C["text"], 0, 0, 0, 0, None),
        }
        fonts = {k: tkfont.Font(win, font=v[0]) for k, v in styles.items()}
        made_tags = set()

        def tag_for(kind, first, last):
            name = f"{kind}:{int(first)}{int(last)}"
            if name not in made_tags:
                font, fg, ind1, ind2, sp1, sp3, bg = styles[kind]
                indent = ind1 if first else ind2
                opts = dict(font=font, foreground=fg, lmargin1=indent, lmargin2=indent,
                            spacing1=sp1 if first else 1, spacing3=sp3 if last else 1)
                if bg:
                    opts["background"] = bg
                text.tag_configure(name, **opts)
                text.tag_configure(name + "~", **opts)  # identical twin, see insert_line()
                made_tags.add(name)
            return name

        def insert_line(line, tag):
            """
            Tk on Windows draws text in 200-byte (UTF-8) pieces; a piece boundary inside a Thai
            cluster splits a consonant from its tone/vowel mark. Insert the line in <=180-byte
            pieces cut at spaces, alternating two identical tags so Tk draws each piece separately.
            """
            if len(line.encode("utf-8")) <= 180:
                text.insert("end", line + "\n", tag)
                return
            pieces, cur = [], ""
            for word in line.split(" "):
                cand = word if not cur else cur + " " + word
                if cur and len(cand.encode("utf-8")) > 180:
                    pieces.append(cur + " ")
                    cur = word
                else:
                    cur = cand
            pieces.append(cur)
            for k, piece in enumerate(pieces):
                text.insert("end", piece, tag if k % 2 == 0 else tag + "~")
            text.insert("end", "\n", tag)

        def wrap_line(line, fnt, first_w, cont_w):
            if fnt.measure(line) <= first_w:
                return [line]
            out, cur, width = [], "", first_w
            for word in line.split(" "):
                cand = word if not cur else cur + " " + word
                if cur and fnt.measure(cand) > width:
                    out.append(cur)
                    cur, width = word, cont_w
                else:
                    cur = cand
            out.append(cur)
            return out

        layout_state = {"width": None, "job": None}

        def layout(force=False):
            layout_state["job"] = None
            avail = text.winfo_width() - 2 * PADX - 12
            if avail < 100 or (not force and avail == layout_state["width"]):
                return
            layout_state["width"] = avail
            top = text.yview()[0]
            text.configure(state="normal")
            text.delete("1.0", "end")
            for kind, content in HELP_SECTIONS:
                _, _, ind1, ind2, *_ = styles[kind]
                fnt = fonts[kind]
                lines = []
                for i, para in enumerate(content.split("\n")):
                    if kind == "code":
                        lines.append(para)
                    else:
                        lines += wrap_line(para, fnt, avail - (ind1 if i == 0 else ind2), avail - ind2)
                for i, ln in enumerate(lines):
                    insert_line(ln, tag_for(kind, i == 0, i == len(lines) - 1))
            text.configure(state="disabled")
            text.yview_moveto(top)

        def on_resize(_event=None):
            if layout_state["job"]:
                win.after_cancel(layout_state["job"])
            layout_state["job"] = win.after(60, layout)

        text.bind("<Configure>", on_resize)
        win.update_idletasks()
        layout(force=True)

        tk.Label(bottom, text=f"{__app_name__}  v{__version__}", bg=C["bg"], fg=C["text_dim"],
                 font=(f, 9)).pack(side="left")
        tk.Button(bottom, text="ปิด", font=(f, 11, "bold"), bg=C["btn_run"], fg=C["btn_run_fg"],
                  activebackground=C["btn_run_hover"], activeforeground="#fff", relief="flat", bd=0,
                  cursor="hand2", padx=22, pady=4, command=win.destroy).pack(side="right")
        win.bind("<Escape>", lambda e: win.destroy())
        win.focus_force()

    def _on_auto_match_toggle(self):
        auto = self.auto_match_var.get()
        entry = getattr(self, "kw_entry", None)
        if entry is not None:
            entry.configure(state="disabled" if auto else "normal")
        for name in ("photo_ocr_cb", "set_fallback_cb", "verify_photos_cb"):
            cb = getattr(self, name, None)
            if cb is not None:
                cb.configure(state="normal" if auto else "disabled")
        if auto:
            self._log("Auto-match ON: serials are found anywhere in each table; keyword and S/N cell are ignored.")
        else:
            self._log("Auto-match OFF: using keyword and configured S/N / Image cells.")

    def _on_stop(self):
        if self._running:
            self._cancel_event.set()
            self.status_label.configure(text="Stopping...", foreground=self.COLORS["warn"])
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
        auto_match = self.auto_match_var.get()
        if not keyword and not auto_match:
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

                result = run_injection(
                    docx, sn_root, output,
                    keyword=keyword,
                    sn_row_col=self.sn_row_col, img_row_col=self.img_row_col,
                    auto_rotate=self.auto_rotate_var.get(),
                    smart_rotate=self.smart_rotate_var.get(),
                    max_image_px=INSERT_MAX_PX if self.shrink_photos_var.get() else 0,
                    start_table=start_table,
                    selected_models=selected_models,
                    auto_match=auto_match,
                    photo_ocr=self.photo_ocr_var.get(),
                    set_number_fallback=self.set_fallback_var.get(),
                    verify_photos=self.verify_photos_var.get(),
                    log_callback=self._log,
                    progress_callback=self._update_progress,
                    cancel_event=self._cancel_event
                )
                injected, img_count, no_folder, errors = result
                summary_file = result.get("summary_path")

                def finish():
                    self._set_running(False)
                    self._last_summary = summary_file
                    if summary_file and os.path.isfile(summary_file):
                        self.summary_btn.configure(state="normal", bg=self.COLORS["soft"])
                    cancelled = self._cancel_event.is_set()
                    if cancelled:
                        self.status_label.configure(
                            text=f"Cancelled: {injected} tables saved before stop",
                            foreground=self.COLORS["warn"])
                    elif errors:
                        self.status_label.configure(
                            text=f"Done: {injected} tables, {img_count} images, {len(errors)} error(s)",
                            foreground=self.COLORS["btn_stop"])
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
def cli_main(argv: list[str]) -> int:
    """
    Headless run (Auto mode), e.g. for batch jobs or testing the built .exe:
      app.exe --cli --docx DOC.docx --photos PHOTO_ROOT --out OUT.docx [--log run.log]
              [--no-ocr] [--no-set-fallback] [--no-smart-rotate] [--start-table N] [--verify]
    The windowed .exe has no console, so use --log to see progress. Exit code 0 = images injected.
    """
    import argparse
    p = argparse.ArgumentParser(prog="Unit_Test_Image_Injector --cli")
    p.add_argument("--cli", action="store_true")
    p.add_argument("--docx", required=True)
    p.add_argument("--photos", required=True)
    p.add_argument("--out", required=True)
    p.add_argument("--log")
    p.add_argument("--start-table", type=int, default=1)
    p.add_argument("--no-ocr", action="store_true")
    p.add_argument("--no-set-fallback", action="store_true")
    p.add_argument("--no-smart-rotate", action="store_true")
    p.add_argument("--verify", action="store_true",
                   help="OCR every S/N / set folder and move photos of another device (slower)")
    p.add_argument("--original-size", action="store_true", help="embed photos without shrinking")
    p.add_argument("--workers", type=int, default=None)
    args = p.parse_args(argv)

    import time
    log_file = open(args.log, "w", encoding="utf-8") if args.log else None
    t0 = time.time()

    def log(msg):
        line = f"[{time.time() - t0:7.1f}s] {msg}"
        if log_file:
            log_file.write(line + "\n")
            log_file.flush()
        else:
            try:
                print(line, flush=True)
            except Exception:
                pass

    try:
        try:
            import pillow_heif
            pillow_heif.register_heif_opener()
        except Exception:
            pass
        res = run_injection(args.docx, args.photos, args.out, auto_match=True,
                            photo_ocr=not args.no_ocr, set_number_fallback=not args.no_set_fallback,
                            verify_photos=args.verify,
                            start_table=args.start_table, auto_rotate=True,
                            smart_rotate=not args.no_smart_rotate,
                            max_image_px=0 if args.original_size else INSERT_MAX_PX,
                            workers=args.workers, log_callback=log)
        if res.get("summary_path"):
            log(f"SUMMARY {res['summary_path']}")
        log(f"RESULT injected={res.injected} images={res.img_count} "
            f"no_folder={res.no_folder} errors={len(res.errors)}")
        return 0 if res.injected > 0 else 1
    except Exception as e:
        import traceback
        log(f"FATAL: {e}\n{traceback.format_exc()}")
        return 2
    finally:
        cleanup_temp_files()
        if log_file:
            log_file.close()


if __name__ == "__main__":
    if "--cli" in sys.argv[1:]:
        sys.exit(cli_main(sys.argv[1:]))
    app = App()
    app.mainloop()
