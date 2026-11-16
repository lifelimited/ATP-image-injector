"""
ocr_engine.py — Pluggable OCR backend for the OCR platform
==========================================================

Default backend: **OpenThai2.0-OCR-Text** (iApp Technology, Apache-2.0)
    https://huggingface.co/iapp/OpenThai2.0-OCR-Text
    - Thai + English printed text, CPU only (ONNX Runtime or OpenVINO)
    - ~1M params (det 1.8 MB + rec 2.4 MB), ~0.1-0.4 s per photo on CPU
    - Built-in page orientation (90/270), deskew and reading order

Legacy backend: **EasyOCR** (PyTorch, optional) — kept as a fallback.

Select with the environment variable ``OCR_BACKEND`` = ``openthai`` (default)
or ``easyocr``. Override the model folder with ``OPENTHAI_OCR_DIR``.
If the model files are missing they are downloaded from Hugging Face once.

Public API
----------
    eng = get_engine(log_cb=print)        # thread-safe singleton
    res = eng.read(pil_or_ndarray_or_path) # OCRResult(text, lines, rotation)
    angle, score = detect_orientation(pil_img, eng)  # PIL CCW angle to fix
"""

from __future__ import annotations

import importlib.util
import os
import sys
import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Optional

def _resolve_default_model_dir() -> Path:
    """Locate bundled models/OpenThai2.0-OCR-Text in both source and PyInstaller _MEIPASS runs."""
    candidates = []
    if getattr(sys, "_MEIPASS", None):
        candidates.append(Path(sys._MEIPASS) / "models" / "OpenThai2.0-OCR-Text")
    if getattr(sys, "frozen", False):
        candidates.append(Path(sys.executable).resolve().parent / "models" / "OpenThai2.0-OCR-Text")
    candidates.append(Path(__file__).resolve().parent / "models" / "OpenThai2.0-OCR-Text")
    for c in candidates:
        if (c / "det.onnx").exists() and (c / "rec.onnx").exists():
            return c
    return candidates[-1]


HF_REPO_ID = "iapp/OpenThai2.0-OCR-Text"
DEFAULT_MODEL_DIR = _resolve_default_model_dir()
REQUIRED_FILES = ("det.onnx", "rec.onnx", "rec_dict.txt", "config.json",
                  "runtime/__init__.py", "runtime/line_engine.py", "runtime/thai_norm.py")

LogCb = Optional[Callable[[str], None]]


# ─────────────────────────────────────────────
# Result types
# ─────────────────────────────────────────────
@dataclass
class OCRLine:
    text: str
    confidence: float
    box: list = field(default_factory=list)
    dropped: bool = False          # filtered by the engine (low confidence / junk)


@dataclass
class OCRResult:
    text: str
    lines: list[OCRLine]
    rotation: int = 0              # clockwise quarter-turn the engine applied (0/90/180/270)

    def texts(self, include_dropped: bool = False) -> list[str]:
        return [ln.text for ln in self.lines if include_dropped or not ln.dropped]

    def score(self, min_conf: float = 0.4, min_len: int = 2) -> float:
        """Readable-text volume: sum(chars x confidence) of confident lines."""
        s = 0.0
        for ln in self.lines:
            t = ln.text.replace(" ", "")
            if not ln.dropped and ln.confidence >= min_conf and len(t) >= min_len:
                s += len(t) * ln.confidence
        return s


# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────
def _to_pil(image: Any):
    from PIL import Image
    import numpy as np
    if isinstance(image, Image.Image):
        return image.convert("RGB")
    if isinstance(image, (str, os.PathLike)):
        # PIL handles non-ASCII (Thai) Windows paths; cv2.imread does not.
        with Image.open(image) as im:
            return im.convert("RGB")
    if isinstance(image, np.ndarray):  # assume RGB (PIL/easyocr convention)
        if image.ndim == 2:
            return Image.fromarray(image).convert("RGB")
        return Image.fromarray(image[:, :, :3])
    raise TypeError(f"Unsupported image type: {type(image)!r}")


def ensure_model(model_dir: Path, log_cb: LogCb = None) -> Path:
    """Make sure the OpenThai model files exist; download them if missing."""
    missing = [f for f in REQUIRED_FILES if not (model_dir / f).exists()]
    if not missing:
        return model_dir
    if log_cb:
        log_cb(f"      Downloading {HF_REPO_ID} (~4 MB) to {model_dir} ...")
    try:
        from huggingface_hub import snapshot_download
    except ImportError as e:
        raise RuntimeError(
            f"OpenThai model files missing ({', '.join(missing)}) and huggingface_hub "
            f"is not installed. Run: pip install huggingface_hub  — or download "
            f"https://huggingface.co/{HF_REPO_ID} into {model_dir}") from e
    snapshot_download(repo_id=HF_REPO_ID, local_dir=str(model_dir),
                      ignore_patterns=["pytorch/*", "*.pt"])
    still = [f for f in REQUIRED_FILES if not (model_dir / f).exists()]
    if still:
        raise RuntimeError(f"Download incomplete, still missing: {', '.join(still)}")
    return model_dir


def _load_runtime(model_dir: Path):
    """Import the model's bundled ``runtime`` package under a unique name
    (the folder name 'OpenThai2.0-OCR-Text' is not importable, and a bare
    ``runtime`` package name could clash with other modules)."""
    pkg_name = "openthai_ocr_runtime"
    if f"{pkg_name}.line_engine" in sys.modules:
        return sys.modules[f"{pkg_name}.line_engine"]
    rt_dir = model_dir / "runtime"
    spec = importlib.util.spec_from_file_location(
        pkg_name, rt_dir / "__init__.py", submodule_search_locations=[str(rt_dir)])
    pkg = importlib.util.module_from_spec(spec)
    sys.modules[pkg_name] = pkg
    spec.loader.exec_module(pkg)
    return importlib.import_module(f"{pkg_name}.line_engine")


# ─────────────────────────────────────────────
# Backends
# ─────────────────────────────────────────────
class OpenThaiEngine:
    """OpenThai2.0-OCR-Text (CPU, ONNX Runtime / OpenVINO)."""
    name = "OpenThai2.0-OCR-Text"
    auto_orients = True            # handles 90/270 itself; 180 is checked by detect_orientation

    def __init__(self, model_dir: Path | str | None = None, threads: int | None = None,
                 log_cb: LogCb = None):
        model_dir = Path(model_dir or os.environ.get("OPENTHAI_OCR_DIR") or DEFAULT_MODEL_DIR)
        ensure_model(model_dir, log_cb)
        line_engine = _load_runtime(model_dir)
        threads = threads or max(1, min(4, (os.cpu_count() or 2) // 2))
        self._ocr = line_engine.LineOCR(model_dir, threads=threads)
        self.backend = self._ocr.backend
        self._lock = threading.Lock()  # one inference at a time per engine instance

    def read(self, image: Any) -> OCRResult:
        import numpy as np
        pil = _to_pil(image)
        bgr = np.ascontiguousarray(np.asarray(pil)[:, :, ::-1])
        with self._lock:
            r = self._ocr.read(bgr)
        lines = [OCRLine(l["text"], float(l["confidence"]), l.get("box", [])) for l in r["lines"]]
        lines += [OCRLine(l["text"], float(l["confidence"]), l.get("box", []), dropped=True)
                  for l in r.get("dropped", []) if l.get("text")]
        return OCRResult(text=r["text"], lines=lines, rotation=int(r.get("rotation", 0)))


class EasyOCREngine:
    """Legacy EasyOCR backend (PyTorch). Install with: pip install easyocr"""
    name = "EasyOCR"
    auto_orients = False

    def __init__(self, langs=("en",), log_cb: LogCb = None):
        import easyocr
        try:
            self._reader = easyocr.Reader(list(langs), gpu=True)
            self.backend = "gpu"
        except Exception as e:  # pragma: no cover - hardware dependent
            if log_cb:
                log_cb(f"      [OCR Warning] GPU failed ({e}). Falling back to CPU...")
            self._reader = easyocr.Reader(list(langs), gpu=False)
            self.backend = "cpu"
        self._lock = threading.Lock()

    def read(self, image: Any) -> OCRResult:
        import numpy as np
        arr = np.asarray(_to_pil(image))
        with self._lock:
            res = self._reader.readtext(arr)
        lines = [OCRLine(t, float(p), [list(map(float, pt)) for pt in b]) for b, t, p in res]
        return OCRResult(text="\n".join(l.text for l in lines), lines=lines, rotation=0)


# ─────────────────────────────────────────────
# Singleton factory
# ─────────────────────────────────────────────
_engine = None
_engine_lock = threading.Lock()


def get_engine(backend: str | None = None, log_cb: LogCb = None):
    """Return the shared OCR engine (created on first use)."""
    global _engine
    backend = (backend or os.environ.get("OCR_BACKEND") or "openthai").lower()
    with _engine_lock:
        if _engine is not None and getattr(_engine, "_key", None) == backend:
            return _engine
        if backend in ("openthai", "openthai2", "iapp"):
            if log_cb:
                log_cb("      Initializing OpenThai2.0-OCR-Text (CPU)...")
            eng = OpenThaiEngine(log_cb=log_cb)
        elif backend == "easyocr":
            if log_cb:
                log_cb("      Initializing EasyOCR (legacy)... this might take a moment.")
            eng = EasyOCREngine(log_cb=log_cb)
        else:
            raise ValueError(f"Unknown OCR_BACKEND '{backend}' (use 'openthai' or 'easyocr')")
        eng._key = backend
        if log_cb:
            log_cb(f"      OCR ready: {eng.name} [{eng.backend}]")
        _engine = eng
        return eng


# ─────────────────────────────────────────────
# Orientation detection
# ─────────────────────────────────────────────
def detect_orientation(pil_img, engine=None, min_score: float = 15.0,
                       max_size: int = 1600) -> tuple[int, float]:
    """
    Find the rotation that makes the text upright.

    Returns ``(angle, score)`` where ``angle`` is the counter-clockwise angle to
    pass to ``PIL.Image.rotate(angle, expand=True)`` (0/90/180/270), and
    ``score`` the readable-text volume at that orientation. If the best score
    is below ``min_score`` the image is treated as text-less and 0 is returned.
    """
    engine = engine or get_engine()
    img = _to_pil(pil_img)
    if max(img.size) > max_size:
        img = img.copy()
        img.thumbnail((max_size, max_size))

    if getattr(engine, "auto_orients", False):
        # OpenThai turns sideways pages itself (reports a clockwise rotation)
        # but only flips 180 on pages that read badly — so compare as-is vs
        # pre-flipped and keep whichever reads better.
        r0 = engine.read(img)
        r1 = engine.read(img.rotate(180, expand=True))
        s0, s1 = r0.score(), r1.score()
        if s1 > s0 * 1.15 and s1 > s0 + 5:
            angle, best = (180 - r1.rotation) % 360, s1
        else:
            angle, best = (-r0.rotation) % 360, s0
    else:
        angle, best = 0, -1.0
        for a in (0, 90, 180, 270):
            s = engine.read(img if a == 0 else img.rotate(a, expand=True)).score()
            if s > best:
                angle, best = a, s

    if best < min_score:
        return 0, best
    return angle, best

