"""
Unit tests for ocr_engine.py (OpenThai2.0-OCR-Text backend).
Self-contained: renders its own test image, no user photos needed.

    python -m pytest test/test_ocr_engine.py -v
"""
import sys
from pathlib import Path

import numpy as np
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

cv2 = pytest.importorskip("cv2")
from PIL import Image  # noqa: E402

from ocr_engine import OpenThaiEngine, detect_orientation, get_engine  # noqa: E402

LINES = ["Phone Information", "MAC Address: BC:AB:F5:A2:C5:35",
         "Host Name: SEPBCABF5A2C535", "Active Load ID: sip78xx.14-1-1"]


@pytest.fixture(scope="module")
def engine():
    eng = get_engine("openthai")
    assert isinstance(eng, OpenThaiEngine)
    return eng


@pytest.fixture(scope="module")
def page():
    img = np.full((420, 1200, 3), 255, np.uint8)
    for i, t in enumerate(LINES):
        cv2.putText(img, t, (40, 80 + i * 90), cv2.FONT_HERSHEY_SIMPLEX, 1.5, (0, 0, 0), 3, cv2.LINE_AA)
    return Image.fromarray(img)


def test_reads_english_text(engine, page):
    res = engine.read(page)
    joined = " ".join(res.texts()).upper()
    assert "PHONE" in joined and "MAC" in joined
    assert res.score() > 30


def test_extracts_mac_from_ocr(engine, page):
    import re
    text = " ".join(engine.read(page).texts(include_dropped=True)).upper()
    cleaned = re.sub(r"[^0-9A-F]", "", text)
    assert "BCABF5A2C535" in cleaned


@pytest.mark.parametrize("applied", [0, 90, 180, 270])
def test_orientation(engine, page, applied):
    angle, score = detect_orientation(page.rotate(applied, expand=True), engine)
    assert angle == (360 - applied) % 360, f"applied {applied}, got {angle} (score {score:.1f})"


def test_blank_image_returns_zero(engine):
    blank = Image.new("RGB", (800, 600), "white")
    assert detect_orientation(blank, engine)[0] == 0



def test_model_download_is_pinned_to_v2_0_0(tmp_path, monkeypatch):
    """A missing model is downloaded from the v2.0.0 tag, never the repository's moving main (now 2.1.x)."""
    import types
    import ocr_engine
    calls = {}

    def fake_download(**kw):
        calls.update(kw)
        for f in ocr_engine.REQUIRED_FILES:
            dst = tmp_path / f
            dst.parent.mkdir(parents=True, exist_ok=True)
            dst.write_text("x")
    monkeypatch.setitem(sys.modules, "huggingface_hub", types.SimpleNamespace(snapshot_download=fake_download))
    ocr_engine.ensure_model(tmp_path)
    assert calls["revision"] == "v2.0.0"
    assert calls["repo_id"] == ocr_engine.HF_REPO_ID
