"""
Unit tests for Device Model Detection and Filtering.
Tests:
- detect_docx_models()
- detect_folder_models()
- model_matches()
- run_injection(selected_models=...)
"""
import sys
from pathlib import Path
import pytest
from docx import Document
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
TEST_DIR = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
if str(TEST_DIR) not in sys.path:
    sys.path.insert(0, str(TEST_DIR))

import app
from test_table_selection import create_mock_docx, create_mock_photo_root, count_drawings_in_cell


def test_detect_docx_models_multi_model(tmp_path):
    """detect_docx_models should extract unique, sorted device models from document tables."""
    docx_path = tmp_path / "multi_models.docx"
    specs = [
        {"sn": "WZP001", "model": "CP-7821-K9="},
        {"sn": "FVH002", "model": "CP-8841-K9="},
        {"sn": "WZP003", "model": "CP-7821-K9="},
        {"sn": "FCH004", "model": "CP-8865-K9="},
    ]
    create_mock_docx(docx_path, specs)

    detected = app.detect_docx_models(str(docx_path))
    assert isinstance(detected, list)
    assert detected == ["CP-7821-K9=", "CP-8841-K9=", "CP-8865-K9="]


def test_detect_docx_models_empty_doc(tmp_path):
    """Empty docx or docx with no tables should return an empty list."""
    empty_docx = tmp_path / "empty.docx"
    Document().save(empty_docx)

    detected = app.detect_docx_models(str(empty_docx))
    assert detected == []


def test_detect_docx_models_multi_row_formats(tmp_path):
    """Verify detection on both Cisco Name/P/N format and Huawei Part Number + Description summary format."""
    docx_path = tmp_path / "huawei_style.docx"
    doc = Document()
    # Summary table (Part Number -> Description)
    summary = doc.add_table(rows=2, cols=4)
    summary.rows[0].cells[0].text = "No"
    summary.rows[0].cells[1].text = "Part Number"
    summary.rows[0].cells[2].text = "Description"
    summary.rows[0].cells[3].text = "Serial Number"
    summary.rows[1].cells[1].text = "02355UEL"
    summary.rows[1].cells[2].text = "S6750-H48Y8C"
    summary.rows[1].cells[3].text = "SN0001"
    # Device table (5x2 key-value format)
    dev_tbl = doc.add_table(rows=5, cols=2)
    dev_tbl.rows[2].cells[0].text = "Part Number:"
    dev_tbl.rows[2].cells[1].text = "02355UEL"
    dev_tbl.rows[3].cells[0].text = "Serial Number:"
    dev_tbl.rows[3].cells[1].text = "SN0001"
    doc.save(docx_path)

    models = app.detect_docx_models(str(docx_path), sn_row_col=(3, 1))
    assert "S6750-H48Y8C (02355UEL)" in models


def test_detect_folder_models_hierarchy(tmp_path):
    """detect_folder_models should detect model directories containing photos under root."""
    root = tmp_path / "photos_root"
    for m, sn in [("7821", "WZP1001"), ("8841", "FVH2001"), ("CP-8865", "FCH3001")]:
        sn_dir = root / m / sn
        sn_dir.mkdir(parents=True)
        img = Image.new("RGB", (50, 50), color="white")
        img.save(sn_dir / "sample.jpg", "JPEG")
    (root / "ignore_dir").mkdir(parents=True)  # non-model dir without photos

    detected = app.detect_folder_models(str(root))
    assert isinstance(detected, list)
    assert "7821" in detected
    assert "8841" in detected
    assert "CP-8865" in detected
    assert "ignore_dir" not in detected


def test_detect_folder_models_empty_or_missing(tmp_path):
    """Empty or non-existent folder should return empty list."""
    empty_dir = tmp_path / "empty_dir"
    empty_dir.mkdir()
    assert app.detect_folder_models(str(empty_dir)) == []
    assert app.detect_folder_models(str(tmp_path / "non_existent")) == []


def test_model_matches_logic():
    """Verify flexible model matching rules (prefix stripping, suffix stripping, word boundary)."""
    assert app.model_matches("7821", "CP-7821-K9=") is True
    assert app.model_matches("CP-7821-K9=", "7821") is True
    assert app.model_matches("CP-7821", "CP-7821-K9=") is True
    assert app.model_matches("7821", "Cisco IP Phone 7821") is True
    assert app.model_matches("8841", "CP-8841-K9=") is True
    assert app.model_matches("8841", "CP-7821-K9=") is False
    assert app.model_matches("782", "CP-7821-K9=") is False
    assert app.model_matches("", "CP-7821-K9=") is False
    assert app.model_matches("CP-7821-K9=", "") is False


@pytest.fixture
def multi_model_env(tmp_path):
    docx_in = tmp_path / "multi_in.docx"
    docx_out = tmp_path / "multi_out.docx"
    photos_dir = tmp_path / "photos"

    specs = [
        {"sn": "WZP001", "model": "CP-7821-K9=", "placeholder": "UNTOUCHED_7821"},
        {"sn": "FVH002", "model": "CP-8841-K9=", "placeholder": "UNTOUCHED_8841"},
    ]
    create_mock_docx(docx_in, specs)
    create_mock_photo_root(photos_dir, ["WZP001", "FVH002"])

    return docx_in, docx_out, photos_dir


def test_run_injection_filter_single_model(multi_model_env):
    """Filtering by CP-7821-K9= must only inject table 1 and leave table 2 untouched."""
    docx_in, docx_out, photos_dir = multi_model_env
    res = app.run_injection(str(docx_in), str(photos_dir), str(docx_out),
                            keyword="",
                            selected_models=["CP-7821-K9="],
                            auto_rotate=False, smart_rotate=False)
    injected = res[0] if isinstance(res, tuple) else res.get('injected')
    assert injected == 1
    assert res.injected == 1

    out_doc = Document(docx_out)
    # Table 1 injected
    assert count_drawings_in_cell(out_doc.tables[0].rows[12].cells[0]) >= 1
    # Table 2 untouched
    t2_cell = out_doc.tables[1].rows[12].cells[0]
    assert t2_cell.text.strip() == "UNTOUCHED_8841"
    assert count_drawings_in_cell(t2_cell) == 0


def test_run_injection_filter_alternate_model(multi_model_env):
    """Filtering by CP-8841-K9= must only inject table 2 and leave table 1 untouched."""
    docx_in, docx_out, photos_dir = multi_model_env
    res = app.run_injection(str(docx_in), str(photos_dir), str(docx_out),
                            keyword="",
                            selected_models=["CP-8841-K9="],
                            auto_rotate=False, smart_rotate=False)
    injected = res[0] if isinstance(res, tuple) else res.get('injected')
    assert injected == 1
    assert res.injected == 1

    out_doc = Document(docx_out)
    # Table 1 untouched
    t1_cell = out_doc.tables[0].rows[12].cells[0]
    assert t1_cell.text.strip() == "UNTOUCHED_7821"
    assert count_drawings_in_cell(t1_cell) == 0
    # Table 2 injected
    assert count_drawings_in_cell(out_doc.tables[1].rows[12].cells[0]) >= 1


def test_run_injection_filter_none_or_empty_injects_all(multi_model_env):
    """selected_models=None or empty list should bypass filtering and inject all matching tables."""
    docx_in, docx_out, photos_dir = multi_model_env
    res = app.run_injection(str(docx_in), str(photos_dir), str(docx_out),
                            keyword="",
                            selected_models=None,
                            auto_rotate=False, smart_rotate=False)
    injected = res[0] if isinstance(res, tuple) else res.get('injected')
    assert injected == 2
    assert res.injected == 2

    out_doc = Document(docx_out)
    assert count_drawings_in_cell(out_doc.tables[0].rows[12].cells[0]) >= 1
    assert count_drawings_in_cell(out_doc.tables[1].rows[12].cells[0]) >= 1


def test_run_injection_filter_unmatched_model(multi_model_env):
    """Filtering by non-existent model leaves all tables untouched."""
    docx_in, docx_out, photos_dir = multi_model_env
    res = app.run_injection(str(docx_in), str(photos_dir), str(docx_out),
                            keyword="",
                            selected_models=["CP-9999-K9="],
                            auto_rotate=False, smart_rotate=False)
    injected = res[0] if isinstance(res, tuple) else res.get('injected')
    assert injected == 0
    assert res.injected == 0

    if docx_out.exists():
        out_doc = Document(docx_out)
        assert out_doc.tables[0].rows[12].cells[0].text.strip() == "UNTOUCHED_7821"
        assert out_doc.tables[1].rows[12].cells[0].text.strip() == "UNTOUCHED_8841"


def test_run_injection_combined_start_table_and_model_filter(tmp_path):
    """Verify combination of start_table and selected_models filtering."""
    docx_in = tmp_path / "comb_in.docx"
    docx_out = tmp_path / "comb_out.docx"
    photos_dir = tmp_path / "photos"

    specs = [
        {"sn": "SN001", "model": "CP-7821-K9=", "placeholder": "UNTOUCHED_1"},  # Table 1: Model A
        {"sn": "SN002", "model": "CP-7821-K9=", "placeholder": "UNTOUCHED_2"},  # Table 2: Model A
        {"sn": "SN003", "model": "CP-8841-K9=", "placeholder": "UNTOUCHED_3"},  # Table 3: Model B
        {"sn": "SN004", "model": "CP-7821-K9=", "placeholder": "UNTOUCHED_4"},  # Table 4: Model A
    ]
    create_mock_docx(docx_in, specs)
    create_mock_photo_root(photos_dir, ["SN001", "SN002", "SN003", "SN004"])

    # start_table=3 skips Tables 1 and 2.
    # selected_models=['CP-7821-K9='] skips Table 3 (which is CP-8841-K9=).
    # Only Table 4 should be injected.
    res = app.run_injection(str(docx_in), str(photos_dir), str(docx_out),
                            keyword="",
                            start_table=3,
                            selected_models=["CP-7821-K9="],
                            auto_rotate=False, smart_rotate=False)
    injected = res[0] if isinstance(res, tuple) else res.get('injected')
    assert injected == 1
    assert res.injected == 1

    out_doc = Document(docx_out)
    assert out_doc.tables[0].rows[12].cells[0].text.strip() == "UNTOUCHED_1"
    assert out_doc.tables[1].rows[12].cells[0].text.strip() == "UNTOUCHED_2"
    assert out_doc.tables[2].rows[12].cells[0].text.strip() == "UNTOUCHED_3"
    assert count_drawings_in_cell(out_doc.tables[3].rows[12].cells[0]) >= 1


def test_model_matches_whitespace_does_not_match():
    """Verify whitespace-only strings do not match as equal models."""
    assert app.model_matches("   ", "   ") is False
    assert app.model_matches("\t", "   ") is False
    assert app.model_matches("   ", "") is False
    assert app.model_matches("", "   ") is False
    assert app.model_matches("  \n  ", "\t") is False


def test_detect_docx_models_does_not_extract_general_text(tmp_path):
    """Verify non-model text in table cells (e.g. cell (5, 2)) is not detected as a model."""
    docx_path = tmp_path / "general_text.docx"
    doc = Document()
    t = doc.add_table(rows=6, cols=4)
    t.rows[5].cells[2].text = "General Description"
    doc.save(docx_path)

    models = app.detect_docx_models(str(docx_path))
    assert models == []


def test_detect_docx_models_preserves_trailing_equal_sign_without_duplicates(tmp_path):
    """Verify document containing Name : CP-7821-K9= and standalone CP-7821-K9= returns exactly ['CP-7821-K9=']."""
    docx_path = tmp_path / "equal_sign_dup.docx"
    doc = Document()
    t1 = doc.add_table(rows=6, cols=4)
    t1.rows[2].cells[0].text = "Name : CP-7821-K9="
    t2 = doc.add_table(rows=6, cols=4)
    t2.rows[2].cells[0].text = "CP-7821-K9="
    doc.save(docx_path)

    models = app.detect_docx_models(str(docx_path))
    assert models == ["CP-7821-K9="]

