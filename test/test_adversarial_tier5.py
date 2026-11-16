"""
Adversarial Hardening Test Suite (Tier 5) for Milestone 3.
Verifies all 8 vulnerability and defect remediations identified by Challengers 1 & 2:
1. Guard start_table against OverflowError (e.g. float('inf'), float('-inf'))
2. Coerce selected_models items to string safely (prevent AttributeError on int/enum)
3. Fix false-positive model detection in detect_folder_models for camera subdirectories
4. Update MODEL_LABEL_REGEX to allow spaces and strip trailing whitespace
5. Capture all device models per table in detect_docx_models
6. Reset start_table_var to 1 when docx has 0 tables in _update_docx_info
7. Correct start table clamping logic before mutating _total_tables in _update_docx_info
8. Coerce initial_table_idx safely in TableSelectionDialog
"""
import os
import sys
from pathlib import Path
import pytest
from docx import Document
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import app


@pytest.fixture(scope="session")
def headless_root():
    """Session-scoped root App instance for UI component tests."""
    root = app.App()
    root.withdraw()
    yield root
    try:
        root.destroy()
    except Exception:
        pass


def _create_simple_doc(path: Path, table_count: int = 1, sn_list: list[str] = None,
                       model: str = "CP-7821-K9=") -> Path:
    """Helper to generate a clean docx with table_count tables."""
    doc = Document()
    if sn_list is None:
        sn_list = [f"WZP{i+1:04d}" for i in range(table_count)]

    for i in range(table_count):
        tbl = doc.add_table(rows=14, cols=4)
        tbl.rows[2].cells[0].text = f"Name : {model} P/N : {model}"
        tbl.rows[5].cells[2].text = model
        tbl.rows[5].cells[3].text = sn_list[i] if i < len(sn_list) else f"WZP{i+1:04d}"
        tbl.rows[12].cells[0].text = "IMAGE_PLACEHOLDER"

    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(path))
    return path


def _create_photo_dir(path: Path, sn_list: list[str], camera_subdir: str = None) -> Path:
    """Helper to generate photo folders for S/Ns, optionally with camera subdirectories."""
    for sn in sn_list:
        folder = path / sn
        if camera_subdir:
            folder = folder / camera_subdir
        folder.mkdir(parents=True, exist_ok=True)
        img = Image.new("RGB", (30, 30), color="blue")
        img.save(folder / "IMG_0001.JPG", "JPEG")
    return path


# ─────────────────────────────────────────────────────────────────────────────
# Task 1: Guard start_table against OverflowError
# ─────────────────────────────────────────────────────────────────────────────

def test_task1_start_table_overflow_error_guard(tmp_path):
    """Calling run_injection with float('inf') or float('-inf') must not raise OverflowError."""
    docx_file = _create_simple_doc(tmp_path / "test.docx", table_count=2, sn_list=["WZP0001", "WZP0002"])
    photo_dir = _create_photo_dir(tmp_path / "photos", ["WZP0001", "WZP0002"])
    out_file = tmp_path / "out.docx"

    # float('inf') should safely default to 1 and process all tables
    res_inf = app.run_injection(str(docx_file), str(photo_dir), str(out_file), start_table=float("inf"))
    assert res_inf.injected == 2
    assert res_inf.no_folder == 0

    # float('-inf') should safely default to 1 and process all tables
    res_neginf = app.run_injection(str(docx_file), str(photo_dir), str(out_file), start_table=float("-inf"))
    assert res_neginf.injected == 2

    # float('nan') should safely default to 1
    res_nan = app.run_injection(str(docx_file), str(photo_dir), str(out_file), start_table=float("nan"))
    assert res_nan.injected == 2


# ─────────────────────────────────────────────────────────────────────────────
# Task 2: Coerce selected_models items to string safely
# ─────────────────────────────────────────────────────────────────────────────

def test_task2_selected_models_safe_string_coercion(tmp_path):
    """Passing integer or mixed non-string types in selected_models must not raise AttributeError."""
    docx_file = _create_simple_doc(tmp_path / "test.docx", table_count=2, sn_list=["WZP0001", "FVH0002"],
                                   model="CP-7821-K9=")
    photo_dir = _create_photo_dir(tmp_path / "photos", ["WZP0001", "FVH0002"])
    out_file = tmp_path / "out.docx"

    # Integer model identifier: [7821] instead of ["7821"]
    res = app.run_injection(str(docx_file), str(photo_dir), str(out_file), selected_models=[7821])
    assert res.injected == 2

    # Mixed list with integers, None, whitespace strings, valid strings
    res_mixed = app.run_injection(str(docx_file), str(photo_dir), str(out_file),
                                  selected_models=[7821, None, "", "   ", 8841])
    assert res_mixed.injected == 2


# ─────────────────────────────────────────────────────────────────────────────
# Task 3: Fix false-positive model detection for camera subdirectories
# ─────────────────────────────────────────────────────────────────────────────

def test_task3_folder_models_camera_subdirectories_ignored(tmp_path):
    """S/N folders with camera subfolders (100APPLE, DCIM, CAMERA) must NOT be detected as models."""
    sn_root = tmp_path / "sn_photos"

    # Flat S/N folders containing camera subdirectories
    _create_photo_dir(sn_root, ["WZP0001"], camera_subdir="100APPLE")
    _create_photo_dir(sn_root, ["WZP0002"], camera_subdir="DCIM")
    _create_photo_dir(sn_root, ["WZP0003"], camera_subdir="CAMERA")
    _create_photo_dir(sn_root, ["Raw_Archive"], camera_subdir="100MEDIA")

    detected = app.detect_folder_models(str(sn_root))
    # Serial numbers and archive folders must NOT be detected as device models
    assert detected == []


def test_task3_folder_models_valid_hierarchy_still_detected(tmp_path):
    """Genuine model folders (e.g. 7821, CP-8841-K9=) must still be detected even if S/N has camera subdirs."""
    photos_root = tmp_path / "model_photos"

    # Nested hierarchy: model/SN/camera_subdir/images
    m7821_dir = photos_root / "7821"
    _create_photo_dir(m7821_dir, ["WZP0001"], camera_subdir="100APPLE")

    m8841_dir = photos_root / "CP-8841-K9="
    _create_photo_dir(m8841_dir, ["FVH0001"], camera_subdir="DCIM")

    detected = app.detect_folder_models(str(photos_root))
    assert "7821" in detected
    assert "CP-8841-K9=" in detected
    assert "WZP0001" not in detected
    assert "FVH0001" not in detected


# ─────────────────────────────────────────────────────────────────────────────
# Task 4: Allow spaces in MODEL_LABEL_REGEX and strip whitespace
# ─────────────────────────────────────────────────────────────────────────────

def test_task4_model_label_regex_handles_spaces_and_strips():
    """MODEL_LABEL_REGEX matches spaced model names and does not truncate at spaces."""
    # Test spaced model names
    m = app.MODEL_LABEL_REGEX.search("Model : Cisco IP Phone 7821")
    assert m is not None
    assert m.group(1).strip() == "Cisco IP Phone 7821"

    # Test that adjacent fields (e.g. P/N :) are not greedily absorbed
    m2 = app.MODEL_LABEL_REGEX.search("Name : CP-7821-K9= P/N : CP-7821-K9=")
    assert m2 is not None
    assert m2.group(1).strip() == "CP-7821-K9="

    # Test trailing whitespace is cleanly stripped
    m3 = app.MODEL_LABEL_REGEX.search("Model : Cisco IP Phone 8841   ")
    assert m3 is not None
    assert m3.group(1).strip() == "Cisco IP Phone 8841"

    # Test model_matches with spaced target
    assert app.model_matches("7821", "Cisco IP Phone 7821") is True
    assert app.model_matches("CP-7821", "Cisco IP Phone 7821") is True


# ─────────────────────────────────────────────────────────────────────────────
# Task 5: Multi-model per table extraction in detect_docx_models
# ─────────────────────────────────────────────────────────────────────────────

def test_task5_detect_docx_models_multi_model_per_table(tmp_path):
    """When a table contains multiple models (e.g. main device + secondary), both must be detected."""
    docx_path = tmp_path / "multi_model_table.docx"
    doc = Document()
    tbl = doc.add_table(rows=14, cols=4)

    # Row 2 contains main model
    tbl.rows[2].cells[0].text = "Name : CP-7821-K9="
    # Row 5 detail mentions accessory or secondary model
    tbl.rows[5].cells[1].text = "Expansion Module CP-8841-K9="
    tbl.rows[5].cells[2].text = "CP-7821-K9="
    tbl.rows[5].cells[3].text = "WZP0001"
    tbl.rows[12].cells[0].text = "IMAGE_HERE"
    doc.save(str(docx_path))

    detected = app.detect_docx_models(str(docx_path))
    assert "CP-7821-K9=" in detected
    assert "CP-8841-K9=" in detected

    # Verify injection matching on secondary model
    photo_dir = _create_photo_dir(tmp_path / "photos", ["WZP0001"])
    out_path = tmp_path / "multi_out.docx"
    res = app.run_injection(str(docx_path), str(photo_dir), str(out_path),
                            selected_models=["CP-8841-K9="])
    assert res.injected == 1


# ─────────────────────────────────────────────────────────────────────────────
# Task 6: Reset start_table_var on 0-table document in _update_docx_info
# ─────────────────────────────────────────────────────────────────────────────

def test_task6_start_table_reset_on_zero_table_document(tmp_path, headless_root):
    """When loading a document with 0 tables, start_table_var must reset cleanly to '1'."""
    zero_doc = tmp_path / "zero_tables.docx"
    Document().save(str(zero_doc))

    # Pre-set start table to an arbitrary value
    headless_root.start_table_var.set("8")
    headless_root._update_docx_info(str(zero_doc))

    assert headless_root.start_table_var.get() == "1"
    if hasattr(headless_root, "total_tables_label"):
        assert headless_root.total_tables_label.cget("text") == " (0 tables)"


# ─────────────────────────────────────────────────────────────────────────────
# Task 7: Proper clamping logic in _update_docx_info
# ─────────────────────────────────────────────────────────────────────────────

def test_task7_start_table_clamping_logic(tmp_path, headless_root):
    """When loading a document with fewer tables than start_table_var, it clamps to total."""
    doc_3tables = _create_simple_doc(tmp_path / "doc3.docx", table_count=3)

    # Set start table higher than document table count
    headless_root.start_table_var.set("10")
    headless_root._update_docx_info(str(doc_3tables))

    assert headless_root.start_table_var.get() == "3"
    if hasattr(headless_root, "total_tables_label"):
        assert headless_root.total_tables_label.cget("text") == " (of 3)"


# ─────────────────────────────────────────────────────────────────────────────
# Task 8: TableSelectionDialog initial_table_idx coercion
# ─────────────────────────────────────────────────────────────────────────────

def test_task8_dialog_initial_table_idx_coercion(tmp_path, headless_root):
    """TableSelectionDialog must gracefully handle non-integer initial_table_idx values."""
    docx_file = _create_simple_doc(tmp_path / "dialog_test.docx", table_count=3)

    # initial_table_idx as string "1"
    dlg_str = app.TableSelectionDialog(headless_root, str(docx_file), (5, 3), (12, 0), initial_table_idx="1")
    assert dlg_str.current_table_idx == 1
    dlg_str.destroy()

    # initial_table_idx as None -> defaults to 0
    dlg_none = app.TableSelectionDialog(headless_root, str(docx_file), (5, 3), (12, 0), initial_table_idx=None)
    assert dlg_none.current_table_idx == 0
    dlg_none.destroy()

    # initial_table_idx as invalid string -> defaults to 0
    dlg_invalid = app.TableSelectionDialog(headless_root, str(docx_file), (5, 3), (12, 0), initial_table_idx="abc")
    assert dlg_invalid.current_table_idx == 0
    dlg_invalid.destroy()

    # initial_table_idx as float('inf') -> defaults to 0 or clamped to max
    dlg_inf = app.TableSelectionDialog(headless_root, str(docx_file), (5, 3), (12, 0), initial_table_idx=float("inf"))
    assert dlg_inf.current_table_idx in (0, 2)
    dlg_inf.destroy()
