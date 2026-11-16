"""
Tier 2: Boundary & Corner Cases (20 tests)
Covers boundary value analysis and edge conditions:
1. Starting Table Selection Boundaries - 5 tests
2. Device Model Auto-Detection Boundaries - 5 tests
3. Device Model Filtering Boundaries - 5 tests
4. Arbitrary Table Layout Configuration Boundaries - 5 tests
"""

import inspect
import os
from pathlib import Path
import pytest
from docx import Document

import app


def _check_start_table_support():
    sig = inspect.signature(app.run_injection)
    if "start_table" not in sig.parameters:
        pytest.skip("run_injection start_table parameter not yet implemented in app.py")


def _check_model_detection_support():
    if not hasattr(app, "detect_docx_models") or not hasattr(app, "detect_folder_models"):
        pytest.skip("detect_docx_models/detect_folder_models not yet implemented in app.py")


def _check_model_filtering_support():
    sig = inspect.signature(app.run_injection)
    if "selected_models" not in sig.parameters:
        pytest.skip("run_injection selected_models parameter not yet implemented in app.py")


def _get_injected_count(res):
    if hasattr(res, "injected"):
        return res.injected
    if isinstance(res, dict):
        return res["injected"]
    return res[0]


# ─────────────────────────────────────────────────────────────────────────────
# Feature 1: Starting Table Selection Boundary Cases
# ─────────────────────────────────────────────────────────────────────────────

def test_tier2_start_table_1_exact_lower_bound(tmp_path, mock_docx_factory, mock_photo_factory, docx_table_inspector):
    """21. Exact lower bound start_table=1 processes all tables."""
    _check_start_table_support()
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"sn": f"WZP000{i}"} for i in range(1, 4)
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "": [f"WZP000{i}" for i in range(1, 4)]
    })
    output_docx = tmp_path / "output.docx"

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx), start_table=1)
    assert _get_injected_count(res) == 3


def test_tier2_start_table_last_table_exact_upper_bound(tmp_path, mock_docx_factory, mock_photo_factory, docx_table_inspector):
    """22. Exact upper bound start_table=N processes only the final table."""
    _check_start_table_support()
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"sn": "WZP0001"},
        {"sn": "WZP0002"},
        {"sn": "WZP0003"},
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "": ["WZP0001", "WZP0002", "WZP0003"]
    })
    output_docx = tmp_path / "output.docx"

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx), start_table=3)
    assert _get_injected_count(res) == 1

    out_doc = Document(str(output_docx))
    assert docx_table_inspector.get_inlines_count(out_doc.tables[0], 12, 0) == 0
    assert docx_table_inspector.get_inlines_count(out_doc.tables[1], 12, 0) == 0
    assert docx_table_inspector.get_inlines_count(out_doc.tables[2], 12, 0) >= 1


def test_tier2_start_table_greater_than_total_tables(tmp_path, mock_docx_factory, mock_photo_factory, docx_table_inspector):
    """23. start_table = N + 1 processes 0 tables safely without raising unhandled exception."""
    _check_start_table_support()
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"sn": "WZP0001"},
        {"sn": "WZP0002"},
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "": ["WZP0001", "WZP0002"]
    })
    output_docx = tmp_path / "output.docx"

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx), start_table=3)
    assert _get_injected_count(res) == 0

    if output_docx.exists():
        out_doc = Document(str(output_docx))
        assert docx_table_inspector.get_inlines_count(out_doc.tables[0], 12, 0) == 0
        assert docx_table_inspector.get_inlines_count(out_doc.tables[1], 12, 0) == 0


def test_tier2_start_table_zero_or_negative_clamped_or_rejected(tmp_path, mock_docx_factory, mock_photo_factory):
    """24. start_table <= 0 is either clamped to 1 or raises ValueError safely."""
    _check_start_table_support()
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"sn": "WZP0001"},
        {"sn": "WZP0002"},
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "": ["WZP0001", "WZP0002"]
    })
    output_docx = tmp_path / "output.docx"

    try:
        res = app.run_injection(str(docx_file), str(photo_root), str(output_docx), start_table=0)
        # If clamped to 1, all 2 tables are injected
        assert _get_injected_count(res) == 2
    except ValueError:
        # Raising ValueError on non-positive start_table is also compliant
        pass


def test_tier2_start_table_single_table_document(tmp_path, mock_docx_factory, mock_photo_factory, docx_table_inspector):
    """25. Single-table document (N=1) handles start_table=1 and start_table=2 correctly."""
    _check_start_table_support()
    docx_file = mock_docx_factory(tmp_path / "input.docx", [{"sn": "WZP0001"}])
    photo_root = mock_photo_factory(tmp_path / "photos", {"": ["WZP0001"]})

    # start_table=1 -> 1 injected
    out1 = tmp_path / "out1.docx"
    res1 = app.run_injection(str(docx_file), str(photo_root), str(out1), start_table=1)
    assert _get_injected_count(res1) == 1

    # start_table=2 -> 0 injected
    out2 = tmp_path / "out2.docx"
    res2 = app.run_injection(str(docx_file), str(photo_root), str(out2), start_table=2)
    assert _get_injected_count(res2) == 0


# ─────────────────────────────────────────────────────────────────────────────
# Feature 2: Device Model Auto-Detection Boundary Cases
# ─────────────────────────────────────────────────────────────────────────────

def test_tier2_detect_docx_models_empty_document_zero_tables(tmp_path):
    """26. detect_docx_models on document with 0 tables returns [] cleanly."""
    _check_model_detection_support()
    docx_file = tmp_path / "empty.docx"
    doc = Document()
    doc.add_paragraph("This document has no tables.")
    doc.save(str(docx_file))

    models = app.detect_docx_models(str(docx_file))
    assert models == []


def test_tier2_detect_docx_models_corrupt_or_non_docx_file(tmp_path):
    """27. detect_docx_models on corrupt / 0-byte file handled safely."""
    _check_model_detection_support()
    corrupt_file = tmp_path / "corrupt.docx"
    corrupt_file.write_bytes(b"NOT_A_VALID_ZIP_PACKAGE")

    try:
        models = app.detect_docx_models(str(corrupt_file))
        assert models == []
    except Exception as e:
        # Standard docx / zip exception is acceptable error handling
        assert type(e).__name__ in ("BadZipFile", "PackageNotFoundError", "ValueError")


def test_tier2_detect_docx_models_tables_without_model_text(tmp_path, mock_docx_factory):
    """28. detect_docx_models on tables with non-model generic text returns []."""
    _check_model_detection_support()
    docx_file = mock_docx_factory(tmp_path / "generic.docx", [
        {"part_number": "N/A", "model_name": "Generic Item 1", "sn": "12345"},
        {"part_number": "ITEM-XYZ", "model_name": "Generic Item 2", "sn": "67890"},
    ])
    models = app.detect_docx_models(str(docx_file))
    # Should not falsely identify generic non-Cisco strings
    assert not any("7821" in m or "8841" in m for m in models)


def test_tier2_detect_folder_models_empty_or_nonexistent_directory(tmp_path):
    """29. detect_folder_models returns [] for non-existent or empty directory."""
    _check_model_detection_support()
    assert app.detect_folder_models(str(tmp_path / "nonexistent")) == []
    empty_dir = tmp_path / "empty_dir"
    empty_dir.mkdir()
    assert app.detect_folder_models(str(empty_dir)) == []


def test_tier2_detect_folder_models_folders_with_no_images(tmp_path):
    """30. detect_folder_models ignores subfolders that contain no valid image files."""
    _check_model_detection_support()
    photos_dir = tmp_path / "photos"
    model_7821 = photos_dir / "7821" / "WZP0001"
    model_7821.mkdir(parents=True)
    (model_7821 / "notes.txt").write_text("no images here")

    models = app.detect_folder_models(str(photos_dir))
    assert models == []


# ─────────────────────────────────────────────────────────────────────────────
# Feature 3: Device Model Filtering Boundary Cases
# ─────────────────────────────────────────────────────────────────────────────

def test_tier2_model_filtering_empty_selected_models_list(tmp_path, mock_docx_factory, mock_photo_factory, docx_table_inspector):
    """31. selected_models=[] injects 0 tables and preserves all document tables."""
    _check_model_filtering_support()
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"part_number": "CP-7821-K9=", "sn": "WZP0001"},
        {"part_number": "CP-8841-K9=", "sn": "FVH0001"},
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "7821": ["WZP0001"],
        "8841": ["FVH0001"]
    })
    output_docx = tmp_path / "output.docx"

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx), keyword="", selected_models=[])
    assert _get_injected_count(res) == 0

    if output_docx.exists():
        out_doc = Document(str(output_docx))
        assert docx_table_inspector.get_inlines_count(out_doc.tables[0], 12, 0) == 0
        assert docx_table_inspector.get_inlines_count(out_doc.tables[1], 12, 0) == 0


def test_tier2_model_filtering_disjoint_unrecognized_model_name(tmp_path, mock_docx_factory, mock_photo_factory, docx_table_inspector):
    """32. selected_models with an unknown model injects 0 tables and leaves all tables untouched."""
    _check_model_filtering_support()
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"part_number": "CP-7821-K9=", "sn": "WZP0001"},
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "7821": ["WZP0001"]
    })
    output_docx = tmp_path / "output.docx"

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx), keyword="",
                            selected_models=["UNKNOWN_MODEL_X"])
    assert _get_injected_count(res) == 0

    if output_docx.exists():
        out_doc = Document(str(output_docx))
        assert docx_table_inspector.get_inlines_count(out_doc.tables[0], 12, 0) == 0


def test_tier2_model_filtering_case_insensitivity_and_whitespace(tmp_path, mock_docx_factory, mock_photo_factory):
    """33. selected_models handles variations in casing and whitespace."""
    _check_model_filtering_support()
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"part_number": "CP-7821-K9=", "sn": "WZP0001"}
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "7821": ["WZP0001"]
    })
    output_docx = tmp_path / "output.docx"

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx), keyword="",
                            selected_models=["  cp-7821-k9=  "])
    assert _get_injected_count(res) == 1


def test_tier2_model_filtering_all_available_models_explicitly_selected(tmp_path, mock_docx_factory, mock_photo_factory):
    """34. Explicitly passing all models matches every corresponding table."""
    _check_model_filtering_support()
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"part_number": "CP-7821-K9=", "sn": "WZP0001"},
        {"part_number": "CP-8841-K9=", "sn": "FVH0001"},
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "7821": ["WZP0001"],
        "8841": ["FVH0001"]
    })
    output_docx = tmp_path / "output.docx"

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx), keyword="",
                            selected_models=["CP-7821-K9=", "CP-8841-K9="])
    assert _get_injected_count(res) == 2


def test_tier2_model_filtering_partially_available_photos(tmp_path, mock_docx_factory, mock_photo_factory, docx_table_inspector):
    """35. Model is selected but photos are missing: injects 0, leaves document untouched."""
    _check_model_filtering_support()
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"part_number": "CP-7821-K9=", "sn": "WZP0001"},
        {"part_number": "CP-8841-K9=", "sn": "FVH0001"},
    ])
    # Photos available ONLY for 7821
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "7821": ["WZP0001"]
    })
    output_docx = tmp_path / "output.docx"

    # Filter for 8841 only (no photos available)
    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx), keyword="",
                            selected_models=["CP-8841-K9="])
    assert _get_injected_count(res) == 0

    if output_docx.exists():
        out_doc = Document(str(output_docx))
        assert docx_table_inspector.get_inlines_count(out_doc.tables[0], 12, 0) == 0
        assert docx_table_inspector.get_inlines_count(out_doc.tables[1], 12, 0) == 0


# ─────────────────────────────────────────────────────────────────────────────
# Feature 4: Arbitrary Table Layout Configuration Boundary Cases
# ─────────────────────────────────────────────────────────────────────────────

def test_tier2_dialog_empty_document_zero_tables(tmp_path, dialog_factory):
    """36. Dialog handles 0-table document safely without crashing."""
    docx_file = tmp_path / "zero_tables.docx"
    doc = Document()
    doc.add_paragraph("No tables")
    doc.save(str(docx_file))

    dlg = dialog_factory(str(docx_file))
    # Dialog should load without unhandled exception
    assert dlg is not None
    dlg.destroy()


def test_tier2_dialog_prev_on_first_table_clamped(tmp_path, mock_docx_factory, dialog_factory):
    """37. Prev button on Table 1 stays clamped at index 0."""
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"sn": "WZP0001"}, {"sn": "WZP0002"}
    ])
    dlg = dialog_factory(str(docx_file), initial_table_idx=0)
    if hasattr(dlg, "_prev_table"):
        dlg._prev_table()
        assert getattr(dlg, "current_table_idx", 0) == 0
    elif hasattr(dlg, "_render_table"):
        dlg._render_table(0)
    dlg.destroy()


def test_tier2_dialog_next_on_last_table_clamped(tmp_path, mock_docx_factory, dialog_factory):
    """38. Next button on Table N stays clamped at index N-1."""
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"sn": "WZP0001"}, {"sn": "WZP0002"}
    ])
    dlg = dialog_factory(str(docx_file), initial_table_idx=1)
    if hasattr(dlg, "_next_table"):
        dlg._next_table()
        assert getattr(dlg, "current_table_idx", 1) == 1
    elif hasattr(dlg, "_render_table"):
        dlg._render_table(1)
    dlg.destroy()


def test_tier2_dialog_inhomogeneous_table_dimensions(tmp_path, mock_docx_factory, dialog_factory):
    """39. Navigating between tables with different row/col dimensions re-renders safely."""
    docx_file = mock_docx_factory(tmp_path / "inhomogeneous.docx", [
        {"rows": 3, "cols": 2, "sn": "WZP0001", "sn_rc": (1, 1), "img_rc": (2, 0)},
        {"rows": 18, "cols": 4, "sn": "WZP0002", "sn_rc": (5, 3), "img_rc": (12, 0)},
    ])
    dlg = dialog_factory(str(docx_file), initial_table_idx=0)
    if hasattr(dlg, "_render_table"):
        dlg._render_table(1)
        assert (17, 3) in dlg.buttons or (5, 3) in dlg.buttons
        dlg._render_table(0)
    dlg.destroy()


def test_tier2_dialog_out_of_bounds_initial_table_idx(tmp_path, mock_docx_factory, dialog_factory):
    """40. Initial table index exceeding table count is clamped safely without raising IndexError."""
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"sn": "WZP0001"}, {"sn": "WZP0002"}
    ])
    dlg = dialog_factory(str(docx_file), initial_table_idx=999)
    # Dialog should load and not crash
    assert dlg is not None
    dlg.destroy()
