"""
Tier 3: Cross-Feature Combinations (11 tests)
Validates pairwise interactions between features:
- start_table + selected_models
- detect_models + selected_models
- TableSelectionDialog coordinates + start_table / selected_models
- cancel_event during filtered start_table injection
"""

import inspect
import threading
from pathlib import Path
import pytest
from docx import Document

import app


def _check_m1_support():
    sig = inspect.signature(app.run_injection)
    if "start_table" not in sig.parameters or "selected_models" not in sig.parameters:
        pytest.skip("run_injection start_table / selected_models parameters not yet implemented in app.py")


def _get_injected_count(res):
    if hasattr(res, "injected"):
        return res.injected
    if isinstance(res, dict):
        return res["injected"]
    return res[0]


def test_tier3_start_table_with_single_model_filter(tmp_path, mock_docx_factory, mock_photo_factory, docx_table_inspector):
    """41. start_table = 3 with selected_models=['CP-7821-K9='] on alternating document."""
    _check_m1_support()
    # 6 tables alternating: T1=7821, T2=8841, T3=7821, T4=8841, T5=7821, T6=8841
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"part_number": "CP-7821-K9=", "sn": "WZP0001"},
        {"part_number": "CP-8841-K9=", "sn": "FVH0001"},
        {"part_number": "CP-7821-K9=", "sn": "WZP0002"},
        {"part_number": "CP-8841-K9=", "sn": "FVH0002"},
        {"part_number": "CP-7821-K9=", "sn": "WZP0003"},
        {"part_number": "CP-8841-K9=", "sn": "FVH0003"},
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "7821": ["WZP0001", "WZP0002", "WZP0003"],
        "8841": ["FVH0001", "FVH0002", "FVH0003"]
    })
    output_docx = tmp_path / "output.docx"

    # Start at table 3 (WZP0002) and filter for 7821 only -> should inject T3 and T5 (2 tables)
    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx),
                            keyword="", start_table=3, selected_models=["CP-7821-K9="])
    assert _get_injected_count(res) == 2

    out_doc = Document(str(output_docx))
    assert docx_table_inspector.get_inlines_count(out_doc.tables[0], 12, 0) == 0  # T1: <start_table
    assert docx_table_inspector.get_inlines_count(out_doc.tables[1], 12, 0) == 0  # T2: <start_table
    assert docx_table_inspector.get_inlines_count(out_doc.tables[2], 12, 0) >= 1  # T3: injected
    assert docx_table_inspector.get_inlines_count(out_doc.tables[3], 12, 0) == 0  # T4: wrong model
    assert docx_table_inspector.get_inlines_count(out_doc.tables[4], 12, 0) >= 1  # T5: injected
    assert docx_table_inspector.get_inlines_count(out_doc.tables[5], 12, 0) == 0  # T6: wrong model


def test_tier3_start_table_with_multi_model_filter(tmp_path, mock_docx_factory, mock_photo_factory, docx_table_inspector):
    """42. start_table = 4 with both models selected injects tables 4, 5, 6."""
    _check_m1_support()
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"part_number": "CP-7821-K9=", "sn": "WZP0001"},
        {"part_number": "CP-8841-K9=", "sn": "FVH0001"},
        {"part_number": "CP-7821-K9=", "sn": "WZP0002"},
        {"part_number": "CP-8841-K9=", "sn": "FVH0002"},
        {"part_number": "CP-7821-K9=", "sn": "WZP0003"},
        {"part_number": "CP-8841-K9=", "sn": "FVH0003"},
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "7821": ["WZP0001", "WZP0002", "WZP0003"],
        "8841": ["FVH0001", "FVH0002", "FVH0003"]
    })
    output_docx = tmp_path / "output.docx"

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx),
                            keyword="", start_table=4, selected_models=["CP-7821-K9=", "CP-8841-K9="])
    assert _get_injected_count(res) == 3

    out_doc = Document(str(output_docx))
    assert docx_table_inspector.get_inlines_count(out_doc.tables[0], 12, 0) == 0
    assert docx_table_inspector.get_inlines_count(out_doc.tables[1], 12, 0) == 0
    assert docx_table_inspector.get_inlines_count(out_doc.tables[2], 12, 0) == 0
    assert docx_table_inspector.get_inlines_count(out_doc.tables[3], 12, 0) >= 1
    assert docx_table_inspector.get_inlines_count(out_doc.tables[4], 12, 0) >= 1
    assert docx_table_inspector.get_inlines_count(out_doc.tables[5], 12, 0) >= 1


def test_tier3_start_table_beyond_selected_model_tables(tmp_path, mock_docx_factory, mock_photo_factory, docx_table_inspector):
    """43. start_table is past all tables matching the selected model -> 0 injected."""
    _check_m1_support()
    # Tables 1-3 are 7821; Tables 4-6 are 8841
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"part_number": "CP-7821-K9=", "sn": "WZP0001"},
        {"part_number": "CP-7821-K9=", "sn": "WZP0002"},
        {"part_number": "CP-7821-K9=", "sn": "WZP0003"},
        {"part_number": "CP-8841-K9=", "sn": "FVH0001"},
        {"part_number": "CP-8841-K9=", "sn": "FVH0002"},
        {"part_number": "CP-8841-K9=", "sn": "FVH0003"},
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "7821": ["WZP0001", "WZP0002", "WZP0003"],
        "8841": ["FVH0001", "FVH0002", "FVH0003"]
    })
    output_docx = tmp_path / "output.docx"

    # Start at table 4, but filter for 7821 -> all 7821 tables are before table 4
    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx),
                            keyword="", start_table=4, selected_models=["CP-7821-K9="])
    assert _get_injected_count(res) == 0


def test_tier3_auto_detect_docx_models_then_filter_injection(tmp_path, mock_docx_factory, mock_photo_factory, docx_table_inspector):
    """44. Detect models from docx dynamically, select the first detected model, and inject."""
    if not hasattr(app, "detect_docx_models"):
        pytest.skip("detect_docx_models not yet implemented in app.py")
    _check_m1_support()

    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"part_number": "CP-7821-K9=", "sn": "WZP0001"},
        {"part_number": "CP-8841-K9=", "sn": "FVH0001"},
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "7821": ["WZP0001"],
        "8841": ["FVH0001"]
    })
    output_docx = tmp_path / "output.docx"

    detected_models = app.detect_docx_models(str(docx_file))
    assert len(detected_models) >= 1
    selected = [detected_models[0]]

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx),
                            keyword="", selected_models=selected)
    assert _get_injected_count(res) == 1


def test_tier3_auto_detect_folder_models_then_filter_injection(tmp_path, mock_docx_factory, mock_photo_factory):
    """45. Detect models from folders dynamically, filter docx injection by detected model."""
    if not hasattr(app, "detect_folder_models"):
        pytest.skip("detect_folder_models not yet implemented in app.py")
    _check_m1_support()

    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"part_number": "CP-7821-K9=", "sn": "WZP0001"},
        {"part_number": "CP-8841-K9=", "sn": "FVH0001"},
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "CP-7821-K9=": ["WZP0001"],
        "CP-8841-K9=": ["FVH0001"]
    })
    output_docx = tmp_path / "output.docx"

    detected_models = app.detect_folder_models(str(photo_root))
    assert len(detected_models) >= 1
    # Filter by the detected 8841 model
    target = [m for m in detected_models if "8841" in m]
    assert len(target) > 0

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx),
                            keyword="", selected_models=target)
    assert _get_injected_count(res) == 1


def test_tier3_dialog_custom_coords_with_start_table(tmp_path, mock_docx_factory, mock_photo_factory, dialog_factory, docx_table_inspector):
    """46. Coordinates selected in dialog combined with start_table=2."""
    _check_m1_support()
    # Non-standard table layout: SN at (2, 1), Image at (6, 0)
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"rows": 10, "cols": 4, "sn_rc": (2, 1), "img_rc": (6, 0), "sn": "WZP0001"},
        {"rows": 10, "cols": 4, "sn_rc": (2, 1), "img_rc": (6, 0), "sn": "WZP0002"},
        {"rows": 10, "cols": 4, "sn_rc": (2, 1), "img_rc": (6, 0), "sn": "WZP0003"},
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "": ["WZP0001", "WZP0002", "WZP0003"]
    })
    output_docx = tmp_path / "output.docx"

    # Simulate user setting coordinates via dialog
    dlg = dialog_factory(str(docx_file))
    dlg.sn_rc = (2, 1)
    dlg.img_rc = (6, 0)
    dlg._approve()
    sn_coord, img_coord = dlg.result

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx),
                            sn_row_col=sn_coord, img_row_col=img_coord, start_table=2)
    assert _get_injected_count(res) == 2

    out_doc = Document(str(output_docx))
    assert docx_table_inspector.get_inlines_count(out_doc.tables[0], 6, 0) == 0
    assert docx_table_inspector.get_inlines_count(out_doc.tables[1], 6, 0) >= 1
    assert docx_table_inspector.get_inlines_count(out_doc.tables[2], 6, 0) >= 1


def test_tier3_dialog_navigation_on_multi_model_document_then_filter(tmp_path, mock_docx_factory, mock_photo_factory, dialog_factory, docx_table_inspector):
    """47. Dialog approval followed by model filtered injection."""
    _check_m1_support()
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"part_number": "CP-7821-K9=", "sn": "WZP0001"},
        {"part_number": "CP-8841-K9=", "sn": "FVH0001"},
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "7821": ["WZP0001"],
        "8841": ["FVH0001"]
    })
    output_docx = tmp_path / "output.docx"

    dlg = dialog_factory(str(docx_file))
    dlg._approve()
    sn_rc, img_rc = dlg.result

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx),
                            sn_row_col=sn_rc, img_row_col=img_rc,
                            keyword="", selected_models=["CP-8841-K9="])
    assert _get_injected_count(res) == 1

    out_doc = Document(str(output_docx))
    assert docx_table_inspector.get_inlines_count(out_doc.tables[0], 12, 0) == 0
    assert docx_table_inspector.get_inlines_count(out_doc.tables[1], 12, 0) >= 1


def test_tier3_start_table_with_keyword_and_model_filter(tmp_path, mock_docx_factory, mock_photo_factory):
    """48. Combines keyword prefix filtering with start_table and model filter."""
    _check_m1_support()
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"part_number": "CP-8841-K9=", "sn": "FVH0001"},
        {"part_number": "CP-8841-K9=", "sn": "FVH0002"},
        {"part_number": "CP-7821-K9=", "sn": "WZP0001"},
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "8841": ["FVH0001", "FVH0002"],
        "7821": ["WZP0001"]
    })
    output_docx = tmp_path / "output.docx"

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx),
                            keyword="FVH", start_table=2, selected_models=["CP-8841-K9="])
    assert _get_injected_count(res) == 1


def test_tier3_cancel_event_during_filtered_start_table_injection(tmp_path, mock_docx_factory, mock_photo_factory):
    """49. Early cancellation stops processing gracefully during filtered start_table run."""
    _check_m1_support()
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"part_number": "CP-7821-K9=", "sn": f"WZP{i:04d}"} for i in range(1, 10)
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "7821": [f"WZP{i:04d}" for i in range(1, 10)]
    })
    output_docx = tmp_path / "output.docx"

    cancel_evt = threading.Event()
    # Trigger cancellation immediately
    cancel_evt.set()

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx),
                            keyword="", start_table=2, selected_models=["CP-7821-K9="],
                            cancel_event=cancel_evt)
    assert _get_injected_count(res) == 0


def test_tier3_start_table_at_last_table_with_model_filter(tmp_path, mock_docx_factory, mock_photo_factory, docx_table_inspector):
    """50. start_table = N with model filter matches only the last table."""
    _check_m1_support()
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"part_number": "CP-7821-K9=", "sn": "WZP0001"},
        {"part_number": "CP-7821-K9=", "sn": "WZP0002"},
        {"part_number": "CP-8841-K9=", "sn": "FVH0001"},
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "7821": ["WZP0001", "WZP0002"],
        "8841": ["FVH0001"]
    })
    output_docx = tmp_path / "output.docx"

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx),
                            keyword="", start_table=3, selected_models=["CP-8841-K9="])
    assert _get_injected_count(res) == 1

    out_doc = Document(str(output_docx))
    assert docx_table_inspector.get_inlines_count(out_doc.tables[2], 12, 0) >= 1


def test_tier3_dialog_custom_coords_with_auto_detect_and_model_filter(tmp_path, mock_docx_factory, mock_photo_factory, dialog_factory):
    """51. Dialog custom coordinates, auto detection, and filtered injection together."""
    if not hasattr(app, "detect_docx_models"):
        pytest.skip("detect_docx_models not yet implemented in app.py")
    _check_m1_support()

    # Non-standard table layout: SN at (3, 2), Image at (8, 0)
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"rows": 12, "cols": 4, "sn_rc": (3, 2), "img_rc": (8, 0), "part_number": "CP-7821-K9=", "sn": "WZP0001"},
        {"rows": 12, "cols": 4, "sn_rc": (3, 2), "img_rc": (8, 0), "part_number": "CP-8841-K9=", "sn": "FVH0001"},
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "7821": ["WZP0001"],
        "8841": ["FVH0001"]
    })
    output_docx = tmp_path / "output.docx"

    models = app.detect_docx_models(str(docx_file), sn_row_col=(3, 2))
    assert "CP-7821-K9=" in models

    dlg = dialog_factory(str(docx_file))
    dlg.sn_rc = (3, 2)
    dlg.img_rc = (8, 0)
    dlg._approve()
    sn_rc, img_rc = dlg.result

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx),
                            sn_row_col=sn_rc, img_row_col=img_rc,
                            keyword="", selected_models=["CP-7821-K9="])
    assert _get_injected_count(res) == 1
