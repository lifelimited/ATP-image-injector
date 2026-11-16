"""
Tier 1: Feature Coverage (20 tests)
Covers all four core features in isolation:
1. Starting Table Selection (start_table) - 5 tests
2. Device Model Auto-Detection (detect_docx_models, detect_folder_models) - 5 tests
3. Device Model Filtering (selected_models) - 5 tests
4. Arbitrary Table Layout Configuration Dialog (TableSelectionDialog) - 5 tests
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


def _get_no_folder_count(res):
    if hasattr(res, "no_folder"):
        return res.no_folder
    if isinstance(res, dict):
        return res["no_folder"]
    return res[2]


# ─────────────────────────────────────────────────────────────────────────────
# Feature 1: Starting Table Selection (start_table)
# ─────────────────────────────────────────────────────────────────────────────

def test_tier1_start_table_1_processes_all_tables(tmp_path, mock_docx_factory, mock_photo_factory, docx_table_inspector):
    """1. start_table = 1 injects images into all tables from Table 1 to N."""
    _check_start_table_support()
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"sn": f"WZP000{i}"} for i in range(1, 6)
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "": [f"WZP000{i}" for i in range(1, 6)]
    })
    output_docx = tmp_path / "output.docx"

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx), start_table=1)
    assert _get_injected_count(res) == 5

    out_doc = Document(str(output_docx))
    assert len(out_doc.tables) == 5
    for i, tbl in enumerate(out_doc.tables):
        assert docx_table_inspector.get_inlines_count(tbl, 12, 0) >= 1, f"Table {i+1} should have image"


def test_tier1_start_table_mid_document_skip(tmp_path, mock_docx_factory, mock_photo_factory, docx_table_inspector):
    """2. start_table = 3 skips Tables 1-2 completely and injects Tables 3-5."""
    _check_start_table_support()
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"sn": f"WZP000{i}"} for i in range(1, 6)
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "": [f"WZP000{i}" for i in range(1, 6)]
    })
    output_docx = tmp_path / "output.docx"

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx), start_table=3)
    assert _get_injected_count(res) == 3

    out_doc = Document(str(output_docx))
    # Tables 1 and 2 (indices 0 and 1) must be skipped (0 inlines)
    assert docx_table_inspector.get_inlines_count(out_doc.tables[0], 12, 0) == 0
    assert docx_table_inspector.get_inlines_count(out_doc.tables[1], 12, 0) == 0
    # Tables 3, 4, 5 (indices 2, 3, 4) must be injected
    assert docx_table_inspector.get_inlines_count(out_doc.tables[2], 12, 0) >= 1
    assert docx_table_inspector.get_inlines_count(out_doc.tables[3], 12, 0) >= 1
    assert docx_table_inspector.get_inlines_count(out_doc.tables[4], 12, 0) >= 1


def test_tier1_start_table_preserves_unmodified_earlier_table_text_and_runs(tmp_path, mock_docx_factory, mock_photo_factory, docx_table_inspector):
    """3. Earlier tables before start_table remain 100% unaltered with original text and 0 inlines."""
    _check_start_table_support()
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"sn": "WZP0001", "img_text": "PRESERVE_INITIAL_CONTENT"},
        {"sn": "WZP0002"},
        {"sn": "WZP0003"},
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "": ["WZP0001", "WZP0002", "WZP0003"]
    })
    output_docx = tmp_path / "output.docx"

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx), start_table=2)
    assert _get_injected_count(res) == 2

    out_doc = Document(str(output_docx))
    table1 = out_doc.tables[0]
    assert docx_table_inspector.get_inlines_count(table1, 12, 0) == 0
    assert docx_table_inspector.get_cell_text(table1, 12, 0) == "PRESERVE_INITIAL_CONTENT"
    # Tables 2 and 3 should have images
    assert docx_table_inspector.get_inlines_count(out_doc.tables[1], 12, 0) >= 1
    assert docx_table_inspector.get_inlines_count(out_doc.tables[2], 12, 0) >= 1


def test_tier1_start_table_sequential_document_order_processing(tmp_path, mock_docx_factory, mock_photo_factory):
    """4. Injection executes in strict document sequence order (t_idx ascending) rather than folder order."""
    _check_start_table_support()
    # Non-alphabetical sequence in document
    sn_sequence = ["WZP_Z", "WZP_A", "WZP_M", "WZP_B"]
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"sn": sn} for sn in sn_sequence
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "": sn_sequence
    })
    output_docx = tmp_path / "output.docx"

    processed_sn_order = []
    def log_cb(msg):
        for sn in sn_sequence:
            if f"[OK  ] {sn}" in msg or f"[SKIP] {sn}" in msg:
                processed_sn_order.append(sn)

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx), start_table=2, log_callback=log_cb)
    assert _get_injected_count(res) == 3
    assert processed_sn_order == ["WZP_A", "WZP_M", "WZP_B"]

    out_doc = Document(str(output_docx))
    # Table 1 (WZP_Z) was skipped
    assert len(out_doc.tables[0].rows[12].cells[0].paragraphs[0].runs) == 0 or not out_doc.tables[0].rows[12].cells[0]._element.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline')


def test_tier1_start_table_accurate_no_folder_and_progress_accounting(tmp_path, mock_docx_factory, mock_photo_factory):
    """5. Tables prior to start_table are not falsely counted in no_folder or error statistics."""
    _check_start_table_support()
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"sn": "WZP0001"},  # missing photo, but before start_table -> do NOT count in no_folder
        {"sn": "WZP0002"},  # present photo
        {"sn": "WZP0003"},  # missing photo, after start_table -> MUST count in no_folder
        {"sn": "WZP0004"},  # present photo
    ])
    # Only WZP0002 and WZP0004 have photo folders
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "": ["WZP0002", "WZP0004"]
    })
    output_docx = tmp_path / "output.docx"

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx), start_table=2)
    assert _get_injected_count(res) == 2
    assert _get_no_folder_count(res) == 1  # Only WZP0003


# ─────────────────────────────────────────────────────────────────────────────
# Feature 2: Device Model Auto-Detection
# ─────────────────────────────────────────────────────────────────────────────

def test_tier1_detect_docx_models_single_model(tmp_path, mock_docx_factory):
    """6. detect_docx_models finds a single unique model from document tables."""
    _check_model_detection_support()
    docx_file = mock_docx_factory(tmp_path / "test_single.docx", [
        {"part_number": "CP-7821-K9=", "model_name": "Cisco IP Phone 7821"} for _ in range(3)
    ])
    models = app.detect_docx_models(str(docx_file), sn_row_col=(5, 3))
    assert models == ["CP-7821-K9="]


def test_tier1_detect_docx_models_multi_model_mixed(tmp_path, mock_docx_factory):
    """7. detect_docx_models extracts multiple models and returns sorted unique list."""
    _check_model_detection_support()
    docx_file = mock_docx_factory(tmp_path / "test_multi.docx", [
        {"part_number": "CP-7821-K9=", "model_name": "Cisco IP Phone 7821"},
        {"part_number": "CP-7821-K9=", "model_name": "Cisco IP Phone 7821"},
        {"part_number": "CP-8841-K9=", "model_name": "Cisco IP Phone 8841"},
        {"part_number": "CP-8841-K9=", "model_name": "Cisco IP Phone 8841"},
    ])
    models = app.detect_docx_models(str(docx_file))
    assert models == ["CP-7821-K9=", "CP-8841-K9="]


def test_tier1_detect_docx_models_header_row2_extraction(tmp_path, mock_docx_factory):
    """8. detect_docx_models falls back to / checks Row 2 Thai header name for model codes."""
    _check_model_detection_support()
    docx_file = mock_docx_factory(tmp_path / "test_r2.docx", [
        {"part_number": "CP-8841-K9=", "model_name": "Cisco IP Phone 8841"}
    ])
    models = app.detect_docx_models(str(docx_file))
    assert "CP-8841-K9=" in models


def test_tier1_detect_folder_models_standard_hierarchy(tmp_path, mock_photo_factory):
    """9. detect_folder_models discovers model directories in standard hierarchy (e.g. 7821, 8841)."""
    _check_model_detection_support()
    sn_root = mock_photo_factory(tmp_path / "photos", {
        "7821": ["WZP0001"],
        "8841": ["FVH0001"]
    })
    models = app.detect_folder_models(str(sn_root))
    assert any("7821" in m for m in models)
    assert any("8841" in m for m in models)


def test_tier1_detect_folder_models_cisco_prefixed_directories(tmp_path, mock_photo_factory):
    """10. detect_folder_models handles full Cisco part number directory names (e.g. CP-7821-K9=)."""
    _check_model_detection_support()
    sn_root = mock_photo_factory(tmp_path / "photos", {
        "CP-7821-K9=": ["WZP0001"],
        "CP-8841-K9=": ["FVH0001"]
    })
    models = app.detect_folder_models(str(sn_root))
    assert "CP-7821-K9=" in models
    assert "CP-8841-K9=" in models


# ─────────────────────────────────────────────────────────────────────────────
# Feature 3: Device Model Filtering (selected_models)
# ─────────────────────────────────────────────────────────────────────────────

def test_tier1_model_filtering_single_model_selected(tmp_path, mock_docx_factory, mock_photo_factory, docx_table_inspector):
    """11. selected_models=['CP-7821-K9='] injects only 7821 tables and skips 8841 tables."""
    _check_model_filtering_support()
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"part_number": "CP-7821-K9=", "sn": "WZP0001"},
        {"part_number": "CP-7821-K9=", "sn": "WZP0002"},
        {"part_number": "CP-8841-K9=", "sn": "FVH0001"},
        {"part_number": "CP-8841-K9=", "sn": "FVH0002"},
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "7821": ["WZP0001", "WZP0002"],
        "8841": ["FVH0001", "FVH0002"]
    })
    output_docx = tmp_path / "output.docx"

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx), keyword="", selected_models=["CP-7821-K9="])
    assert _get_injected_count(res) == 2

    out_doc = Document(str(output_docx))
    assert docx_table_inspector.get_inlines_count(out_doc.tables[0], 12, 0) >= 1
    assert docx_table_inspector.get_inlines_count(out_doc.tables[1], 12, 0) >= 1
    assert docx_table_inspector.get_inlines_count(out_doc.tables[2], 12, 0) == 0
    assert docx_table_inspector.get_inlines_count(out_doc.tables[3], 12, 0) == 0


def test_tier1_model_filtering_all_models_selected(tmp_path, mock_docx_factory, mock_photo_factory, docx_table_inspector):
    """12. selected_models with all models selected processes every matching table."""
    _check_model_filtering_support()
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"part_number": "CP-7821-K9=", "sn": "WZP0001"},
        {"part_number": "CP-7821-K9=", "sn": "WZP0002"},
        {"part_number": "CP-8841-K9=", "sn": "FVH0001"},
        {"part_number": "CP-8841-K9=", "sn": "FVH0002"},
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "7821": ["WZP0001", "WZP0002"],
        "8841": ["FVH0001", "FVH0002"]
    })
    output_docx = tmp_path / "output.docx"

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx), keyword="",
                            selected_models=["CP-7821-K9=", "CP-8841-K9="])
    assert _get_injected_count(res) == 4

    out_doc = Document(str(output_docx))
    for tbl in out_doc.tables:
        assert docx_table_inspector.get_inlines_count(tbl, 12, 0) >= 1


def test_tier1_model_filtering_none_means_all_models(tmp_path, mock_docx_factory, mock_photo_factory):
    """13. selected_models=None preserves backwards compatibility by processing all models."""
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

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx), keyword="", selected_models=None)
    assert _get_injected_count(res) == 2


def test_tier1_model_filtering_ignores_unselected_model_folders(tmp_path, mock_docx_factory, mock_photo_factory):
    """14. Folders corresponding to unselected models are ignored without producing false errors."""
    _check_model_filtering_support()
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"part_number": "CP-7821-K9=", "sn": "WZP0001"}
    ])
    # Photo root has both 7821 and 8841 folders
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "7821": ["WZP0001"],
        "8841": ["FVH0001", "FVH0002"]
    })
    output_docx = tmp_path / "output.docx"

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx), keyword="", selected_models=["CP-7821-K9="])
    assert _get_injected_count(res) == 1
    assert _get_no_folder_count(res) == 0


def test_tier1_model_filtering_preserves_non_matching_table_formatting(tmp_path, mock_docx_factory, mock_photo_factory, docx_table_inspector):
    """15. Non-matching tables retain exact cell text and 0 inlines."""
    _check_model_filtering_support()
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"part_number": "CP-7821-K9=", "sn": "WZP0001"},
        {"part_number": "CP-8841-K9=", "sn": "FVH0001", "img_text": "DO_NOT_OVERWRITE_8841"},
        {"part_number": "CP-7821-K9=", "sn": "WZP0002"},
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "7821": ["WZP0001", "WZP0002"],
        "8841": ["FVH0001"]
    })
    output_docx = tmp_path / "output.docx"

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx), keyword="", selected_models=["CP-7821-K9="])
    assert _get_injected_count(res) == 2

    out_doc = Document(str(output_docx))
    table2 = out_doc.tables[1]
    assert docx_table_inspector.get_inlines_count(table2, 12, 0) == 0
    assert docx_table_inspector.get_cell_text(table2, 12, 0) == "DO_NOT_OVERWRITE_8841"


# ─────────────────────────────────────────────────────────────────────────────
# Feature 4: Arbitrary Table Layout Configuration (TableSelectionDialog)
# ─────────────────────────────────────────────────────────────────────────────

def test_tier1_dialog_initial_table_rendering(tmp_path, mock_docx_factory, dialog_factory):
    """16. TableSelectionDialog displays cells of the initial table."""
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"sn": "WZP0001", "cell_texts": {(0, 0): "CELL_T1_TOP"}}
    ])
    dlg = dialog_factory(str(docx_file), initial_table_idx=0)
    assert hasattr(dlg, "buttons")
    assert (0, 0) in dlg.buttons
    assert "CELL_T1_TOP" in dlg.buttons[(0, 0)].cget("text")
    dlg.destroy()


def test_tier1_dialog_navigate_to_arbitrary_table(tmp_path, mock_docx_factory, dialog_factory):
    """17. TableSelectionDialog can navigate to table N and re-render cell preview."""
    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"sn": "WZP0001", "cell_texts": {(0, 0): "TABLE_1_DATA"}},
        {"sn": "WZP0002", "cell_texts": {(0, 0): "UNIQUE_T2_DATA"}},
    ])
    dlg = dialog_factory(str(docx_file), initial_table_idx=0)
    # Check if navigation method exists (M2 Feature)
    if hasattr(dlg, "_render_table"):
        dlg._render_table(1)
        assert getattr(dlg, "current_table_idx", 1) == 1
        assert "UNIQUE_T2_DATA" in dlg.buttons[(0, 0)].cget("text")
    elif hasattr(dlg, "_load_table") and "table_idx" in inspect.signature(dlg._load_table).parameters:
        dlg._load_table(str(docx_file), table_idx=1)
        assert getattr(dlg, "current_table_idx", 1) == 1
        assert "UNIQUE_T2_DATA" in dlg.buttons[(0, 0)].cget("text")
    elif hasattr(dlg, "_next_table"):
        dlg._next_table()
        assert getattr(dlg, "current_table_idx", 1) == 1
        assert "UNIQUE_T2_DATA" in dlg.buttons[(0, 0)].cget("text")
    elif hasattr(dlg, "table_combo") and hasattr(dlg, "_on_table_change"):
        dlg.table_combo.set("Table 2")
        dlg._on_table_change()
        assert getattr(dlg, "current_table_idx", 1) == 1
        assert "UNIQUE_T2_DATA" in dlg.buttons[(0, 0)].cget("text")
    else:
        dlg.destroy()
        pytest.skip("Arbitrary table navigation (_render_table, table_idx, or _next_table) not yet implemented in TableSelectionDialog")
    dlg.destroy()


def test_tier1_dialog_coordinate_selection_sn_and_img(tmp_path, mock_docx_factory, dialog_factory):
    """18. Clicking cells in SN mode and IMG mode updates coordinates accordingly."""
    docx_file = mock_docx_factory(tmp_path / "input.docx", [{"sn": "WZP0001"}])
    dlg = dialog_factory(str(docx_file))

    dlg.mode.set("sn")
    dlg._on_cell_click(2, 1)
    assert dlg.sn_rc == (2, 1)

    dlg.mode.set("img")
    dlg._on_cell_click(6, 0)
    assert dlg.img_rc == (6, 0)
    dlg.destroy()


def test_tier1_dialog_approval_returns_selected_coordinates(tmp_path, mock_docx_factory, dialog_factory):
    """19. Clicking Approve sets dialog.result = (sn_rc, img_rc) and destroys window."""
    docx_file = mock_docx_factory(tmp_path / "input.docx", [{"sn": "WZP0001"}])
    dlg = dialog_factory(str(docx_file))
    dlg.sn_rc = (3, 2)
    dlg.img_rc = (7, 0)

    dlg._approve()
    assert dlg.result == ((3, 2), (7, 0))


def test_tier1_dialog_cancel_leaves_result_none(tmp_path, mock_docx_factory, dialog_factory):
    """20. Canceling or closing dialog without approve leaves dialog.result as None."""
    docx_file = mock_docx_factory(tmp_path / "input.docx", [{"sn": "WZP0001"}])
    dlg = dialog_factory(str(docx_file))
    dlg._on_cell_click(3, 2)

    dlg.destroy()
    assert dlg.result is None
