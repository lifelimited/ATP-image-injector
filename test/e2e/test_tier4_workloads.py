"""
Tier 4: Realistic Multi-Model Application Workloads (6 tests)
Simulates end-to-end user journeys and complex production scenarios:
1. Multi-model full document injection (7821 + 8841)
2. Selective single model batch injection (skip unselected model tables)
3. Mid-document resume workflow after partial run
4. Dialog inspection followed by selective injection
5. Deeply nested photo structures with mixed file extensions
6. Non-consecutive multi-model multi-pass injection (idempotence & preservation)
"""

import inspect
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


def test_tier4_multi_model_full_document_injection(tmp_path, mock_docx_factory, mock_photo_factory, docx_table_inspector):
    """52. Full production-style ATP document containing both 7821 and 8841 tables with full injection."""
    _check_m1_support()

    # 6 tables (3x 7821 and 3x 8841)
    docx_file = mock_docx_factory(tmp_path / "production_atp.docx", [
        {"part_number": "CP-7821-K9=", "model_name": "Cisco IP Phone 7821", "sn": f"WZP100{i}"} for i in range(1, 4)
    ] + [
        {"part_number": "CP-8841-K9=", "model_name": "Cisco IP Phone 8841", "sn": f"FVH200{i}"} for i in range(1, 4)
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "7821": [f"WZP100{i}" for i in range(1, 4)],
        "8841": [f"FVH200{i}" for i in range(1, 4)]
    })
    output_docx = tmp_path / "production_atp_output.docx"

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx),
                            keyword="", start_table=1,
                            selected_models=["CP-7821-K9=", "CP-8841-K9="])
    assert _get_injected_count(res) == 6

    out_doc = Document(str(output_docx))
    assert len(out_doc.tables) == 6
    for i, tbl in enumerate(out_doc.tables):
        assert docx_table_inspector.get_inlines_count(tbl, 12, 0) >= 1
        # Validate that SN cell remains intact
        expected_sn = f"WZP100{i+1}" if i < 3 else f"FVH200{i-2}"
        assert docx_table_inspector.get_sn(tbl, 5, 3) == expected_sn


def test_tier4_selective_single_model_batch_injection(tmp_path, mock_docx_factory, mock_photo_factory, docx_table_inspector):
    """53. 10-table document: technician selects only 7821; 8841 tables remain completely untouched."""
    _check_m1_support()

    # 10 tables: 5x 7821 and 5x 8841
    docx_file = mock_docx_factory(tmp_path / "atp_mixed_10.docx", [
        {"part_number": "CP-7821-K9=", "sn": f"WZP100{i}"} for i in range(1, 6)
    ] + [
        {"part_number": "CP-8841-K9=", "sn": f"FVH200{i}", "img_text": f"EMPTY_PLACEHOLDER_{i}"} for i in range(1, 6)
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "7821": [f"WZP100{i}" for i in range(1, 6)],
        "8841": [f"FVH200{i}" for i in range(1, 6)]
    })
    output_docx = tmp_path / "atp_mixed_output.docx"

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx),
                            keyword="", start_table=1, selected_models=["CP-7821-K9="])
    assert _get_injected_count(res) == 5

    out_doc = Document(str(output_docx))
    # Tables 0-4 (7821) must have pictures
    for i in range(5):
        assert docx_table_inspector.get_inlines_count(out_doc.tables[i], 12, 0) >= 1
    # Tables 5-9 (8841) must be completely untouched
    for i in range(5, 10):
        tbl = out_doc.tables[i]
        assert docx_table_inspector.get_inlines_count(tbl, 12, 0) == 0
        assert f"EMPTY_PLACEHOLDER_{i-4}" in docx_table_inspector.get_cell_text(tbl, 12, 0)


def test_tier4_mid_document_resume_after_partial_run(tmp_path, mock_docx_factory, mock_photo_factory, docx_table_inspector):
    """54. Resume injection from table 11 of 20-table document; tables 1-10 preserved."""
    _check_m1_support()

    # 20 tables: Tables 1-10 previously processed, Tables 11-20 pending
    docx_file = mock_docx_factory(tmp_path / "resume_doc.docx", [
        {"part_number": "CP-8841-K9=", "sn": f"FVH{i:04d}", "img_text": f"PRIOR_RUN_CONTENT_{i}"} for i in range(1, 11)
    ] + [
        {"part_number": "CP-8841-K9=", "sn": f"FVH{i:04d}"} for i in range(11, 21)
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "8841": [f"FVH{i:04d}" for i in range(11, 21)]
    })
    output_docx = tmp_path / "resume_output.docx"

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx),
                            keyword="", start_table=11, selected_models=["CP-8841-K9="])
    assert _get_injected_count(res) == 10

    out_doc = Document(str(output_docx))
    # Tables 0-9 preserved with original text
    for i in range(10):
        tbl = out_doc.tables[i]
        assert docx_table_inspector.get_inlines_count(tbl, 12, 0) == 0
        assert f"PRIOR_RUN_CONTENT_{i+1}" in docx_table_inspector.get_cell_text(tbl, 12, 0)
    # Tables 10-19 injected
    for i in range(10, 20):
        assert docx_table_inspector.get_inlines_count(out_doc.tables[i], 12, 0) >= 1


def test_tier4_dialog_inspection_followed_by_selective_injection(tmp_path, mock_docx_factory, mock_photo_factory, dialog_factory, docx_table_inspector):
    """55. Full UI workflow: Inspect table layout in dialog, confirm coordinates, then run selective injection."""
    _check_m1_support()

    docx_file = mock_docx_factory(tmp_path / "workflow_doc.docx", [
        {"part_number": "CP-7821-K9=", "sn": "WZP0001"},
        {"part_number": "CP-8841-K9=", "sn": "FVH0001"},
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "7821": ["WZP0001"],
        "8841": ["FVH0001"]
    })
    output_docx = tmp_path / "workflow_output.docx"

    # Step 1: User opens dialog and inspects
    dlg = dialog_factory(str(docx_file))
    dlg._on_cell_click(5, 3)
    dlg.mode.set("img")
    dlg._on_cell_click(12, 0)
    dlg._approve()
    sn_coord, img_coord = dlg.result

    # Step 2: User runs injection for 8841 starting at table 2
    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx),
                            sn_row_col=sn_coord, img_row_col=img_coord,
                            keyword="", start_table=2, selected_models=["CP-8841-K9="])
    assert _get_injected_count(res) == 1

    out_doc = Document(str(output_docx))
    assert docx_table_inspector.get_inlines_count(out_doc.tables[0], 12, 0) == 0
    assert docx_table_inspector.get_inlines_count(out_doc.tables[1], 12, 0) >= 1


def test_tier4_nested_photo_structure_with_mixed_extensions(tmp_path, mock_docx_factory, mock_photo_factory, docx_table_inspector):
    """56. Discovers photos in nested subdirectories with mixed case and file extensions (.jpg, .jpeg, .png)."""
    _check_m1_support()

    docx_file = mock_docx_factory(tmp_path / "input.docx", [
        {"part_number": "CP-7821-K9=", "sn": "WZP0001"},
        {"part_number": "CP-8841-K9=", "sn": "FVH0001"},
    ])
    # Photos in nested folders with mixed extensions
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "7821": ["WZP0001"],
        "8841": ["FVH0001"]
    }, nested=True, extensions=[".jpg", ".png", ".JPEG"])
    output_docx = tmp_path / "output.docx"

    res = app.run_injection(str(docx_file), str(photo_root), str(output_docx),
                            keyword="", start_table=1,
                            selected_models=["CP-7821-K9=", "CP-8841-K9="])
    assert _get_injected_count(res) == 2

    out_doc = Document(str(output_docx))
    assert docx_table_inspector.get_inlines_count(out_doc.tables[0], 12, 0) >= 1
    assert docx_table_inspector.get_inlines_count(out_doc.tables[1], 12, 0) >= 1


def test_tier4_non_consecutive_multi_model_injection_pass(tmp_path, mock_docx_factory, mock_photo_factory, docx_table_inspector):
    """57. Two-pass injection on alternating model document: Pass 1 injects 7821; Pass 2 injects 8841 into Pass 1 output."""
    _check_m1_support()

    # Interleaved document: T1=7821, T2=8841, T3=7821, T4=8841
    docx_file = mock_docx_factory(tmp_path / "interleaved.docx", [
        {"part_number": "CP-7821-K9=", "sn": "WZP0001"},
        {"part_number": "CP-8841-K9=", "sn": "FVH0001"},
        {"part_number": "CP-7821-K9=", "sn": "WZP0002"},
        {"part_number": "CP-8841-K9=", "sn": "FVH0002"},
    ])
    photo_root = mock_photo_factory(tmp_path / "photos", {
        "7821": ["WZP0001", "WZP0002"],
        "8841": ["FVH0001", "FVH0002"]
    })
    pass1_output = tmp_path / "pass1_output.docx"
    final_output = tmp_path / "final_output.docx"

    # Pass 1: Inject only 7821
    res1 = app.run_injection(str(docx_file), str(photo_root), str(pass1_output),
                             keyword="", start_table=1, selected_models=["CP-7821-K9="])
    assert _get_injected_count(res1) == 2

    doc_p1 = Document(str(pass1_output))
    assert docx_table_inspector.get_inlines_count(doc_p1.tables[0], 12, 0) >= 1
    assert docx_table_inspector.get_inlines_count(doc_p1.tables[1], 12, 0) == 0
    assert docx_table_inspector.get_inlines_count(doc_p1.tables[2], 12, 0) >= 1
    assert docx_table_inspector.get_inlines_count(doc_p1.tables[3], 12, 0) == 0

    # Pass 2: Inject 8841 into pass1_output
    res2 = app.run_injection(str(pass1_output), str(photo_root), str(final_output),
                             keyword="", start_table=1, selected_models=["CP-8841-K9="])
    assert _get_injected_count(res2) == 2

    doc_final = Document(str(final_output))
    # Now ALL 4 tables must have images!
    for i, tbl in enumerate(doc_final.tables):
        assert docx_table_inspector.get_inlines_count(tbl, 12, 0) >= 1, f"Table {i+1} should have image"
