"""
Unit tests for Starting Table / Page Selection in run_injection().
Self-contained: generates mock docx and mock images dynamically via tmp_path.
"""
import sys
from pathlib import Path
import pytest
from docx import Document
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import app


def create_mock_docx(file_path: Path, tables_spec: list[dict]) -> Path:
    """
    Creates a docx file with tables according to tables_spec.
    Each spec dict: {'sn': 'WZP001', 'model': 'CP-7821-K9=', 'placeholder': 'UNTOUCHED_0'}
    """
    doc = Document()
    for spec in tables_spec:
        table = doc.add_table(rows=14, cols=4)
        # Row 2: Model header
        table.rows[2].cells[0].text = f"ลำดับที่ 1\nName : {spec.get('model', 'CP-7821-K9=')} P/N : {spec.get('model', 'CP-7821-K9=')}"
        # Row 4: Header
        table.rows[4].cells[1].text = "Detail"
        table.rows[4].cells[2].text = "Part Number"
        table.rows[4].cells[3].text = "Serial Number"
        # Row 5: S/N row (sn_row_col = (5, 3))
        table.rows[5].cells[1].text = f"Cisco Phone {spec.get('model', 'CP-7821-K9=')}"
        table.rows[5].cells[2].text = spec.get('model', 'CP-7821-K9=')
        table.rows[5].cells[3].text = spec['sn']
        # Row 12: Image cell (img_row_col = (12, 0))
        table.rows[12].cells[0].text = spec.get('placeholder', f"UNTOUCHED_{spec['sn']}")
    doc.save(file_path)
    return file_path


def create_mock_photo_root(base_dir: Path, sn_list: list[str]) -> Path:
    """Creates directory structure with a dummy JPEG image in each SN subfolder."""
    for sn in sn_list:
        sn_dir = base_dir / sn
        sn_dir.mkdir(parents=True, exist_ok=True)
        img = Image.new("RGB", (50, 50), color="white")
        img.save(sn_dir / "sample.jpg", "JPEG")
    return base_dir


def count_drawings_in_cell(cell) -> int:
    return len(cell._element.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline'))


@pytest.fixture
def table_test_env(tmp_path):
    docx_in = tmp_path / "input.docx"
    docx_out = tmp_path / "output.docx"
    photos_dir = tmp_path / "photos"

    specs = [
        {"sn": "WZP1001", "model": "CP-7821-K9=", "placeholder": "UNTOUCHED_TABLE_1"},
        {"sn": "WZP1002", "model": "CP-7821-K9=", "placeholder": "UNTOUCHED_TABLE_2"},
        {"sn": "WZP1003", "model": "CP-7821-K9=", "placeholder": "UNTOUCHED_TABLE_3"},
    ]
    create_mock_docx(docx_in, specs)
    create_mock_photo_root(photos_dir, ["WZP1001", "WZP1002", "WZP1003"])

    return docx_in, docx_out, photos_dir


def test_start_table_default_processes_all(table_test_env):
    """Default start_table=1 should inject all 3 tables."""
    docx_in, docx_out, photos_dir = table_test_env
    res = app.run_injection(str(docx_in), str(photos_dir), str(docx_out), auto_rotate=False, smart_rotate=False, start_table=1)
    injected = res[0] if isinstance(res, tuple) else res.get('injected')
    assert injected == 3
    assert res.img_count == 3
    assert res.no_folder == 0

    out_doc = Document(docx_out)
    for t_idx in range(3):
        cell = out_doc.tables[t_idx].rows[12].cells[0]
        assert count_drawings_in_cell(cell) >= 1
        assert "UNTOUCHED" not in cell.text


def test_start_table_skips_first_table(table_test_env):
    """start_table=2 must skip table 1 (untouched) and inject tables 2 and 3."""
    docx_in, docx_out, photos_dir = table_test_env
    res = app.run_injection(str(docx_in), str(photos_dir), str(docx_out), auto_rotate=False, smart_rotate=False, start_table=2)
    injected = res[0] if isinstance(res, tuple) else res.get('injected')
    assert injected == 2

    out_doc = Document(docx_out)
    # Table 1 (index 0) must be 100% untouched
    t1_cell = out_doc.tables[0].rows[12].cells[0]
    assert t1_cell.text.strip() == "UNTOUCHED_TABLE_1"
    assert count_drawings_in_cell(t1_cell) == 0

    # Tables 2 and 3 must have images injected
    for t_idx in (1, 2):
        cell = out_doc.tables[t_idx].rows[12].cells[0]
        assert count_drawings_in_cell(cell) >= 1
        assert "UNTOUCHED" not in cell.text


def test_start_table_last_table_only(table_test_env):
    """start_table=3 must skip tables 1 and 2 and only inject table 3."""
    docx_in, docx_out, photos_dir = table_test_env
    res = app.run_injection(str(docx_in), str(photos_dir), str(docx_out), auto_rotate=False, smart_rotate=False, start_table=3)
    injected = res[0] if isinstance(res, tuple) else res.get('injected')
    assert injected == 1

    out_doc = Document(docx_out)
    # Tables 1 and 2 must remain untouched
    assert out_doc.tables[0].rows[12].cells[0].text.strip() == "UNTOUCHED_TABLE_1"
    assert count_drawings_in_cell(out_doc.tables[0].rows[12].cells[0]) == 0
    assert out_doc.tables[1].rows[12].cells[0].text.strip() == "UNTOUCHED_TABLE_2"
    assert count_drawings_in_cell(out_doc.tables[1].rows[12].cells[0]) == 0

    # Table 3 must be injected
    t3_cell = out_doc.tables[2].rows[12].cells[0]
    assert count_drawings_in_cell(t3_cell) >= 1
    assert "UNTOUCHED" not in t3_cell.text


def test_start_table_greater_than_total(table_test_env):
    """start_table greater than total tables skips all tables without modification."""
    docx_in, docx_out, photos_dir = table_test_env
    res = app.run_injection(str(docx_in), str(photos_dir), str(docx_out), auto_rotate=False, smart_rotate=False, start_table=10)
    injected = res[0] if isinstance(res, tuple) else res.get('injected')
    assert injected == 0
    assert res.img_count == 0

    # If file was saved, verify all tables are untouched
    if docx_out.exists():
        out_doc = Document(docx_out)
        for t_idx in range(3):
            cell = out_doc.tables[t_idx].rows[12].cells[0]
            assert count_drawings_in_cell(cell) == 0
            assert cell.text.startswith("UNTOUCHED_")


def test_start_table_zero_or_negative_fallback(table_test_env):
    """start_table <= 0 should gracefully fallback to table 1 or clamp."""
    docx_in, docx_out, photos_dir = table_test_env
    res = app.run_injection(str(docx_in), str(photos_dir), str(docx_out), auto_rotate=False, smart_rotate=False, start_table=0)
    injected = res[0] if isinstance(res, tuple) else res.get('injected')
    assert injected == 3


def test_sequential_document_table_order(tmp_path):
    """Verify that tables are executed in document table order (t_idx ascending), not serial number order."""
    docx_in = tmp_path / "order_doc.docx"
    docx_out = tmp_path / "order_out.docx"
    photos_dir = tmp_path / "order_photos"

    # Table 1 has WZP999 (alphabetically last), Table 2 has WZP111 (alphabetically first)
    specs = [
        {"sn": "WZP999", "model": "CP-7821-K9=", "placeholder": "UNTOUCHED_FIRST"},
        {"sn": "WZP111", "model": "CP-7821-K9=", "placeholder": "UNTOUCHED_SECOND"},
    ]
    create_mock_docx(docx_in, specs)
    create_mock_photo_root(photos_dir, ["WZP999", "WZP111"])

    execution_order = []
    def log_tracker(msg):
        if "WZP" in msg and ("OK" in msg or "SKIP" in msg):
            execution_order.append(msg)

    # When running all tables, Table 1 (WZP999) must be processed before Table 2 (WZP111)
    res_all = app.run_injection(str(docx_in), str(photos_dir), str(docx_out), auto_rotate=False, smart_rotate=False, start_table=1, log_callback=log_tracker)
    assert res_all.injected == 2
    assert len(execution_order) >= 2
    assert "WZP999" in execution_order[0]
    assert "WZP111" in execution_order[1]

    # When start_table=2, Table 1 (WZP999) must be skipped, Table 2 (WZP111) must be injected
    execution_order.clear()
    res = app.run_injection(str(docx_in), str(photos_dir), str(docx_out), auto_rotate=False, smart_rotate=False, start_table=2, log_callback=log_tracker)
    injected = res[0] if isinstance(res, tuple) else res.get('injected')
    assert injected == 1

    out_doc = Document(docx_out)
    assert out_doc.tables[0].rows[12].cells[0].text.strip() == "UNTOUCHED_FIRST"
    assert count_drawings_in_cell(out_doc.tables[1].rows[12].cells[0]) >= 1


# ─────────────────────────────────────────────
# Milestone 2: TableSelectionDialog & App GUI Tests
# ─────────────────────────────────────────────

@pytest.fixture(scope="session")
def headless_root():
    """Session-scoped root App instance for stable Tkinter UI testing."""
    root = app.App()
    root.withdraw()
    yield root
    try:
        root.destroy()
    except Exception:
        pass


def test_dialog_initial_table_idx_navigation(tmp_path, headless_root):
    """Verify dialog opens at specified initial_table_idx and renders correct table."""
    docx_file = tmp_path / "multi.docx"
    doc = Document()
    t1 = doc.add_table(rows=3, cols=2)
    t1.rows[0].cells[0].text = "T1_DATA"
    t2 = doc.add_table(rows=3, cols=2)
    t2.rows[0].cells[0].text = "T2_DATA"
    doc.save(str(docx_file))

    dlg = app.TableSelectionDialog(headless_root, str(docx_file), (0, 0), (1, 0), initial_table_idx=1)
    assert dlg.current_table_idx == 1
    assert "T2_DATA" in dlg.buttons[(0, 0)].cget("text")
    dlg.destroy()


def test_dialog_prev_next_stepping(tmp_path, headless_root):
    """Verify _next_table and _prev_table step sequentially with clamping."""
    docx_file = tmp_path / "multi3.docx"
    doc = Document()
    for i in range(3):
        t = doc.add_table(rows=2, cols=2)
        t.rows[0].cells[0].text = f"T{i+1}_CONTENT"
    doc.save(str(docx_file))

    dlg = app.TableSelectionDialog(headless_root, str(docx_file), (0, 0), (1, 0), initial_table_idx=0)
    assert dlg.current_table_idx == 0

    dlg._next_table()
    assert dlg.current_table_idx == 1
    assert "T2_CONTENT" in dlg.buttons[(0, 0)].cget("text")

    dlg._next_table()
    assert dlg.current_table_idx == 2
    assert "T3_CONTENT" in dlg.buttons[(0, 0)].cget("text")

    dlg._next_table()  # Clamped at end
    assert dlg.current_table_idx == 2

    dlg._prev_table()
    assert dlg.current_table_idx == 1

    dlg._prev_table()
    assert dlg.current_table_idx == 0

    dlg._prev_table()  # Clamped at start
    assert dlg.current_table_idx == 0
    dlg.destroy()


def test_dialog_single_table_disables_nav_buttons(tmp_path, headless_root):
    """Verify single-table document disables both Prev and Next buttons."""
    docx_file = tmp_path / "single.docx"
    doc = Document()
    doc.add_table(rows=2, cols=2)
    doc.save(str(docx_file))

    dlg = app.TableSelectionDialog(headless_root, str(docx_file), (0, 0), (1, 0))
    if hasattr(dlg, "prev_btn"):
        assert str(dlg.prev_btn.cget("state")) == "disabled"
    if hasattr(dlg, "next_btn"):
        assert str(dlg.next_btn.cget("state")) == "disabled"
    dlg.destroy()


def test_dialog_inhomogeneous_grid_cleanup(tmp_path, headless_root):
    """Verify grid cleanly removes old buttons when switching table dimensions."""
    docx_file = tmp_path / "inhomogeneous.docx"
    doc = Document()
    doc.add_table(rows=10, cols=4)  # 40 buttons
    doc.add_table(rows=2, cols=2)   # 4 buttons
    doc.save(str(docx_file))

    dlg = app.TableSelectionDialog(headless_root, str(docx_file), (0, 0), (1, 0), initial_table_idx=0)
    assert len(dlg.buttons) == 40

    dlg._render_table(1)
    assert len(dlg.buttons) == 4
    assert (9, 3) not in dlg.buttons
    dlg.destroy()


def test_dialog_preserves_selected_coordinates(tmp_path, headless_root):
    """Verify coordinates set on table A remain selected when navigating to table B."""
    docx_file = tmp_path / "multi_coords.docx"
    doc = Document()
    doc.add_table(rows=6, cols=4)
    doc.add_table(rows=6, cols=4)
    doc.save(str(docx_file))

    dlg = app.TableSelectionDialog(headless_root, str(docx_file), (0, 0), (1, 0))
    dlg._on_cell_click(5, 3)  # sn_rc
    dlg.mode.set("img")
    dlg._on_cell_click(4, 2)  # img_rc
    assert dlg.sn_rc == (5, 3)
    assert dlg.img_rc == (4, 2)

    dlg._next_table()
    assert dlg.sn_rc == (5, 3)
    assert dlg.img_rc == (4, 2)
    # Check button highlight colors
    assert dlg.buttons[(5, 3)].cget("bg") == app.UI_COLORS["select_sn"]
    assert dlg.buttons[(4, 2)].cget("bg") == app.UI_COLORS["select_img"]
    dlg.destroy()


def test_app_gui_start_table_validation_and_clamping(tmp_path, headless_root):
    """Verify App Card 1 start table validation, bounds clamping, and dynamic max table updating."""
    main_app = headless_root

    # Initial state
    main_app.start_table_var.set("1")
    main_app._total_tables = 1
    assert main_app.start_table_var.get() == "1"
    assert main_app._get_start_table() == 1

    # Clamp logic
    main_app.start_table_var.set("-5")
    assert main_app._get_start_table() == 1
    main_app.start_table_var.set("0")
    assert main_app._get_start_table() == 1
    main_app.start_table_var.set("not_a_number")
    assert main_app._get_start_table() == 1

    # Dynamic docx total tables update
    docx_file = tmp_path / "five_tables.docx"
    doc = Document()
    for _ in range(5):
        doc.add_table(rows=2, cols=2)
    doc.save(str(docx_file))

    main_app._update_docx_info(str(docx_file))
    assert main_app._total_tables == 5
    assert main_app.start_table_spinbox.cget("to") == 5
    assert "5" in main_app.total_tables_label.cget("text")

    # Clamped to total_tables
    main_app.start_table_var.set("99")
    assert main_app._get_start_table() == 5
    main_app._on_start_table_changed()
    assert main_app.start_table_var.get() == "5"


def test_app_gui_dynamic_model_filtering_and_merging(tmp_path, headless_root):
    """Verify App Card 2 device model checklist detection, synonym merging, and select/deselect all."""
    docx_file = tmp_path / "models.docx"
    doc = Document()
    t1 = doc.add_table(rows=6, cols=4)
    t1.rows[2].cells[0].text = "Name : CP-7821-K9="
    t1.rows[5].cells[3].text = "WZP1001"
    t2 = doc.add_table(rows=6, cols=4)
    t2.rows[2].cells[0].text = "Name : CP-8841-K9="
    t2.rows[5].cells[3].text = "FVH1001"
    doc.save(str(docx_file))

    photo_root = tmp_path / "photos"
    p1 = photo_root / "7821"
    p1.mkdir(parents=True)
    (p1 / "test.jpg").write_bytes(b"dummy")

    p2 = photo_root / "8865"
    p2.mkdir(parents=True)
    (p2 / "test.jpg").write_bytes(b"dummy")

    main_app = headless_root
    main_app.docx_var.set(str(docx_file))
    main_app.sn_root_var.set(str(photo_root))
    main_app._refresh_detected_models()

    assert "CP-7821-K9=" in main_app.model_vars
    assert "CP-8841-K9=" in main_app.model_vars
    assert "8865" in main_app.model_vars
    assert "7821" not in main_app.model_vars  # Merged under CP-7821-K9=

    # All initially selected
    assert all(var.get() for var in main_app.model_vars.values())

    # Deselect all
    main_app._deselect_all_models()
    assert all(not var.get() for var in main_app.model_vars.values())

    # Select all
    main_app._select_all_models()
    assert all(var.get() for var in main_app.model_vars.values())


def test_file_scan_runs_in_background_and_updates_ui(headless_root, tmp_path):
    """Selecting a document/folder reads it in a worker thread; the UI is updated when done."""
    import time
    import docx as _docx
    docx_file = tmp_path / "bg.docx"
    doc = _docx.Document()
    for _ in range(4):
        t = doc.add_table(rows=6, cols=4)
        t.rows[2].cells[0].text = "Name : CP-7821-K9="
    doc.save(str(docx_file))
    photo_root = tmp_path / "photos_bg" / "8865"
    photo_root.mkdir(parents=True)
    (photo_root / "a.jpg").write_bytes(b"dummy")

    app_ = headless_root
    called = []
    app_.sn_root_var.set(str(photo_root.parent))
    app_._start_file_scan("test", then=lambda: called.append(True))
    app_.docx_var.set(str(docx_file))            # newer scan replaces the first one
    deadline = time.time() + 15
    while time.time() < deadline and not ("CP-7821-K9=" in app_.model_vars and "(of 4)" in app_.total_tables_label.cget("text")):
        app_.update()
        time.sleep(0.02)
    assert "(of 4)" in app_.total_tables_label.cget("text")
    assert "CP-7821-K9=" in app_.model_vars and "8865" in app_.model_vars
    assert called == [True]                       # 'then' runs once, after the newest scan
    assert getattr(app_, "_busy_win", None) is None
