import os
from pathlib import Path
import pytest
from PIL import Image
from docx import Document
import tkinter as tk

# Ensure Tcl/Tk libraries are located correctly on Windows
os.environ.setdefault("TCL_LIBRARY", r"C:\Python313\tcl\tcl8.6")
os.environ.setdefault("TK_LIBRARY", r"C:\Python313\tcl\tk8.6")


@pytest.fixture
def dummy_image_factory():
    """Generates small valid JPEG/PNG images for testing."""
    def _create(path: Path, color: str = "blue", size: tuple = (50, 50), fmt: str = "JPEG"):
        path.parent.mkdir(parents=True, exist_ok=True)
        img = Image.new("RGB", size, color=color)
        img.save(str(path), fmt)
        return path
    return _create


@pytest.fixture
def mock_docx_factory():
    """Builds realistic ATP Word documents with N tables."""
    def _create(path: Path, tables: list[dict], default_rows: int = 18, default_cols: int = 4):
        doc = Document()
        for idx, t_data in enumerate(tables):
            seq = t_data.get("seq", idx + 1)
            model_name = t_data.get("model_name", "Cisco IP Phone 7821")
            part_number = t_data.get("part_number", "CP-7821-K9=")
            sn = t_data.get("sn", f"WZP{seq:04d}")
            sn_rc = t_data.get("sn_rc", (5, 3))
            img_rc = t_data.get("img_rc", (12, 0))
            rows = t_data.get("rows", default_rows)
            cols = t_data.get("cols", default_cols)

            tbl = doc.add_table(rows=rows, cols=cols)
            if rows > 2 and cols > 0:
                tbl.rows[2].cells[0].text = f"ลำดับที่ {seq} Name : {part_number} P/N : {part_number}"
            if rows > 4 and cols > 3:
                tbl.rows[4].cells[1].text = "Detail"
                tbl.rows[4].cells[2].text = "Part Number"
                tbl.rows[4].cells[3].text = "Serial Number"
            if rows > sn_rc[0] and cols > sn_rc[1]:
                if cols > 0:
                    tbl.rows[sn_rc[0]].cells[0].text = f"1.{seq}"
                if cols > 1:
                    tbl.rows[sn_rc[0]].cells[1].text = model_name
                if cols > 2:
                    tbl.rows[sn_rc[0]].cells[2].text = part_number
                tbl.rows[sn_rc[0]].cells[sn_rc[1]].text = sn

            if "img_text" in t_data and rows > img_rc[0] and cols > img_rc[1]:
                tbl.rows[img_rc[0]].cells[img_rc[1]].text = t_data["img_text"]

            if "cell_texts" in t_data:
                for (r, c), val in t_data["cell_texts"].items():
                    if r < rows and c < cols:
                        tbl.rows[r].cells[c].text = str(val)

            if idx < len(tables) - 1:
                doc.add_page_break()

        path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(path))
        return path
    return _create


@pytest.fixture
def mock_photo_factory(dummy_image_factory):
    """Generates structured photo directories."""
    def _create(root_dir: Path, model_structure: dict[str, list[str]], nested: bool = False,
                count_per_sn: int = 2, extensions: list[str] = None):
        """
        model_structure: { "7821": ["WZP0001", "WZP0002"], "8841": ["FVH0001"] }
        or { "": ["WZP0001", "WZP0002"] } for flat layout.
        """
        if extensions is None:
            extensions = [".jpg"]
        for model_dir, sn_list in model_structure.items():
            base = root_dir / model_dir if model_dir else root_dir
            for sn in sn_list:
                sn_folder = base / sn / sn if nested else base / sn
                sn_folder.mkdir(parents=True, exist_ok=True)
                for i in range(count_per_sn):
                    ext = extensions[i % len(extensions)]
                    fmt = "PNG" if ext.lower() == ".png" else "JPEG"
                    fname = f"IMG_{i+1:04d}{ext}"
                    dummy_image_factory(sn_folder / fname, color="blue" if i % 2 == 0 else "green", fmt=fmt)
        return root_dir
    return _create


@pytest.fixture
def docx_table_inspector():
    """Helper methods to verify injected docx content."""
    class Inspector:
        @staticmethod
        def get_inlines(table, row=12, col=0):
            cell = table.rows[row].cells[col]
            return cell._element.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline')

        @staticmethod
        def get_inlines_count(table, row=12, col=0):
            return len(Inspector.get_inlines(table, row, col))

        @staticmethod
        def get_sn(table, row=5, col=3):
            return table.rows[row].cells[col].text.strip()

        @staticmethod
        def get_cell_text(table, row, col):
            return table.rows[row].cells[col].text.strip()
    return Inspector()


@pytest.fixture(scope="session")
def _session_tk():
    """Session-scoped root Tkinter instance for stable UI testing."""
    root = tk.Tk()
    root.withdraw()
    yield root
    try:
        root.destroy()
    except Exception:
        pass


@pytest.fixture
def headless_tk(_session_tk):
    """Provides a hidden root Tkinter instance with required theme colors."""
    _session_tk.COLORS = {
        "bg": "#1a1a2e", "card": "#16213e", "card_alt": "#0f3460",
        "input_bg": "#0a0a1a", "accent": "#e94560", "text": "#ffffff",
        "text_dim": "#a0a0b0", "btn_run": "#00b894", "btn_run_fg": "#ffffff",
        "btn_stop": "#d63031", "btn_stop_fg": "#ffffff"
    }
    return _session_tk


@pytest.fixture
def dialog_factory(headless_tk):
    """Factory creating TableSelectionDialog with backwards compatibility."""
    def _create(docx_path: str, current_sn_rc=(5, 3), current_img_rc=(12, 0), initial_table_idx=0):
        from app import TableSelectionDialog
        import inspect
        sig = inspect.signature(TableSelectionDialog.__init__)
        if "initial_table_idx" in sig.parameters:
            return TableSelectionDialog(headless_tk, docx_path, current_sn_rc, current_img_rc, initial_table_idx=initial_table_idx)
        else:
            dlg = TableSelectionDialog(headless_tk, docx_path, current_sn_rc, current_img_rc)
            if hasattr(dlg, "_render_table"):
                dlg._render_table(initial_table_idx)
            return dlg
    return _create


@pytest.fixture(autouse=True)
def mock_ocr(monkeypatch):
    """Ensures smart auto-rotate doesn't invoke heavy OCR models during tests."""
    import app
    monkeypatch.setattr(app, "get_best_orientation", lambda img_path, log_cb=None: 0)
