# Changelog

All notable changes to this project will be documented in this file.

## [1.5.0] - 2026-11-16
### Added
- **Starting Page / Table Selection**: Added a 1-based **Start Table** selector in the main GUI (`start_table` in `run_injection()`) so users can skip cover, revision, and summary tables and begin photo injection at any table in the document.
- **Arbitrary Table Layout Configuration**: Upgraded `TableSelectionDialog` with `◀ Previous`, `Next ▶`, and a direct `Table X of N` selector dropdown so S/N and Image cell coordinates can be inspected and mapped on any table in the document (no longer restricted to Table 1).
- **Automatic Device Model Detection & Multi-Select Filtering**:
  - Added `detect_docx_models()`, `detect_folder_models()`, `extract_doc_part_descriptions()`, and `model_matches()` to automatically detect device models from both Word tables and photo folder hierarchies.
  - Supports multi-vendor formats including Cisco (`CP-7821-K9=`, `CP-8841-K9=`), Huawei switches/APs/servers (`S6750-H48Y8C`, `S5755-H24UM4Y2CZ`, `AirEngine9700-M1`, `AirEngine8776-X7T-T`, `AirEngine6776I-X6ETH`, `TaiShan 200`), and TP-Link (`ARCHER-TBE400E`) by cross-referencing `Part Number:` rows with summary table descriptions.
  - Added a **Device Models Filter** checklist card in the main GUI with **Select All** / **Deselect All** controls so only selected device models are injected.
- **Test Suites**: Added comprehensive unit, GUI, adversarial, and E2E test suites (`test/test_table_selection.py`, `test/test_model_filtering.py`, `test/test_adversarial_tier5.py`, and `test/e2e/`).

### Fixed
- **Camera Subfolder False Positives**: Excluded camera subdirectories (`DCIM`, `100APPLE`, `CAMERA`, `100MEDIA`, etc.) from being misclassified as device model folders.
- **Boundary & Input Hardening**: Added safe bounds clamping and type coercion for `start_table`, `initial_table_idx`, and `selected_models`, including 0-table document reset handling.

## [1.4.0] - 2026-11-15
### Changed
- **OCR engine upgraded to [iapp/OpenThai2.0-OCR-Text](https://huggingface.co/iapp/OpenThai2.0-OCR-Text)** (Thai + English, ONNX, CPU-only). EasyOCR/PyTorch is no longer required.
- Smart Rotate uses the engine's built-in sideways detection plus a 180° comparison (2 OCR passes instead of 4).
- `mac_scanner.py` v3.0 uses the shared engine; includes low-confidence lines for LCD text.

### Added
- `ocr_engine.py`: pluggable OCR backend (`OCR_BACKEND=openthai|easyocr`), thread-safe singleton, auto-download of model files.
- `models/OpenThai2.0-OCR-Text/`: bundled model + runtime (Apache-2.0, see `NOTICE`).
- `test/test_ocr_engine.py`: self-contained OCR / orientation / MAC-extraction tests.

## [1.3.2] - 2026-05-30
### Added
- **Universal Document Support**: The application can now process any document image (not just IP-phones), intelligently calculating orientation based on OCR text density.
- **Improved Smart-Rotation Algorithm**: Replaced keyword-specific tracking with an adaptive text-density scoring algorithm to robustly handle generic documentation processing.

### Fixed
- **HEIC Processing Bug**: Resolved a bug where `.HEIC` and `.HEIF` images were incorrectly bypassed during the format conversion phase, ensuring reliable processing and injection.

## [1.2.1] - Previous Versions
- Added support for macOS compatibility (directory separator normalization).
- Initial framework for image processing and document injection based on IP-Phone specific tracking (keyword boosting).
