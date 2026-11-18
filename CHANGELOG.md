# Changelog

All notable changes to this project will be documented in this file.

## [1.6.2] - 2026-11-19
### Fixed
- **Photos in `Model/By_Number/<n>/` folders were lost or OCR'd for nothing.** A folder named only by a number (no `-total`, e.g. `S5755-H24UM4Y2CZ/by_Number/12/`) was taken as a serial folder named `12`, so the same numbers of different models collided (`[WARN] 53 duplicate folder name(s), using first`) and one model's photos were dropped, and all such photos went through OCR. They are now ชุดที่ folders: set n, total = the number of numbered sibling folders or the highest number (`_plain_set_number()`), mapped by model name and total like `12-280` folders. Real data (OneDrive export, 5 model folders, `By_SN` folders empty): 285 → 338 of 453 tables with photos, S5755 1 → 30 of 30, OCR 781 → 221 photos.

## [1.6.1] - 2026-11-18
### Fixed
- **The window no longer freezes while a document or photo folder is loaded.** Picking a .docx opened it up to 5 times on the UI thread (a 613-table document with photos takes ~8 s per open). It is now read once in the background with a small *Loading...* popup (file name and a moving bar, shown only when loading takes over 0.3 s), and the opened document is cached for the table count, device models and Configure Layout (`open_docx_readonly()`), which now opens instantly.
- **Read S/N from photos OCR'd photos of devices that are not in the document.** A photo root holding every device of a project (`7821/<serial>/`, `8841/<serial>/`), used with a document of only some of them, sent every other folder through OCR (~4–5 s per photo). Two checks now run before OCR:
  - **Model folders**: the document's device models are read first, and folders under a model folder whose model is not in the document, or is unticked in the Device Models filter, are skipped (`_ocr_skip_model_folders()`). This only applies when at least one model folder matches a document model, so a document whose models cannot be read is OCR'd as before.
  - **Serial-named folders**: a folder named by a serial that is not in the document is skipped (`_folder_named_like_other_serial()`); a name within distance 2 of a document serial (a typo) is still OCR'd.
  - Both are logged as `Skipping OCR in N folder(s) …`.
- Real data: 654 folders / 1,967 files (7821 + 8841), document with 43 serials of 8841 only: 1,827 photos queued for OCR (about 2 h+) → **21 s**, 43/43 tables filled.
### Changed
- **Read S/N from photos is faster and safer in per-device folders** (2+ sibling folders, up to 12 photos each):
  - reading stops at the first exact S/N label of a folder, and the folder's other photos get that serial (`OCR_EARLY_STOP_MAX_PHOTOS`);
  - a folder where no label can be read is no longer handed to the neighbouring folder's device by shooting time; its photos stay unassigned with a `[WARN]`;
  - OCR runs on up to 4 workers, one OpenThai engine each (`OCR_GROUP_MAX_WORKERS`);
  - a single flat folder (devices mixed) is still read photo by photo and grouped by shooting time.
- Real data, 43 devices / 129 photos (mostly HEIC) in folders not named by serial (OCR grouping only): **147 s → 49 s**; photos given to the wrong device **35 → 0** (correct 87 → 129 of 129). In 1.6.0 a device shot whose label could not be read went by shooting time to whichever device was photographed at a similar time, often in another folder.
- 4 new tests in `test/test_auto_match.py`.

## [1.6.0] - 2026-11-17
### Added
- **Auto Mode (Auto-match S/N)**: New `auto_match` option in `run_injection()` and an **Auto-match S/N (scan whole table)** checkbox (on by default). Every cell of each table is scanned for a serial matching a photo folder name, so no keyword or S/N/Image cell setup is needed and documents with different table layouts work as-is.
  - Serial matching ignores case, spaces and dashes (`wzp-301 000` = `WZP301000`), finds serials inside text (`S/N: WZP301000`) and in folder names with extra words (`Core2 10299X000016`); partial serials never match.
  - The photo cell is detected per table: the first empty merged row below the S/N, else the empty row under a `รูป/ภาพ/Photo` label, else the configured Image cell only if it is empty. Otherwise the table is skipped and nothing is overwritten.
  - Summary/list tables are skipped (more than one known serial, or 3+ serial-like values in the S/N column).
  - The log shows `[FIND]` with the cells used per table and lists folders not found in the document.
- **Nested Folder Scan** (`find_sn_leaf_folders()`): Auto mode finds image folders at any depth, e.g. `Model/By_SN/<serial>/`. Sequence folders like `195-280` are ignored and images placed directly in the selected folder are included.
- **Read S/N from Photos (OCR)**: New `photo_ocr` option and a **Read S/N from photos (OCR)** checkbox. Photos in a folder that is not named by a serial (e.g. 200 timestamp-named photos in one folder) are OCR'd for the S/N on the box label and matched against the serials labelled in the document (`extract_doc_serials()`), allowing up to 2 OCR character errors when the match is unique. Photos are ordered by EXIF capture time, and each run of photos with no readable S/N (wide device shots) is split at its largest time gap, so it does not matter whether the label was shot before or after the device photos, even within one set (`photo_capture_time()`, `assign_unread_photos()`). Without EXIF times, file-name order is used and each serial is topped up to the usual photo count. Warnings flag unusual photo counts, non-adjacent repeats and photos with no serial.
- **ชุดที่ (Set Number) Fallback**: New `set_number_fallback` option and a **Fallback: match ชุดที่ to folders like 12-280** checkbox (Auto mode, on by default in the GUI). Tables that still have no photos after S/N matching are filled from set-number folders (`Model/By_Number/12-280/`). The document's `ชุดที่ : N/T` rows are numbered continuously per Part Number across sections (e.g. 229 + 51 = 280), and model folders are matched to device groups by total and by name against the summary-table Description (`extract_set_entries()`, `find_set_number_folders()`, `map_set_folders_to_tables()`). Ambiguous folders are skipped and logged, never guessed. On the delivery document, the 104 `By_Number` folders gave exactly the same photos per table as the `By_SN` folders.
- **Run summary report** (`<output>_summary.txt`, **📄 Summary** button, `SUMMARY` line in `--cli`): per device group of the document (Description + Part Number, total units) it lists units with photos, *partial* units (fewer photos than usual) and *missing* units with ชุดที่, S/N and the reason (no photo folder, photo folder is empty, no photo cell in the table, …). Then: S/N repeated in several document tables, photo folders not in the document, photos whose S/N could not be read, and all warnings/errors. One-row summary tables are not counted as extra devices.
- **Command-line mode**: `app.py --cli` / `.exe --cli --docx ... --photos ... --out ... [--log ...]` runs Auto mode headless (batch jobs, testing the built exe).
- **Shrink photos** checkbox / `max_image_px` (default 2000 px long side, ~360 dpi at 14 cm): much smaller and faster `.docx` files. `--original-size` or unticking keeps the original photos.
- `test/test_auto_match.py` (23 tests).

### Performance
Full run on 669 photos / 336 tables (Smart Rotate on): **924 s → 257 s (3.6× faster)**, output **1.58 GB → 335 MB**.
- python-docx re-hashed every existing image on each `add_picture()` (duplicate check), which is O(n²) and took 249 s. The SHA-1 is now cached per image part.
- Each photo is decoded once, rotated in memory, shrunk and saved once (previously decoded and re-saved at full size 2–3 times).
- Photos are prepared (HEIC, EXIF, Smart Rotate, resize) in a thread pool with one OCR engine per worker, while insertion into Word stays in document order.

### Improved
- **Set-number sticker on the box ('31/100')**: when no S/N label can be read, the sticker on the box is read and mapped to the document's `ชุดที่ 31/100` and from there to its S/N (`build_set_sticker_lookup()`, `serial_from_set_sticker()`, `_read_foreground_sticker()`). Totals follow the document's continuous numbering (e.g. `230/280` → section 2 `1/51`). Big bold serif digits are read by zooming into the sticker in 9 variants (height × horizontal squeeze), and a result needs ≥ 2 agreeing variants. Safeguards:
  - only the **foreground** sticker counts (size × closeness to the photo centre), never boxes stacked behind (`271/280` behind `71/280`);
  - ambiguous stickers are skipped (`1/2` when two device types have 2 units);
  - a sticker read is ignored when photos shot within 90 s have a readable S/N label that disagrees; when both the S/N label and the sticker are read and disagree, the S/N label wins and the log warns (wrong sticker on a box).
  On the 200-photo `all/all` set, stickers alone: 100 correct, 0 wrong. Combined: 200/200 photos grouped correctly, only 4 left to the time-gap rule (was 84).
- **Small / far-away S/N labels**: if a photo gives no serial, each line that looks like `S/N …` or a serial is cropped and enlarged (text ~48 and ~72 px tall) and read again (`read_serial_from_photo()`). Fuzzy matching treats OCR look-alikes (O/0, 7/1, 8/B, 5/S, 6/G, 2/Z) as 1/3 of an error, accepts misread prefixes (`3/N`, `5/N`), and for 12+ character serials allows up to 2.5 errors only when the best match beats the next by ≥ 0.5. Example: a 991×742 photo of a TP-Link box (label text ~10 px) is now read as `33770F2001042`. On the 200-photo `all/all` set: 130 photos read (was 116), 0 wrong.

### Fixed
- **Same S/N in several document tables** (a copy error in the Word file): the photos were inserted into every one of those tables. Now they go only into the first table, the others fall through to the ชุดที่ fallback (their own set folder, i.e. the real box), and the log and summary report flag the document error.
- OCR photo grouping: photos with no readable S/N are only attached to a device through a chain of shots ≤ 90 s apart. Before, photos of other devices with unreadable labels (e.g. servers) were pulled onto the nearest readable serial (11 photos landed on one switch). They are now left unassigned and listed in the log.
- The "unusual photo count" warning ignores OCR groups that were discarded because an S/N folder exists.
- README build command now bundles the OCR model and its runtime imports.
- Serials starting with 3 digits (e.g. `21599990000GS0000208`) are no longer mistaken for camera folders like `100APPLE` in Auto mode.

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
