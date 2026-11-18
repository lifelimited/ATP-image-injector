# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

## [1.7.0-beta] - 2026-11-18 (pre-release)
### Added
- **Verify photos (OCR every folder)**: new `verify_photos` option in `run_injection()`, a **Verify photos (OCR every folder, slower)** checkbox under Auto-match (off by default) and `--verify` in `--cli`. Photos in S/N folders and set-number folders are OCR'd too, so a photo of device B left in device A's folder goes into B's table (`verify_folder_photos()`, `resolve_photo_moves()`). A wrong move is worse than no move, so it is strict:
  - only an exact or near-exact S/N label read (fuzzy distance ≤ 1.0, `VERIFY_MAX_FUZZY`) moves a photo; a set-number sticker may confirm the folder's own device but never moves a photo;
  - a read is not trusted for a move when another document serial differs from it only by OCR look-alikes (`P6H` read as `PGH`, both real serials), or when a fuzzy read is close to the folder's own serial (own label `NJ6` read as `NI6` → `NTG`) (`_move_read_doubt()`);
  - such a look-alike read is trusted again only with stronger evidence: the label line is re-read zoomed in 6 variants (40/56/72 px × plain / contrast+sharpen) and must read exactly the same ≥ 5 times with no variant showing the look-alike (`_confirm_label_read()`), **and** at least 2 photos of the folder must pass this while no photo reads the look-alike or the folder's own serial;
  - swapped folders are solved as a pair (`_swap_pair_proposals()`): folder B also moves when A moves whole to B on trusted reads and every label in B names A (no read of B), or when each folder has one label photo confirmed by the zoomed re-read naming the other; never when the two serials are close;
  - all readable labels in a folder name device B (none name A) and there are **at least 2 such label photos** → the whole folder goes to B, unread photos included. With a **single** label photo of B only that photo moves: the unread photos may be A's own (e.g. a stray label photo next to an unreadable wide shot) and stay, unless B's folder in turn moves whole to A (a swap). Logged as `[VERF] … only 1 label photo …`;
  - a mixed folder is split by label, unread photos follow a label shot ≤ 90 s apart, and the log warns;
  - photos never go into a table that keeps its own photos (its S/N folder, set folder or OCR group), and photos of one device found in several folders are not moved; both cases are logged as `[WARN] … NOT moved` (two swapped folders do move);
  - only the Word output changes, files on disk are never moved or renamed;
  - the summary report has a new **Photos moved (wrong folder)** section (photo, from folder, to table with ชุดที่ and S/N, reason), and a table whose folder held only another device's photos says so under *Missing*.
- Validation on hard-link copies of 40 S/N folders (~120 photos) of a delivered document, with 33 photos deliberately misplaced (4 devices' photos dropped into another device's folder, 3 pairs of folder names swapped, 3 single photos moved), three random samples: **0 wrong moves**; 18, 6 and 16 of 33 moved correctly, the rest left in place with a `[WARN]`/`[VERF]` line. The first version (no look-alike guard) made 2 wrong moves; with the guard alone only 6 moved per sample. Runtime 36 s → 134–167 s (≈ +1 s per photo, Smart Rotate off).
  - Verify OCR runs in parallel (up to 4 workers, one OpenThai engine each, `_verify_ocr_reads()`): on correctly sorted real folders 48 photos 75 s → 47 s and 60 photos 109 s → 77 s, both with **0 moves**.
- 16 tests in `test/test_auto_match.py` (fake OCR, made-up serials).

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
