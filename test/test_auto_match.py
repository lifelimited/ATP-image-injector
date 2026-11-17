"""Tests for Auto mode: match S/N anywhere in a table and detect the photo cell."""
import os
import sys

import pytest
from docx import Document
from PIL import Image

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import app  # noqa: E402


def _make_photo(folder, name="a.jpg"):
    os.makedirs(folder, exist_ok=True)
    Image.new("RGB", (40, 30), "blue").save(os.path.join(folder, name))


def _merged_empty_row(table, r):
    row = table.rows[r]
    row.cells[0].merge(row.cells[-1])


def _count_pictures(table):
    return len(table._tbl.xpath('.//*[local-name()="drawing"]'))


def _build_doc(path):
    """Two layouts + a summary table listing two serials."""
    doc = Document()
    # T1: summary table listing both serials (must be skipped)
    t = doc.add_table(rows=3, cols=3)
    t.rows[0].cells[2].text = "Serial Number"
    t.rows[1].cells[2].text = "10299X000012"
    t.rows[2].cells[2].text = "WZP-301000"
    # T2: label/value layout, photo row = merged empty last row
    t = doc.add_table(rows=4, cols=2)
    t.rows[0].cells[0].text = "Part Number:"
    t.rows[0].cells[1].text = "02355UEL"
    t.rows[1].cells[0].text = "Serial Number:"
    t.rows[1].cells[1].text = "10299X000012"
    t.rows[2].cells[0].text = "Note"
    _merged_empty_row(t, 3)
    # T3: S/N embedded in text, photo below a 'รูปภาพ' label row (not merged)
    t = doc.add_table(rows=5, cols=2)
    t.rows[0].cells[0].text = "S/N: wzp301000"
    t.rows[1].cells[0].text = "Result"
    t.rows[1].cells[1].text = "Pass"
    t.rows[2].cells[0].text = "รูปภาพ"
    t.rows[4].cells[0].text = "Signature :"
    # T4: serial with no folder
    t = doc.add_table(rows=3, cols=2)
    t.rows[0].cells[1].text = "ZZZ999999"
    _merged_empty_row(t, 2)
    doc.save(path)


def test_normalize_and_index():
    assert app.normalize_sn(" wzp-301 000 ") == "WZP301000"
    idx = app.build_sn_index(["WZP-301000", "Core2 10299X000016", "AP01"])
    assert idx["WZP301000"] == "WZP-301000"
    assert idx["10299X000016"] == "Core2 10299X000016"
    assert "AP01" not in idx  # too short to be a serial


def test_index_drops_ambiguous_tokens():
    idx = app.build_sn_index(["A 10299X000016", "B 10299X000016"])
    assert "10299X000016" not in idx


def test_match_sns_in_text():
    idx = app.build_sn_index(["WZP301000"])
    assert app.match_sns_in_text("Serial Number: WZP301000", idx) == ["WZP301000"]
    assert app.match_sns_in_text("WZP3010001", idx) == []  # no partial matches
    assert app.match_sns_in_text("", idx) == []


def test_auto_locate_layouts(tmp_path):
    p = tmp_path / "doc.docx"
    _build_doc(str(p))
    doc = Document(str(p))
    idx = app.build_sn_index(["10299X000012", "WZP-301000"])

    assert app.auto_locate_in_table(doc.tables[0], idx)["status"] == "ambiguous"

    r = app.auto_locate_in_table(doc.tables[1], idx)
    assert r["status"] == "ok" and r["sn"] == "10299X000012"
    assert r["sn_rc"] == (1, 1) and r["img_rc"] == (3, 0)

    r = app.auto_locate_in_table(doc.tables[2], idx)
    assert r["status"] == "ok" and r["sn"] == "WZP-301000"
    assert r["img_rc"] == (3, 0)

    assert app.auto_locate_in_table(doc.tables[3], idx) is None


def test_no_image_cell_uses_only_empty_fallback(tmp_path):
    doc = Document()
    t = doc.add_table(rows=2, cols=2)
    t.rows[0].cells[0].text = "WZP301000"
    t.rows[1].cells[0].text = "filled"
    t.rows[1].cells[1].text = "filled"
    idx = app.build_sn_index(["WZP301000"])
    assert app.auto_locate_in_table(t, idx, fallback_img_rc=(1, 0))["status"] == "no_image_cell"
    t.rows[1].cells[0].text = ""
    r = app.auto_locate_in_table(t, idx, fallback_img_rc=(1, 0))
    assert r["status"] == "ok" and r["img_rc"] == (1, 0)


def test_run_injection_auto_mode(tmp_path):
    p = tmp_path / "doc.docx"
    out = tmp_path / "out.docx"
    _build_doc(str(p))
    root = tmp_path / "photos"
    _make_photo(str(root / "10299X000012"))
    _make_photo(str(root / "WZP-301000"))
    _make_photo(str(root / "WZP-301000"), "b.jpg")

    logs = []
    res = app.run_injection(str(p), str(root), str(out), keyword="NOPE",
                            auto_match=True, auto_rotate=False, smart_rotate=False,
                            log_callback=logs.append)
    assert res.injected == 2 and res.img_count == 3 and not res.errors

    doc = Document(str(out))
    assert _count_pictures(doc.tables[0]) == 0  # summary table untouched
    assert _count_pictures(doc.tables[1]) == 1
    assert _count_pictures(doc.tables[2]) == 2
    assert doc.tables[2].rows[4].cells[0].text.strip() == "Signature :"
    assert any("summary" in line for line in logs)


def test_run_injection_manual_mode_unchanged(tmp_path):
    """Without auto_match, keyword 'NOPE' matches nothing (legacy behavior)."""
    p = tmp_path / "doc.docx"
    _build_doc(str(p))
    root = tmp_path / "photos"
    _make_photo(str(root / "WZP-301000"))
    res = app.run_injection(str(p), str(root), str(tmp_path / "o.docx"), keyword="NOPE",
                            auto_rotate=False, smart_rotate=False)
    assert res.injected == 0


def test_leaf_scanner_deep_layout(tmp_path):
    """Model/By_SN/<SN>/ + duplicate Model/By_Number/<n-total>/ + camera subfolder."""
    root = tmp_path / "Missing"
    _make_photo(str(root / "AirEngine8776-X7-T" / "By_SN" / "21599990000GS0000208"))
    _make_photo(str(root / "AirEngine8776-X7-T" / "By_Number" / "199-280"))
    _make_photo(str(root / "S5755" / "By_SN" / "4E9990000999" / "DCIM" / "100APPLE"))
    m = app.find_sn_leaf_folders(str(root))
    assert sorted(m) == ["21599990000GS0000208", "4E9990000999"]
    assert m.sn_to_model["21599990000GS0000208"] == "AirEngine8776-X7-T"
    assert m.sn_to_model["4E9990000999"] == "S5755"


def test_list_table_with_single_known_serial_is_skipped():
    doc = Document()
    t = doc.add_table(rows=5, cols=2)
    t.rows[0].cells[1].text = "Serial Number"
    for i, sn in enumerate(["4E9990000962", "4E9990000891", "4E9990000900", "4E9990000999"], start=1):
        t.rows[i].cells[1].text = sn
    idx = app.build_sn_index(["4E9990000999"])
    assert app.auto_locate_in_table(t, idx)["status"] == "ambiguous"


# ── OCR grouping (photos not sorted into S/N folders) ──

class _FakeEngine:
    """Returns OCR lines keyed by the image's pixel colour (set up per test)."""
    def __init__(self, by_color):
        self.by_color = by_color

    def read(self, img):
        color = img.getpixel((0, 0))
        lines = self.by_color.get(color, [])

        class _R:
            def texts(self_inner, include_dropped=False):
                return lines
        return _R()


def _photo(path, color):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    Image.new("RGB", (20, 20), color).save(path)
    return path


def test_extract_doc_serials_by_label():
    doc = Document()
    t = doc.add_table(rows=3, cols=3)
    t.rows[0].cells[0].text = "Description"
    t.rows[0].cells[1].text = "Serial Number"
    t.rows[1].cells[0].text = "AirEngine8776-X7T-T"
    t.rows[1].cells[1].text = "21599990000GS0000208"
    t2 = doc.add_table(rows=2, cols=2)
    t2.rows[0].cells[0].text = "Serial Number:"
    t2.rows[0].cells[1].text = "4E9990000999"
    t2.rows[1].cells[0].text = "S/N: WZP301000"
    ser = app.extract_doc_serials(doc)
    assert set(ser) == {"21599990000GS0000208", "4E9990000999", "WZP301000"}  # model name excluded


def test_find_serial_in_ocr_lines_exact_and_fuzzy():
    ser = {"21599990000GS0000345": "21599990000GS0000345", "21599990000GS0000346": "21599990000GS0000346"}
    assert app.find_serial_in_ocr_lines(["S/N: 21599990000GS0000346"], ser) == ("21599990000GS0000346", "exact")
    # OCR read O for 0 and split the number; O/0 is a look-alike, so it costs only 1/3
    key, how = app.find_serial_in_ocr_lines(["S/N, 215999900O0 GS0000345"], ser)
    assert key == "21599990000GS0000345" and how == "fuzzy0.3"
    # equally close to two serials -> refuse
    assert app.find_serial_in_ocr_lines(["21599990000GS000034X"], ser) == (None, "")
    assert app.find_serial_in_ocr_lines(["Made in China"], ser) == (None, "")


def test_find_serial_small_label_misreads():
    """Zoomed OCR of a small box label (made-up serial): 'S/N:33770F2001042' read as '3/N 3,37DF2O07O4/2'."""
    sibs = ["33770F2001042", "33770F2001073", "3374599003545", "336B765000997", "3374599003084"]
    ser = {s: s for s in sibs}
    assert app.find_serial_in_ocr_lines(["3/N 3,37DF2O07O4/2"], ser)[0] == "33770F2001042"
    assert app.find_serial_in_ocr_lines(["S/N 33970F2001O42"], ser)[0] == "33770F2001042"
    # a word ending in 'sn' is not an S/N prefix
    assert app._OCR_SN_PREFIX_REGEX.search("Version 1") is None


def test_levenshtein_lookalikes_are_cheap():
    assert app._levenshtein("2O07", "2001") < 1.0          # O/0 and 7/1 look alike
    assert app._levenshtein("2X01", "2001") == 1.0         # X/0 is a real mistake


class _BoxLine:
    def __init__(self, text, box):
        self.text, self.box, self.dropped, self.confidence = text, box, True, 0.5


class _ZoomEngine:
    """Full photo: tiny unreadable S/N line with a box. Enlarged crop: readable."""
    def read(self, img):
        big = img.width < 400

        class _R:
            lines = [] if not big else [_BoxLine("S/N-22คลOF2O 1O4 2", [[20, 40], [120, 40], [120, 50], [20, 50]])]

            def texts(self_inner, include_dropped=False):
                return ["S/N 33970F2001O42"] if not big else [ln.text for ln in self_inner.lines]
        return _R()


def test_read_serial_from_photo_zooms_into_small_label(tmp_path):
    p = str(tmp_path / "box.png")
    Image.new("RGB", (300, 120), "white").save(p)
    ser = {s: s for s in ["33770F2001042", "33770F2001073"]}
    key, how = app.read_serial_from_photo(p, _ZoomEngine(), ser)
    assert key == "33770F2001042" and how.endswith("+zoom")


def test_group_photos_label_then_wide_shot(tmp_path):
    """No EXIF time: file-name order, unread photos topped up to the usual count (2)."""
    ser = {"AAA1234567": "AAA1234567", "BBB7654321": "BBB7654321"}
    eng = _FakeEngine({(255, 0, 0): ["S/N: AAA1234567"], (0, 0, 255): ["S/N: BBB7654321"]})
    d = str(tmp_path / "all")
    imgs = [_photo(os.path.join(d, "1.png"), (255, 0, 0)),   # label A
            _photo(os.path.join(d, "2.png"), (0, 255, 0)),   # wide shot -> A
            _photo(os.path.join(d, "3.png"), (0, 0, 255)),   # label B
            _photo(os.path.join(d, "4.png"), (0, 255, 0))]   # wide shot -> B
    groups, unassigned = app.group_photos_by_ocr(imgs, ser, engine=eng)
    assert [os.path.basename(p) for p in groups["AAA1234567"]] == ["1.png", "2.png"]
    assert [os.path.basename(p) for p in groups["BBB7654321"]] == ["3.png", "4.png"]
    assert unassigned == []


def test_assign_unread_device_photo_first_by_count():
    """Device photo shot BEFORE its label, no times: still grouped correctly."""
    seq = [("u1", None), ("A", "A"), ("u2", None), ("B", "B"), ("u3", None), ("C", "C")]
    groups, method = app.assign_unread_photos(seq)
    assert method == "photo count"
    assert groups == {"A": ["u1", "A"], "B": ["u2", "B"], "C": ["u3", "C"]}


def test_assign_unread_mixed_order_by_time_gap():
    """A: label then device; B: device then label; C: label then 2 device photos."""
    seq = [("A", "A"), ("a1", None), ("b1", None), ("B", "B"), ("C", "C"), ("c1", None), ("c2", None)]
    times = [0, 4, 60, 63, 120, 123, 126]
    groups, method = app.assign_unread_photos(seq, times)
    assert method == "time gap"
    assert groups == {"A": ["A", "a1"], "B": ["b1", "B"], "C": ["C", "c1", "c2"]}


def test_group_photos_sorted_by_capture_time(tmp_path):
    """File names out of shooting order (e.g. renamed copies): EXIF time decides."""
    ser = {"AAA1234567": "AAA1234567", "BBB7654321": "BBB7654321"}
    eng = _FakeEngine({(255, 0, 0): ["S/N: AAA1234567"], (0, 0, 255): ["S/N: BBB7654321"]})
    d = str(tmp_path / "all")
    shots = {"1.png": ((0, 0, 255), 63), "2.png": ((0, 255, 0), 60),   # B: device then label
             "3.png": ((0, 255, 0), 3), "4.png": ((255, 0, 0), 0)}     # A: label then device
    imgs = [_photo(os.path.join(d, n), c) for n, (c, _) in shots.items()]
    t = {os.path.join(d, n): ts for n, (_, ts) in shots.items()}
    groups, _ = app.group_photos_by_ocr(imgs, ser, engine=eng, time_fn=lambda p: t[p])
    assert sorted(os.path.basename(p) for p in groups["AAA1234567"]) == ["3.png", "4.png"]
    assert sorted(os.path.basename(p) for p in groups["BBB7654321"]) == ["1.png", "2.png"]


def test_photo_capture_time_reads_exif(tmp_path):
    from PIL import Image as PILImage
    p = str(tmp_path / "x.jpg")
    img = PILImage.new("RGB", (10, 10))
    exif = img.getexif()
    exif.get_ifd(0x8769)[36867] = "2026:11:17 11:28:33"
    img.save(p, exif=exif)
    t1 = app.photo_capture_time(p)
    assert t1 is not None
    assert app.photo_capture_time(str(tmp_path / "missing.jpg")) is None


def test_run_injection_photo_ocr_flat_folder(tmp_path, monkeypatch):
    doc = Document()
    for sn in ("AAA1234567", "BBB7654321"):
        t = doc.add_table(rows=3, cols=2)
        t.rows[0].cells[0].text = "Serial Number:"
        t.rows[0].cells[1].text = sn
        _merged_empty_row(t, 2)
    p = tmp_path / "doc.docx"
    doc.save(str(p))
    root = tmp_path / "all"
    for name, color in (("1.png", (255, 0, 0)), ("2.png", (0, 255, 0)),
                        ("3.png", (0, 0, 255)), ("4.png", (0, 255, 0))):
        _photo(str(root / name), color)
    eng = _FakeEngine({(255, 0, 0): ["S/N: AAA1234567"], (0, 0, 255): ["S/N: BBB7654321"]})
    import ocr_engine
    monkeypatch.setattr(ocr_engine, "get_engine", lambda *a, **k: eng)

    out = tmp_path / "out.docx"
    res = app.run_injection(str(p), str(root), str(out), auto_match=True, photo_ocr=True,
                            auto_rotate=False, smart_rotate=False)
    assert res.injected == 2 and res.img_count == 4 and not res.errors
    d = Document(str(out))
    assert _count_pictures(d.tables[0]) == 2 and _count_pictures(d.tables[1]) == 2

    # Without photo_ocr the flat folder matches nothing
    res2 = app.run_injection(str(p), str(root), str(tmp_path / "o2.docx"), auto_match=True,
                             photo_ocr=False, auto_rotate=False, smart_rotate=False)
    assert res2.injected == 0


# ── ชุดที่ (set number) fallback ──

def _detail_table(doc, item, n, total, pn, sn):
    t = doc.add_table(rows=5, cols=2)
    for r, (k, v) in enumerate([("รายการ:", item), ("ชุดที่ :", f"{n}/{total}"),
                                ("Part Number:", pn), ("Serial Number:", sn)]):
        t.rows[r].cells[0].text = k
        t.rows[r].cells[1].text = v
    _merged_empty_row(t, 4)
    return t


def _summary_table(doc, pn_desc):
    t = doc.add_table(rows=1 + len(pn_desc), cols=3)
    for c, h in enumerate(["ลำดับ", "Part Number", "Description"]):
        t.rows[0].cells[c].text = h
    for i, (pn, desc) in enumerate(pn_desc, start=1):
        t.rows[i].cells[0].text = str(i)
        t.rows[i].cells[1].text = pn
        t.rows[i].cells[2].text = desc


def _set_doc(path):
    """AP group split over 2 sections (2 + 1 = 3), plus two 1-unit groups."""
    doc = Document()
    _summary_table(doc, [("50088471", "AirEngine8776-X7T-T"), ("02355UEL", "S6750-H48Y8C"),
                         ("98012016", "S5735-S24T4XE-V2")])
    _detail_table(doc, "AP indoor A", 1, 2, "50088471", "AP0000000001")
    _detail_table(doc, "AP indoor A", 2, 2, "50088471", "AP0000000002")
    _detail_table(doc, "AP indoor B", 1, 1, "50088471", "AP0000000003")
    _detail_table(doc, "Core Switch", 1, 1, "02355UEL", "CORE00000001")
    _detail_table(doc, "Mgmt Switch", 1, 1, "98012016", "MGMT00000001")
    doc.save(path)


def test_extract_set_entries_numbers_across_sections(tmp_path):
    p = str(tmp_path / "d.docx")
    _set_doc(p)
    entries = app.extract_set_entries(Document(p))
    ap = [(e["n"], e["total"], e["index"], e["group_total"]) for e in entries if e["pn"] == "50088471"]
    assert ap == [(1, 2, 1, 3), (2, 2, 2, 3), (1, 1, 3, 3)]
    assert all(e["last_row"] == 3 for e in entries)


def test_find_set_number_folders_model_name(tmp_path):
    root = tmp_path / "photos"
    _make_photo(str(root / "AirEngine8776-X7-T" / "By_Number" / "3-3"))
    _make_photo(str(root / "AirEngine8776-X7-T" / "By_SN" / "AP0000000003"))  # not a set folder
    found = app.find_set_number_folders(str(root))
    assert [(f["model"], f["n"], f["total"]) for f in found] == [("AirEngine8776-X7-T", 3, 3)]


def test_map_set_folders_by_name_and_total(tmp_path):
    p = str(tmp_path / "d.docx")
    _set_doc(p)
    doc = Document(p)
    entries = app.extract_set_entries(doc)
    pn = app.extract_doc_part_descriptions(doc)
    folders = [{"model": "AirEngine8776-X7-T", "n": 3, "total": 3, "name": "3-3", "imgs": ["x"]},
               {"model": "S6750-H48Y8C", "n": 1, "total": 1, "name": "1-1", "imgs": ["y"]},
               {"model": "Unknown", "n": 1, "total": 1, "name": "1-1", "imgs": ["z"]}]
    logs = []
    m = app.map_set_folders_to_tables(folders, entries, pn, log_cb=logs.append)
    by_sn = {v["entry"]["sn"]: v["imgs"] for v in m.values()}
    assert by_sn == {"AP0000000003": ["x"], "CORE00000001": ["y"]}
    # 'Unknown' x/1 fits two 1-unit groups (Core, Mgmt) and a 1-unit AP section: skipped, not guessed
    assert any("Unknown" in line and "SKIP" in line for line in logs)


def test_run_injection_set_fallback_and_sn_priority(tmp_path):
    p = tmp_path / "d.docx"
    _set_doc(str(p))
    root = tmp_path / "photos"
    _make_photo(str(root / "AirEngine8776-X7-T" / "By_Number" / "1-3"))
    _make_photo(str(root / "AirEngine8776-X7-T" / "By_Number" / "3-3"))
    _make_photo(str(root / "AirEngine8776-X7-T" / "By_Number" / "3-3"), "b.jpg")
    _make_photo(str(root / "AirEngine8776-X7-T" / "By_SN" / "AP0000000001"))  # S/N folder wins for 1-3
    out = tmp_path / "o.docx"
    logs = []
    res = app.run_injection(str(p), str(root), str(out), auto_match=True, set_number_fallback=True,
                            auto_rotate=False, smart_rotate=False, log_callback=logs.append)
    assert res.injected == 2 and res.img_count == 3
    d = Document(str(out))
    assert _count_pictures(d.tables[1]) == 1   # AP 1/2 from its S/N folder
    assert _count_pictures(d.tables[2]) == 0   # AP 2/2 has no folder
    assert _count_pictures(d.tables[3]) == 2   # AP 1/1 = set 3 of 3
    assert any("1 already matched by S/N" in line for line in logs)

    # Off by default
    res2 = app.run_injection(str(p), str(root), str(tmp_path / "o2.docx"), auto_match=True,
                             auto_rotate=False, smart_rotate=False)
    assert res2.injected == 1


def test_assign_unread_far_from_any_label_stays_unassigned():
    """Mixed devices: other devices' photos (labels unreadable) shot minutes away are not pulled in."""
    seq = [("A", "A"), ("a1", None), ("s1", None), ("s2", None), ("B", "B"), ("x1", None), ("x2", None)]
    times = [0, 3, 300, 305, 600, 740, 745]
    groups, _ = app.assign_unread_photos(seq, times)
    assert groups["A"] == ["A", "a1"]
    assert groups["B"] == ["B"]
    assert groups[None] == ["s1", "s2", "x1", "x2"]


def test_assign_unread_count_mode_caps_at_usual_count():
    seq = [("A", "A"), ("a1", None), ("B", "B"), ("b1", None), ("C", "C"), ("c1", None), ("j1", None), ("j2", None)]
    groups, _ = app.assign_unread_photos(seq)  # typical = round(8/3) = 3
    assert groups["C"] == ["C", "c1", "j1"]
    assert groups[None] == ["j2"]


# ── Set-number sticker on the box ('31/100') ──

def _sticker_entries():
    doc = Document()
    _summary_table(doc, [("152502547", "TPLINK ARCHER-TBE400E"), ("02355UEL", "S6750"), ("02354FRN", "AirEngine9700")])
    for n in (1, 2, 3):
        _detail_table(doc, "USB Wi-Fi", n, 3, "152502547", f"33770F200104{n}")
    _detail_table(doc, "Core", 1, 2, "02355UEL", "CORE00000001")
    _detail_table(doc, "Core", 2, 2, "02355UEL", "CORE00000002")
    _detail_table(doc, "WLC", 1, 2, "02354FRN", "WLC000000001")
    _detail_table(doc, "WLC", 2, 2, "02354FRN", "WLC000000002")
    entries = app.extract_set_entries(doc)
    ser = app.extract_doc_serials(doc)
    return app.build_set_sticker_lookup(entries), ser


def test_sticker_maps_to_document_serial():
    lk, ser = _sticker_entries()
    assert app.serial_from_set_sticker(["2/3"], lk, ser) == ("33770F2001042", "2/3")
    assert app.serial_from_set_sticker(["2/3"], lk, ser)[0] == "33770F2001042"
    # OCR look-alikes inside a sticker
    assert app.serial_from_set_sticker(["3/ด"], lk, ser) == (None, "")   # 3/0 is not a valid set
    assert app.serial_from_set_sticker(["l/3"], lk, ser)[0] == "33770F2001041"


def test_sticker_ambiguous_or_embedded_is_ignored():
    lk, ser = _sticker_entries()
    assert app.serial_from_set_sticker(["1/2"], lk, ser) == (None, "")          # Core 1/2 or WLC 1/2?
    assert app.serial_from_set_sticker(["2880/2880 Mbps tri-band speed 2/3"], lk, ser) == (None, "")
    assert app.serial_from_set_sticker(["4/3"], lk, ser) == (None, "")          # n > total


class _StickerLine:
    def __init__(self, text, box):
        self.text, self.box, self.dropped, self.confidence = text, box, False, 0.9


def test_foreground_sticker_beats_background_box():
    """'71/280' on the box in front, '(271/280' on a box behind near the photo edge."""
    from PIL import Image as PILImage
    img = PILImage.new("RGB", (4096, 3072), "white")

    class _R:
        lines = [_StickerLine("(271/280", [[3704, 586], [4090, 586], [4090, 801], [3704, 801]]),
                 _StickerLine("71/280", [[2010, 2032], [2400, 2032], [2400, 2222], [2010, 2222]])]
    assert app._foreground_sticker_line(img, _R()).text == "71/280"


def test_sticker_read_dropped_when_nearby_label_disagrees(tmp_path):
    ser = {"AAA1234567": "AAA1234567", "BBB7654321": "BBB7654321"}
    d = str(tmp_path / "all")
    imgs = [_photo(os.path.join(d, n), (i * 40, 0, 0)) for i, n in enumerate(["1.png", "2.png", "3.png"])]
    reads = {imgs[0]: ("AAA1234567", "exact"), imgs[1]: ("BBB7654321", "sticker 9/9+zoom"),
             imgs[2]: ("BBB7654321", "exact")}
    times = {imgs[0]: 0, imgs[1]: 3, imgs[2]: 300}
    orig = app.read_serial_from_photo
    app.read_serial_from_photo = lambda p, e, s, **k: reads[p]
    try:
        groups, un = app.group_photos_by_ocr(imgs, ser, engine=object(), time_fn=lambda p: times[p],
                                             sticker_lookup={9: []})
    finally:
        app.read_serial_from_photo = orig
    # photo 2's sticker disagrees with photo 1's label shot 3 s earlier -> ignored, joins A by time
    assert [os.path.basename(p) for p in groups["AAA1234567"]] == ["1.png", "2.png"]
    assert [os.path.basename(p) for p in groups["BBB7654321"]] == ["3.png"]


# ── Run summary report + duplicate S/N in the document ──

def _dup_doc(path):
    """3 detail tables; the document wrongly repeats set 1's S/N in set 3."""
    doc = Document()
    _summary_table(doc, [("PN-100", "Switch-X")])
    _detail_table(doc, "Switch", 1, 3, "PN-100", "SWX00000001")
    _detail_table(doc, "Switch", 2, 3, "PN-100", "SWX00000002")
    _detail_table(doc, "Switch", 3, 3, "PN-100", "SWX00000001")   # document error (should be ...03)
    doc.save(path)


def test_duplicate_document_serial_uses_set_folder_and_is_reported(tmp_path):
    p = tmp_path / "d.docx"
    _dup_doc(str(p))
    root = tmp_path / "photos"
    _make_photo(str(root / "Switch-X" / "By_SN" / "SWX00000001"), "a.jpg")
    _make_photo(str(root / "Switch-X" / "By_SN" / "SWX00000001"), "b.jpg")
    _make_photo(str(root / "Switch-X" / "By_SN" / "SWX00000003"), "c.jpg")      # not in the document
    os.makedirs(str(root / "Switch-X" / "By_SN" / "SWX00000002"))               # empty folder
    _make_photo(str(root / "Switch-X" / "By_Number" / "3-3"), "c.jpg")
    out = tmp_path / "o.docx"
    logs = []
    res = app.run_injection(str(p), str(root), str(out), auto_match=True, set_number_fallback=True,
                            auto_rotate=False, smart_rotate=False, log_callback=logs.append)
    d = Document(str(out))
    assert _count_pictures(d.tables[1]) == 2      # set 1: its S/N folder
    assert _count_pictures(d.tables[2]) == 0      # set 2: empty folder
    assert _count_pictures(d.tables[3]) == 1      # set 3: NOT set 1's photos again -> set folder 3-3
    assert any("same S/N in tables" in line for line in logs)

    text = open(res["summary_path"], encoding="utf-8-sig").read()
    assert "Switch-X (Part Number PN-100) (Total: 3 Units)" in text
    assert "2/3" in text and "photo folder is empty (0 files)" in text
    assert "same S/N in more than one table" in text
    assert "SWX00000003" in text                   # folder not found in document


def test_summary_written_when_nothing_injected(tmp_path):
    p = tmp_path / "d.docx"
    _dup_doc(str(p))
    root = tmp_path / "photos"
    _make_photo(str(root / "Other" / "ZZZ99999999"))
    res = app.run_injection(str(p), str(root), str(tmp_path / "o.docx"), auto_match=True,
                            auto_rotate=False, smart_rotate=False)
    assert res.injected == 0
    text = open(res["summary_path"], encoding="utf-8-sig").read()
    assert "NOT saved" in text and "Missing photos: 3 unit(s)" in text


def test_find_empty_folders(tmp_path):
    _make_photo(str(tmp_path / "A" / "SN1"))
    os.makedirs(str(tmp_path / "A" / "SN2"))
    os.makedirs(str(tmp_path / "B" / "SN3" / "DCIM"))
    assert app.find_empty_folders(str(tmp_path)) == ["B", "DCIM", "SN2", "SN3"]
