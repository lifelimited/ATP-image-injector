"""Time-left estimate (eta.py) with a fake clock."""
import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
import eta  # noqa: E402


class Clock:
    def __init__(self):
        self.t = 0.0

    def __call__(self):
        return self.t


def test_format_duration():
    assert eta.format_duration(3) == "5 s"
    assert eta.format_duration(42) == "45 s"
    assert eta.format_duration(95) == "1 min 30 s"
    assert eta.format_duration(62) == "1 min"
    assert eta.format_duration(12 * 60 + 40) == "13 min"
    assert eta.format_duration(2 * 3600 + 5 * 60) == "2 h 05 min"


def test_estimate_uses_measured_speed_and_future_phases():
    c = Clock()
    t = eta.EtaTracker(clock=c, min_interval=0)
    t.rates.update({"ocr": 1.0, "insert": 0.5, "save": 0.1})
    t.set_plan("insert", 100)
    t.set_plan("save", 100)
    t.start("ocr", 100)
    assert abs(t.remaining() - (100 * 1.0 + 100 * 0.5 + 100 * 0.1)) < 1e-6    # nothing measured yet
    c.t = 60.0
    t.advance("ocr", 30)                       # 2 s per photo, past the trust fraction
    assert abs(t.remaining() - (70 * 2.0 + 50 + 10)) < 1e-6
    t.finish("ocr")
    t.start("insert", 80)                      # real count replaces the plan
    assert abs(t.remaining() - (80 * 0.5 + 100 * 0.1)) < 1e-6


def test_measured_speed_blends_in_early():
    c = Clock()
    t = eta.EtaTracker(clock=c, min_interval=0)
    t.rates["ocr"] = 1.0
    t.start("ocr", 100)
    c.t = 15.0
    t.advance("ocr", 5)                        # 3 s/photo measured on 5% of the phase
    rem = t.remaining()
    assert 95 * 1.0 < rem < 95 * 3.0           # between the prior and the measured speed


def test_learned_speed_is_saved_and_reloaded(tmp_path):
    path = str(tmp_path / "stats.json")
    c = Clock()
    t = eta.EtaTracker(clock=c, stats_path=path, variant={"insert": "smart"})
    t.start("insert", 10)
    c.t = 10.0
    t.advance("insert", 10)
    t.finish("insert")
    t.close()
    rates = json.load(open(path))["rates"]
    expected = (1 - eta.LEARN_WEIGHT) * eta.DEFAULT_RATES["insert:smart"] + eta.LEARN_WEIGHT * 1.0
    assert abs(rates["insert:smart"] - expected) < 1e-9
    t2 = eta.EtaTracker(stats_path=path, variant={"insert": "smart"})
    assert abs(t2.rate("insert") - expected) < 1e-9


def test_status_text_and_callback():
    c = Clock()
    seen = []
    t = eta.EtaTracker(clock=c, callback=lambda text, frac: seen.append((text, frac)), min_interval=0)
    t.start("ocr", 50)
    assert "estimating" in seen[-1][0]
    c.t = 20.0
    t.advance("ocr", 10)
    text, frac = seen[-1]
    assert text.startswith("Reading S/N from photos... 10/50") and "left" in text
    assert 0.0 < frac < 1.0


def test_module_helpers_are_noops_without_active_tracker():
    eta.set_active(None)
    eta.start("ocr", 5)
    eta.advance("ocr")
    eta.finish("ocr")
    eta.plan("insert", 3)


def test_run_injection_reports_eta(tmp_path):
    import app
    from docx import Document
    from PIL import Image
    doc = Document()
    for sn in ("AAA1234567", "BBB7654321"):
        tb = doc.add_table(rows=3, cols=2)
        tb.rows[0].cells[0].text = "Serial Number:"
        tb.rows[0].cells[1].text = sn
        row = tb.rows[2]
        row.cells[0].merge(row.cells[1])
        os.makedirs(tmp_path / "p" / sn)
        Image.new("RGB", (40, 30), (200, 100, 50)).save(str(tmp_path / "p" / sn / "a.jpg"))
    doc.save(str(tmp_path / "d.docx"))
    seen = []
    res = app.run_injection(str(tmp_path / "d.docx"), str(tmp_path / "p"), str(tmp_path / "o.docx"),
                            auto_match=True, auto_rotate=False, smart_rotate=False,
                            eta_callback=lambda text, frac: seen.append(text),
                            eta_stats_path=str(tmp_path / "s.json"))
    assert res.injected == 2
    assert any(t.startswith("Inserting photos") for t in seen)
    assert app.run_injection.last_tracker.finished.keys() >= {"insert", "save"}
    assert os.path.exists(tmp_path / "s.json")
