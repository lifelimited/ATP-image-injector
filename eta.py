"""
Time-left estimate for a run of the injector.

A run goes through phases whose sizes are known (or estimated) up front:
    ocr     - Read S/N from photos (photos)
    verify  - Verify photos OCR (photos)
    insert  - prepare + insert photos into Word (photos)
    save    - write the .docx (photos inserted)

Remaining time = remaining work of the current phase / its measured speed, blended with the
learned speed while only a little of the phase is done, plus the planned work of the phases still
to come times their learned seconds per unit. Learned speeds come from earlier runs on this
computer (a small JSON file) or built-in defaults, and are updated at the end of each run.
"""
from __future__ import annotations

import json
import os
import threading
import time

PHASES = ("ocr", "verify", "insert", "save", "summary")
LABELS = {"ocr": "Reading S/N from photos", "verify": "Verify photos",
          "insert": "Inserting photos", "save": "Saving Word file",
          "summary": "Writing summary"}

# Seconds per unit measured on a 16-thread desktop CPU (1.6.1, real HEIC/JPG photos); updated per run.
DEFAULT_RATES = {
    "ocr": 0.40,
    "verify": 0.35,
    "insert:smart": 0.25,
    "insert:plain": 0.05,
    "insert:orig": 0.04,
    "save": 0.02,
    "summary": 1.5,          # seconds per run (summary report and clean-up)
}
LEARN_WEIGHT = 0.5            # new run's speed vs the stored one
TRUST_FRACTION = 0.25         # measured speed fully trusted once this share of a phase is done
MIN_UNITS_FOR_SPEED = 3       # below this many done units the measured speed is ignored
RECENT_SHARE = 0.5            # speed is measured over the most recent half of the work done so far,
                              # so start-up costs (OCR engines loading per worker) do not stay in it


def default_stats_path() -> str:
    base = os.environ.get("UTII_STATS_DIR") or os.path.join(
        os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "UnitTestImageInjector")
    return os.path.join(base, "eta_stats.json")


def format_duration(sec: float) -> str:
    sec = max(0, int(round(sec)))
    if sec < 60:
        return f"{max(5, (sec + 4) // 5 * 5)} s"
    if sec < 3600:
        m, s = divmod(sec, 60)
        if m >= 10:
            return f"{m + (1 if s >= 30 else 0)} min"
        s = s // 10 * 10
        return f"{m} min {s} s" if s else f"{m} min"
    h, m = divmod((sec + 30) // 60, 60)
    return f"{h} h {m:02d} min"


class EtaTracker:
    def __init__(self, callback=None, stats_path: str | None = None, clock=time.monotonic,
                 variant: dict | None = None, min_interval: float = 1.0):
        self.callback = callback
        self.stats_path = stats_path
        self.clock = clock
        self.variant = dict(variant or {})          # e.g. {"insert": "smart"}
        self.min_interval = min_interval
        self.lock = threading.Lock()
        self.rates = dict(DEFAULT_RATES)
        self.rates.update(self._load())
        self.plan: dict[str, float] = {}            # phase -> planned units
        self.done: dict[str, float] = {}
        self.started: dict[str, float] = {}
        self.history: dict[str, list] = {}         # phase -> [(time, done units)]
        self.finished: dict[str, float] = {}        # phase -> seconds taken
        self.current: str | None = None
        self.t0 = clock()
        self.samples: list[tuple[float, float]] = []  # (elapsed, estimated remaining)
        self._last_emit = -1e9

    # ── learned speeds ──
    def _load(self) -> dict:
        if not self.stats_path:
            return {}
        try:
            with open(self.stats_path, encoding="utf-8") as f:
                data = json.load(f)
            return {k: float(v) for k, v in data.get("rates", {}).items() if float(v) > 0}
        except Exception:
            return {}

    def _save(self):
        if not self.stats_path:
            return
        try:
            os.makedirs(os.path.dirname(self.stats_path), exist_ok=True)
            tmp = self.stats_path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump({"rates": self.rates}, f, indent=1)
            os.replace(tmp, self.stats_path)
        except Exception:
            pass

    def _key(self, phase: str) -> str:
        v = self.variant.get(phase)
        return f"{phase}:{v}" if v else phase

    def rate(self, phase: str) -> float:
        return self.rates.get(self._key(phase), DEFAULT_RATES.get(self._key(phase), 0.3))

    # ── phases ──
    def set_plan(self, phase: str, units: float):
        with self.lock:
            if phase not in self.finished:
                self.plan[phase] = max(0.0, float(units))

    def start(self, phase: str, units: float | None = None):
        with self.lock:
            if units is not None:
                self.plan[phase] = max(0.0, float(units))
            self.current = phase
            self.started[phase] = self.clock()
            self.done[phase] = 0.0
            self.history[phase] = [(self.started[phase], 0.0)]
        self._emit(force=True)

    def advance(self, phase: str, n: float = 1):
        with self.lock:
            if phase not in self.started or phase in self.finished:
                return
            self.done[phase] = self.done.get(phase, 0.0) + n
            self.history.setdefault(phase, []).append((self.clock(), self.done[phase]))
        self._emit()

    def finish(self, phase: str):
        with self.lock:
            if phase not in self.started or phase in self.finished:
                return
            took = self.clock() - self.started[phase]
            self.finished[phase] = took
            units = self.done.get(phase, 0.0) or self.plan.get(phase, 0.0)
            if (units >= MIN_UNITS_FOR_SPEED or phase == "summary") and took > 0:
                key = self._key(phase)
                measured = took / units
                old = self.rates.get(key, DEFAULT_RATES.get(key, measured))
                self.rates[key] = (1 - LEARN_WEIGHT) * old + LEARN_WEIGHT * measured
            if self.current == phase:
                self.current = None
        self._emit(force=True)

    def skip(self, phase: str):
        with self.lock:
            self.plan.pop(phase, None)

    def close(self):
        """End of run: finish an open step and store the learned speeds."""
        if "summary" in self.started:
            self.finish("summary")
        self._save()

    # ── estimate ──
    def remaining(self) -> float | None:
        with self.lock:
            now = self.clock()
            total = 0.0
            for ph in PHASES:
                if ph in self.finished or ph not in self.plan:
                    continue
                units = self.plan[ph]
                prior = self.rate(ph)
                if ph in self.started:
                    done = min(self.done.get(ph, 0.0), units) if units else self.done.get(ph, 0.0)
                    left = max(0.0, units - done)
                    elapsed = now - self.started[ph]
                    if done >= MIN_UNITS_FOR_SPEED and elapsed > 0:
                        measured = self._recent_speed(ph, now, done, elapsed)
                        frac = done / units if units else 1.0
                        w = min(1.0, frac / TRUST_FRACTION)
                        sec_per = w * measured + (1 - w) * prior
                    else:
                        sec_per = prior
                        left = max(0.0, units * prior - elapsed) / prior if prior else left
                    total += left * sec_per
                else:
                    total += units * prior
            return total

    def _recent_speed(self, ph, now, done, elapsed) -> float:
        """Seconds per unit over the most recent part of the phase (lock held)."""
        hist = self.history.get(ph) or []
        target = done * (1 - RECENT_SHARE)
        ref_t, ref_done = self.started[ph], 0.0
        for t, d in hist:
            if d <= target:
                ref_t, ref_done = t, d
            else:
                break
        if done - ref_done >= MIN_UNITS_FOR_SPEED and now > ref_t:
            return (now - ref_t) / (done - ref_done)
        return elapsed / done

    def fraction(self) -> float:
        rem = self.remaining() or 0.0
        el = self.clock() - self.t0
        return el / (el + rem) if (el + rem) > 0 else 0.0

    def status(self) -> str:
        with self.lock:
            ph = self.current
            done = self.done.get(ph, 0.0) if ph else 0.0
            units = self.plan.get(ph, 0.0) if ph else 0.0
            started = self.started.get(ph) if ph else None
        rem = self.remaining()
        label = LABELS.get(ph, "Working") if ph else "Working"
        count = f" {int(done)}/{int(units)}" if ph and units and ph != "save" else ""
        early = ph and ph != "save" and done < MIN_UNITS_FOR_SPEED and (self.clock() - (started or 0)) < 5
        if rem is None or early:
            return f"{label}...{count} · estimating time left"
        return f"{label}...{count} · about {format_duration(rem)} left"

    def _emit(self, force: bool = False):
        now = self.clock()
        if not force and now - self._last_emit < self.min_interval:
            return
        self._last_emit = now
        rem = self.remaining()
        if rem is not None:
            self.samples.append((now - self.t0, rem))
        if self.callback:
            try:
                self.callback(self.status(), self.fraction())
            except Exception:
                pass

    def accuracy(self, total_seconds: float | None = None, points=(0.1, 0.3, 0.5, 0.7, 0.9)) -> dict:
        """Error of the estimate at given shares of the run: {share: (estimated, actual, error %)}."""
        total = total_seconds if total_seconds is not None else self.clock() - self.t0
        out = {}
        for p in points:
            t = p * total
            near = [s for s in self.samples if s[0] <= t]
            if not near:
                continue
            el, est = near[-1]
            actual = total - el
            out[p] = (est, actual, (est - actual) / total * 100 if total else 0.0)
        return out


_ACTIVE: EtaTracker | None = None


def set_active(tracker: EtaTracker | None):
    global _ACTIVE
    _ACTIVE = tracker


def advance(phase: str, n: float = 1):
    t = _ACTIVE
    if t is not None:
        t.advance(phase, n)


def plan(phase: str, units: float):
    t = _ACTIVE
    if t is not None:
        t.set_plan(phase, units)


def start(phase: str, units: float | None = None):
    t = _ACTIVE
    if t is not None:
        t.start(phase, units)


def finish(phase: str):
    t = _ACTIVE
    if t is not None:
        t.finish(phase)
