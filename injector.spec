# PyInstaller build of the Windows exe:   python -m PyInstaller --noconfirm injector.spec
# One-file exe by default; set BUILD_ONEDIR=1 for a folder build (starts instantly, ships as a zip).
#
# Why a spec file and not the old one-line command:
# - The exe only bundles what the app uses. PyInstaller pulls in everything installed that anything might
#   import (PySide6 45 MB, jedi, hf_xet, OpenCV's video DLL ...): the one-file exe unpacked 7,574 files /
#   365 MB on every start (6.6 s before the window could open). Now ~1,260 files / ~240 MB (~2.5 s).
# - The Microsoft C++ runtime is taken from Windows (System32). PyInstaller otherwise bundles whichever
#   msvcp140.dll it meets first, e.g. 14.36, and onnxruntime 1.27 then fails with "DLL initialization
#   routine failed" (OCR unavailable). Builds before 1.7.0 only worked because PySide6 brought a newer one.
import os
import re

_here = os.path.abspath(SPECPATH)
_version = re.search(r'^__version__ = "([^"]+)"', open(os.path.join(_here, "app.py"), encoding="utf-8").read(),
                     re.M).group(1)
_name = f"Unit_Test_Image_Injector_v{_version}"
_onedir = os.environ.get("BUILD_ONEDIR") == "1"
_model = os.path.join("models", "OpenThai2.0-OCR-Text")

a = Analysis(
    [os.path.join(_here, "app.py")],
    pathex=[_here],
    datas=[(os.path.join(_here, _model), _model)],
    hiddenimports=["onnxruntime", "pyclipper", "cv2", "unicodedata"],
    excludes=["torch", "torchvision", "easyocr", "scipy", "matplotlib", "pandas", "pytest",
              "PySide6", "shiboken6", "PyQt5", "PyQt6", "jedi", "parso", "IPython",
              "hf_xet", "pydantic", "pydantic_core", "cryptography"],
)

_RUNTIME = ("msvcp140.dll", "msvcp140_1.dll", "msvcp140_2.dll", "msvcp140_atomic_wait.dll",
            "msvcp140_codecvt_ids.dll", "vcruntime140.dll", "vcruntime140_1.dll", "concrt140.dll")
_sys32 = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), "System32")
a.binaries = [b for b in a.binaries
              if os.path.basename(b[0]).lower() not in _RUNTIME and "opencv_videoio_ffmpeg" not in b[0].lower()]
a.binaries += [(n, os.path.join(_sys32, n), "BINARY") for n in _RUNTIME if os.path.exists(os.path.join(_sys32, n))]

pyz = PYZ(a.pure)
if _onedir:
    exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name=_name, console=False, upx=False)
    coll = COLLECT(exe, a.binaries, a.datas, name=_name, upx=False)
else:
    exe = EXE(pyz, a.scripts, a.binaries, a.datas, [], name=_name, console=False, upx=False,
              runtime_tmpdir=None)
