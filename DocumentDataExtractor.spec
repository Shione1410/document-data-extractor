# -*- mode: python ; coding: utf-8 -*-

"""
PyInstaller build specification for Document Data Extractor.

This build intentionally does NOT use collect_all("paddle").
Collecting the full Paddle package can include very deep C++ header paths and
cause COLLECT failures on Windows.

PaddleX checks installed dependency metadata when creating the OCR pipeline.
Therefore the metadata of installed PaddleX dependency packages is copied into
the bundled application.
"""

from importlib import metadata

import paddlex
from PyInstaller.utils.hooks import (
    collect_data_files,
    collect_dynamic_libs,
    copy_metadata,
)


def safe_copy_metadata(package_name):
    try:
        return copy_metadata(package_name)
    except Exception:
        return []


datas = [
    ("OCR.yaml", "."),
]

# PaddleX pipeline configuration/data files.
datas += collect_data_files("paddlex")

# PaddleX itself and PaddleOCR metadata.
datas += safe_copy_metadata("paddlex")
datas += safe_copy_metadata("paddleocr")

# PaddleX checks dependency availability through installed distribution
# metadata. Copy metadata for every installed dependency known to PaddleX.
installed_names = {
    dist.metadata.get("Name")
    for dist in metadata.distributions()
    if dist.metadata.get("Name")
}

paddlex_dependency_names = set(paddlex.utils.deps.BASE_DEP_SPECS.keys())

for package_name in sorted(installed_names & paddlex_dependency_names):
    datas += safe_copy_metadata(package_name)

# Paddle runtime DLLs / native binaries only.
# Do not collect all Paddle data, because include/ headers are unnecessary.
binaries = collect_dynamic_libs("paddle")


a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="DocumentDataExtractor",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="DocumentDataExtractor",
)
