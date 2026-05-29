# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for Manila Polo Club — Instructor Fee Summary
Produces a single-folder distribution in dist\MPCScript\
"""

import sys
from pathlib import Path

ROOT = Path(SPECPATH)   # project root (where this .spec lives)

block_cipher = None

a = Analysis(
    [str(ROOT / 'main.py')],
    pathex=[str(ROOT)],
    binaries=[],
    datas=[
        # Bundle the logo PNG
        (str(ROOT / 'ui' / 'src' / 'MPC-LOGO-white-text-1.png'), 'ui/src'),
    ],
    hiddenimports=[
        # pandas/numpy internals
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.skiplist',
        # openpyxl
        'openpyxl.cell._writer',
        # xlrd
        'xlrd',
        # reportlab
        'reportlab.graphics.barcode.common',
        'reportlab.graphics.barcode.code128',
        # Pillow
        'PIL._tkinter_finder',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'scipy', 'IPython', 'notebook',
        'pygments', 'docutils', 'sphinx',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='MPCScript',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,          # no console window (GUI app)
    icon=str(ROOT / 'ui' / 'src' / 'mpc_icon.ico'),
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='MPCScript',
)
