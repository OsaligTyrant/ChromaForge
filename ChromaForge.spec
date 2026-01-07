# -*- mode: python ; coding: utf-8 -*-

import os
import tkinterdnd2

block_cipher = None

tkdnd_dir = os.path.join(os.path.dirname(tkinterdnd2.__file__), "tkdnd")

a = Analysis(
    ["app\\png_transparency_gui.py"],
    pathex=[],
    binaries=[],
    datas=[
        ("app\\ChromaForge_logo.png", "."),
        ("app\\ChromaForge_logo.ico", "."),
        (tkdnd_dir, "tkinterdnd2\\tkdnd"),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="ChromaForge",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    icon="app\\ChromaForge_logo.ico",
)
