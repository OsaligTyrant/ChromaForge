# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['app\\png_transparency_gui.py'],
    pathex=[],
    binaries=[],
    datas=[('app\\ChromaForge_logo.png', '.'), ('app\\ChromaForge_logo.ico', '.'), ('C:\\Users\\TA-Ko\\AppData\\Local\\Packages\\PythonSoftwareFoundation.Python.3.11_qbz5n2kfra8p0\\LocalCache\\local-packages\\Python311\\site-packages\\tkinterdnd2\\tkdnd', 'tkinterdnd2\\tkdnd')],
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
    a.binaries,
    a.datas,
    [],
    name='ChromaForge',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['app\\ChromaForge_logo.ico'],
)
