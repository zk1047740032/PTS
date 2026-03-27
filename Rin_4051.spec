# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['D:\\Coding\\Project\\PreciTestSystem\\PTS\\独立程序\\Rin_4051.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['pyvisa', 'tkinter', 'PIL', 'matplotlib'],
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
    name='Rin_4051',
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
    icon=['D:\\Coding\\Project\\PreciTestSystem\\PTS\\PreciLasers.ico'],
)
