# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(['main_platform.py'],
             pathex=['d:\Coding\Project\PreciTestSystem\PTS'],
             binaries=[],
             datas=[('PreciLasers.ico', '.')],
             hiddenimports=['pyvisa', 'matplotlib.backends.backend_tkagg', 'PIL._tkinter_finder'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='PTS',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False,
          icon='PreciLasers.ico',
          disable_windowed_traceback=False,
          target_arch=None,
          codesign_identity=None,
          entitlements_file=None)