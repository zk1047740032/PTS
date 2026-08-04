# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(['main_platform.py'],
             pathex=[r'd:\Coding\Project\PTS\zhongzi'],
             binaries=[],
             datas=[
    ('PreciLasers.ico', '.'),
    ('report/templates/template_default.docx', 'report/templates'),
    ('docs/一键测试操作说明.md', 'docs'),
    ('help_images', 'help_images'),
],
             hiddenimports=[
        # 仪器/UI 底层依赖
        'pyvisa', 'pyvisa_py', 'pyvisa_py.tcpip', 'pyvisa_py.usb',
        'serial', 'pywinauto',
        # matplotlib / PIL 与 tkinter 集成
        'matplotlib.backends.backend_tkagg', 'PIL._tkinter_finder',
        # docx 报告生成
        'docxtpl', 'docx', 'jinja2', 'lxml',
        # multiprocessing (Windows + console=False 场景)
        'multiprocessing',
        # === 通过 importlib.import_module 动态加载的测试模块 ===
        'path_a.Rin_FSV3004',
        'path_a.LineWidth_FSV3004',
        'path_a.TimeDomain',
        'path_a.SpectrumSNR',
        'path_a.SingleFrequency',
        'path_b.Power',
        'path_a.PhaseNoise',
        'path_a.WaveLength',
        # path_a.WaveLength 的传递依赖
        'path_a.drivers.wlmData',
        'path_a.drivers.wlmConst',
    ],
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