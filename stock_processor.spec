# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['stock_processor.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('data/stock_info.json', 'data'),  # 如果需要包含資料檔案
    ],
    hiddenimports=[
        'pandas',
        'openpyxl',
        'FinMind',
        'FinMind.data',
        'numpy',
        'pytz',
        'dateutil',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='股票數據處理器',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # 保持命令列視窗，可以看到處理進度
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # 如果有圖示可以指定
)
