# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['Pending\\main_v2.py'],
    pathex=[],
    binaries=[],
    datas=[('Header.jpg', '.'), ('Fonts\\Armata-Regular.ttf', 'Fonts'), ('Fonts\\Novecentowide-Bold.ttf', 'Fonts'), ('Fonts\\Novecentowide-DemiBold_0.ttf', 'Fonts'), ('favicon.ico', '.')],
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
    name='Pending',
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
)
