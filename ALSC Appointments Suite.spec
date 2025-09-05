# -*- mode: python ; coding: utf-8 -*-
import glob, os
datas = [
    ('Header.jpg', '.'),  # copy to app root
]
datas += [(f, os.path.join('Fonts', os.path.basename(f))) for f in glob.glob('Fonts\\*.ttf')]
a = Analysis(
    ['Payments\\main_v2.py'],
    pathex=[],
    binaries=[],
    datas=datas,                # ← here
    hiddenimports=[],
    hookspath=[],
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
    name='ALSC Appointments Suite',
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
    icon=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='ALSC Appointments Suite',
)
