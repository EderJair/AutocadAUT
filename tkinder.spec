# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['tkinder.py'],
    pathex=[],
    binaries=[],
    datas=[('script.py', '.'), ('CONVERTIDOR.xlsx', '.'), ('PLANO1.dxf', '.')],
    hiddenimports=['ezdxf', 'xlwings', 'shapely'],
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
    name='tkinder',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
