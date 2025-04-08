# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['tkinder.py'],
    pathex=[],
    binaries=[],
    datas=[('CONVERTIDOR.xlsx', '.'), ('script.py', '.')],
    hiddenimports=['ezdxf', 'ezdxf.addons', 'shapely', 'shapely.geometry', 'xlwings', 'PIL'],
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
    name='ACERO_SCRIPT',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # Cambia a False para la versión final
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)