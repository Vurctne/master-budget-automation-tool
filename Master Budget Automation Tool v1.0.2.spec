# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

from app_metadata import APP_TITLE


PROJECT_ROOT = Path(globals().get('SPECPATH', Path.cwd()))
VERSION_FILE = PROJECT_ROOT / 'windows_version_info.txt'

a = Analysis(
    ['app.py'],
    pathex=[str(PROJECT_ROOT)],
    binaries=[],
    datas=[
        (str(PROJECT_ROOT / 'assets' / 'compass_step1.png'), 'assets'),
        (str(PROJECT_ROOT / 'assets' / 'compass_step2.png'), 'assets'),
    ],
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
    name=APP_TITLE,
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
    version=str(VERSION_FILE),
    codesign_identity=None,
    entitlements_file=None,
)
