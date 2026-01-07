# -*- mode: python ; coding: utf-8 -*-

import os

extra_datas = []

# Inclui SumatraPDF portátil, se presente
sumatra_candidates = [
    'SumatraPDF.exe',
    os.path.join('vendor', 'SumatraPDF.exe'),
    os.path.join('vendor', 'sumatra', 'SumatraPDF.exe'),
]
for cand in sumatra_candidates:
    if os.path.isfile(cand):
        # Coloca na raiz do dist para facilitar a descoberta
        extra_datas.append((cand, '.'))
        break


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=extra_datas,
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
    name='main',
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
    icon=['gac-logo.ico'],
)
