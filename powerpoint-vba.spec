# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['C:\\_daten.lokal\\_workarea\\git.repos\\onderhold\\vba-editRD\\src\\vba_edit\\powerpoint_vba.py'],
    pathex=['C:\\_daten.lokal\\_workarea\\git.repos\\onderhold\\vba-editRD\\src\\vba_edit'],
    binaries=[],
    datas=[],
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
    name='powerpoint-vba',
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
