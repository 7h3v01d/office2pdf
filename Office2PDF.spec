# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['gui.py'],
    pathex=['G:\\Projects\\_Office\\office2pdf'],
    binaries=[],
    datas=[('G:\\Projects\\_Office\\office2pdf\\assets\\office2pdf.ico', 'assets'), ('G:\\Projects\\_Office\\office2pdf\\LICENSE.txt', '.'), ('G:\\Projects\\_Office\\office2pdf\\THIRD_PARTY_NOTICES.txt', '.'), ('G:\\Projects\\_Office\\office2pdf\\SOURCE_OFFER.txt', '.')],
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
    [],
    exclude_binaries=True,
    name='Office2PDF',
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
    version='G:\\Projects\\_Office\\office2pdf\\windows_version_info.txt',
    icon=['G:\\Projects\\_Office\\office2pdf\\assets\\office2pdf.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Office2PDF',
)
