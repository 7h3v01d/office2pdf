# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['native_office_worker.py'],
    pathex=['G:\\Projects\\_Office\\office2pdf'],
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
    name='Office2PDFNativeWorker',
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
    version='G:\\Projects\\_Office\\office2pdf\\windows_worker_version_info.txt',
    icon=['G:\\Projects\\_Office\\office2pdf\\assets\\office2pdf.ico'],
)
