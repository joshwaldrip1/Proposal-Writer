# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['proposal_gui.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('PROPOSAL FOR PROFESSIONAL SERVICES.dotx', '.'),
        ('Proposal Creator Key.xlsx', '.'),
        ('Civil Engineering Services.xlsx', '.'),
        ('Monday Export.xlsx', '.'),
    ],
    hiddenimports=['win32com', 'win32com.client', 'pythoncom'],
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
    name='proposal_gui',
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
