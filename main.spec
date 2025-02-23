# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
plist = {
    'CFBundleName': 'AI勘误器_Word版',
    'CFBundleDisplayName': 'AI勘误器_Word版',
    'CFBundleGetInfoString': 'AI勘误器_Word版',
    'CFBundleIdentifier': 'com.errata.word',
    'CFBundleVersion': '1.0.0',
    'CFBundleShortVersionString': '1.0.0',
    'NSHighResolutionCapable': 'True'
}

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='AI勘误器_Word版',
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
    icon='errata.icns',
)

app = BUNDLE(
    exe,
    name='AI勘误器_Word版.app',
    icon='errata.icns',
    bundle_identifier='com.errata.word',
    info_plist=plist,
)