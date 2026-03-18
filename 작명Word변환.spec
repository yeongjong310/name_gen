# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['src/name_gen/main.py'],
    pathex=[],
    binaries=[],
    datas=[('src/name_gen/assets/happy_numbers.xlsx', 'name_gen/assets')],
    hiddenimports=['name_gen.parser', 'name_gen.word_writer'],
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
    name='작명Word변환',
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
app = BUNDLE(
    exe,
    name='작명Word변환.app',
    icon=None,
    bundle_identifier=None,
)
