# -*- mode: python ; coding: utf-8 -*-
import re
import os

# Get current version
with open("MC_Recon_UI.py", "r", encoding="utf-8") as f:
    content = f.read()
    version_match = re.search(r"VERSION = '([\d\.]+)'", content)
    current_version = version_match.group(1) if version_match else "unknown"

block_cipher = None

a = Analysis(
    ['MC_Recon_UI.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
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
    name=f'MC_Recon_Tool_v{current_version}',
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
    icon='favicon.ico',
)