# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [
    ('MySchoolChecks\\checks', 'checks'),
    ('MySchoolChecks\\core', 'core'),
    ('MySchoolChecks\\smeae', 'smeae'),
    ('MySchoolChecks\\data', 'data'),
    ('MySchoolChecks\\config.py', '.'),
    ('MySchoolChecks\\editor.py', '.'),
    ('MySchoolChecks\\placements.py', '.'),
    ('MySchoolChecks\\placements_template.xlsx', '.'),
    ('MySchoolChecks\\startup.mp3', '.'),
    ('MySchoolChecks\\app.ico', '.'),
    ('MySchoolChecks_Odigos.pdf', '.'),
]
binaries = []
hiddenimports = [
    'smtplib', 'ssl', 'csv',
    'email.mime.multipart', 'email.mime.base', 'email.mime.text',
    'selenium', 'pandas', 'openpyxl',
]
tmp_ret = collect_all('selenium')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('email')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('pandas')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('openpyxl')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['MySchoolChecks\\main.py'],
    pathex=['MySchoolChecks'],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    name='MySchoolChecks',
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
    icon='MySchoolChecks\\app.ico',
    version=None,
    uac_admin=False,
)