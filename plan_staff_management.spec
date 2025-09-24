# -*- mode: python ; coding: utf-8 -*-

# This .spec file is configured to create a single, standalone executable
# for your PyQt6 application.
#
# To build the executable, run this command in your terminal:
# pyinstaller transport_app.spec

block_cipher = None

a = Analysis(
    ['main.py'],  # The main entry point of your application
    pathex=['.'],  # Search for imports in the current directory
    binaries=[],
    datas=[],
    hiddenimports=[
        'PyQt6.sip',  # A common hidden import needed for PyQt6 apps
        'openpyxl',
        'xlsxwriter',
        'pandas'
    ],
    hookspath=[],
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
    name='TransportOperationsManager',  # The final name of your .exe file
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # This is crucial for a GUI app; it hides the console window.
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # You can specify an icon for your application here
    # icon='path/to/your/icon.ico'
)
