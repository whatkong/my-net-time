# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(
    ['time_sync_tool.py'],
    pathex=['.'],
    binaries=[],
    datas=[('icon.ico', '.')],
    hiddenimports=['pystray', 'PIL', 'PIL.Image', 'PIL.ImageDraw', 'tkinter', 'tkinter.ttk', 'tkinter.messagebox', 'tkinter.simpledialog'],
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
    name='TimeSyncTool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,  # 禁用UPX压缩，避免DLL加载问题
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # 启用控制台输出，方便调试
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
)