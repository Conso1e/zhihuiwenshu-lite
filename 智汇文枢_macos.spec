# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec — macOS .app 打包配置
构建命令: pyinstaller 智汇文枢_macos.spec
"""

import sys
from pathlib import Path

# -------- 图标处理 --------
# macOS 需要 .icns 格式。如果有 icon.icns 文件则使用，否则跳过图标
icon_path = None
candidates = [
    Path(__file__).parent / 'icon.icns',
    Path(__file__).parent / 'icon.png',
]
for c in candidates:
    if c.exists():
        icon_path = str(c)
        break

# -------- hidden imports --------
# customtkinter 及其内部模块 + 项目自有模块
hiddenimports = [
    'customtkinter',
    'customtkinter.windows',
    'customtkinter.windows.widgets',
    'pandas',
    'openpyxl',
    'docxtpl',
    'docx',
    'jinja2',
    'darkdetect',
    # 项目模块
    'config',
    'ui',
    'ui.main_window',
    'ui.widgets',
    'services',
    'services.excel_processor',
    'services.word_generator',
    'models',
    'models.id_parser',
    'models.formula_engine',
    'models.fee_calculator',
]

# -------- Analysis --------
a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',  # customtkinter 不需要标准 tkinter
    ],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

# -------- EXE (生成 Unix 可执行文件) --------
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='智汇文枢',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # macOS: windowed 模式
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,        # None = 当前架构；设为 'universal2' 可同时支持 Intel + ARM
    codesign_identity=None,  # 无签名（本地打开需右键→打开）
    entitlements_file=None,
)

# -------- BUNDLE (收集到 .app 包中) --------
app = BUNDLE(
    exe,
    name='智汇文枢.app',
    icon=icon_path,
    bundle_identifier='com.zhihuiwenshu.app',
    info_plist={
        'NSPrincipalClass': 'NSApplication',
        'NSHighResolutionCapable': 'True',
        'CFBundleShortVersionString': '4.0.0',
        'CFBundleVersion': '4.0.0',
        'CFBundleName': '智汇文枢',
        'CFBundleDisplayName': '智汇文枢',
        'LSMinimumSystemVersion': '11.0',
    },
)
