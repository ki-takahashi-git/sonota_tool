# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['shikyu_check_tool.py'],
    pathex=[],
    binaries=[],
    datas=[('使い方.md', '.'), ('VERSION', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pandas', 'scipy', 'matplotlib', 'cv2', 'pyarrow', 'uvicorn', 'websockets', 'IPython', 'notebook', 'jupyter', 'jupyter_client', 'jupyter_core', 'torch', 'tensorflow', 'sklearn', 'numpy.f2py', 'tornado', 'zmq', 'PyQt5', 'PySide2', 'PySide6', 'PyQt6', 'docx', 'pptx', 'pydantic', 'fastapi', 'gradio', 'streamlit'],
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
    name='マミヤ支給明細照合ツール',
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
