# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files

datas = []
datas += collect_data_files('rapidocr_onnxruntime')

a = Analysis(
    ['pdf_unified_tool.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[
        'fitz',
        'rapidocr_onnxruntime',
        'rapidocr_onnxruntime.main',
        'rapidocr_onnxruntime.utils',
        'rapidocr_onnxruntime.utils.infer_engine',
        'rapidocr_onnxruntime.utils.parse_parameters',
        'rapidocr_onnxruntime.ch_ppocr_det',
        'rapidocr_onnxruntime.ch_ppocr_det.text_detect',
        'rapidocr_onnxruntime.ch_ppocr_rec',
        'rapidocr_onnxruntime.ch_ppocr_rec.text_recognize',
        'rapidocr_onnxruntime.ch_ppocr_cls',
        'rapidocr_onnxruntime.ch_ppocr_cls.text_cls',
        'PIL',
        'numpy',
        'onnxruntime',
        'cv2',
        'shapely',
        'yaml',
    ],
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
    name='PDF智能处理工具',
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
    icon=None,
)
