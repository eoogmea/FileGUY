# -*- mode: python ; coding: utf-8 -*-
# Crea SingleGUY_editor.exe, il programma che back.py fa partire.
# Uso (dalla cartella del programma):  pyinstaller app.spec
# Accanto all'exe vanno messi config.dat (si crea con ConfigManager) e la cartella Fonts
# (se manca usa i font di Windows, con Arial al posto di Helvetica).

from PyInstaller.utils.hooks import collect_data_files

block_cipher = None


a = Analysis(['SingleGUY_editor.py'],
             pathex=[],
             binaries=[],
             # Font usato da python-barcode per scrivere il numero sotto il codice
             datas=collect_data_files('barcode'),
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             # PyMuPDF 2.0 attiva da solo pymupdf.layout (un modello onnx) se è installato:
             # nell'exe non trova i suoi file e il programma non parte. Non serve, quindi si esclude
             # (PyMuPDF lo importa dentro un try/except ImportError)
             excludes=['pymupdf.layout', 'pymupdf4llm', 'onnxruntime'],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='SingleGUY_editor',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=True )
