# -*- mode: python ; coding: utf-8 -*-
# Crea EditorExcel.exe (su Linux: EditorExcel), l'editor dei modelli .gui per i file Excel.
# Uso (dalla cartella del programma):  pyinstaller EditorExcel.spec
# Accanto all'exe va messo config.dat, come per SingleGUY.exe.
# I modelli vengono salvati nella cartella ModelliGUI accanto all'exe.

import sys

from PyInstaller.utils.hooks import collect_data_files

block_cipher = None


a = Analysis(['EditorExcel.py'],
             pathex=[],
             binaries=[],
             # Font usato da python-barcode (il codice del servizio viene caricato per l'anteprima)
             # e la guida aperta dal menu Aiuto
             datas=collect_data_files('barcode') + [('Guida_EditorModelli.html', '.')],
             # Il codice del servizio è importato per nome solo al momento dell'anteprima
             hiddenimports=['SingleGUY'],
             hookspath=[],
             runtime_hooks=[],
             # Come in app.spec: pymupdf.layout nell'exe non trova i suoi file e non serve
             excludes=['pymupdf.layout', 'pymupdf4llm', 'onnxruntime']
                      # cryptography serve solo su Linux per config.dat (su Windows c'è DPAPI)
                      + (['cryptography'] if sys.platform == 'win32' else []),
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
          name='EditorExcel',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False )
