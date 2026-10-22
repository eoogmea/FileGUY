# -*- mode: python ; coding: utf-8 -*-
# Crea EditorModelli.exe (su Linux: EditorModelli), l'editor dei modelli .gui.
# Uso (dalla cartella del programma):  pyinstaller EditorModelli.spec
# Accanto all'exe vanno messi config.dat e la cartella Fonts, come per SingleGUY.exe.
# I modelli vengono salvati nella cartella ModelliGUI accanto all'exe.

import sys

from PyInstaller.utils.hooks import collect_data_files

block_cipher = None


a = Analysis(['EditorModelli.py'],
             pathex=[],
             binaries=[],
             # Font usato da python-barcode (il codice del servizio viene caricato per l'anteprima)
             # e la guida aperta dal menu Aiuto
             datas=collect_data_files('barcode') + [('Guida_EditorModelli.html', '.'), ('Guida_file_gui.html', '.')],
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
          name='EditorModelli',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False )
