"""Stampa di un PDF su una stampante di Windows senza Ghostscript.

Le pagine vengono trasformate in immagini con PyMuPDF (fitz) e mandate alla stampante
con le funzioni di stampa di Windows (pywin32), librerie già usate dal programma.
"""
import os

import fitz
import win32con
import win32print
import win32ui
from PIL import Image, ImageWin

# Risoluzione massima con cui vengono trasformate le pagine (più alta = più nitida ma più lenta)
DPI_MASSIMI = 300


def stampa_pdf(percorso, stampante=None, file_uscita=None):
    """Stampa il PDF sulla stampante indicata (se manca, quella predefinita).

    Le pagine sono stampate a grandezza reale e rimpicciolite solo se non entrano nel foglio;
    se l'orientamento della pagina è diverso da quello del foglio la pagina viene ruotata.
    file_uscita serve solo con stampanti virtuali come "Microsoft Print to PDF",
    per non far comparire la finestra "Salva con nome".
    """
    if stampante is None:
        stampante = win32print.GetDefaultPrinter()

    dc = win32ui.CreateDC()
    dc.CreatePrinterDC(stampante)
    try:
        # Area stampabile del foglio in pixel e risoluzione della stampante
        area_x = dc.GetDeviceCaps(win32con.HORZRES)
        area_y = dc.GetDeviceCaps(win32con.VERTRES)
        dpi_x = dc.GetDeviceCaps(win32con.LOGPIXELSX)
        dpi_y = dc.GetDeviceCaps(win32con.LOGPIXELSY)

        doc = fitz.open(percorso)
        dc.StartDoc(os.path.basename(percorso), file_uscita)
        try:
            for pagina in doc:
                # Dimensioni della pagina in pollici (il PDF misura in punti, 72 per pollice)
                larghezza = pagina.rect.width / 72
                altezza = pagina.rect.height / 72
                ruota = (larghezza > altezza) != (area_x > area_y)
                if ruota:
                    larghezza, altezza = altezza, larghezza

                # Dimensione reale in pixel della stampante, ridotta se non entra nell'area stampabile
                scala = min(1.0, area_x / (larghezza * dpi_x), area_y / (altezza * dpi_y))
                w = int(larghezza * dpi_x * scala)
                h = int(altezza * dpi_y * scala)

                pix = pagina.get_pixmap(dpi=max(1, int(min(DPI_MASSIMI, max(dpi_x, dpi_y) * scala))), alpha=False)
                immagine = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                if ruota:
                    immagine = immagine.rotate(90, expand=True)

                dc.StartPage()
                ImageWin.Dib(immagine).draw(dc.GetHandleOutput(), (0, 0, w, h))
                dc.EndPage()
            dc.EndDoc()
        except Exception:
            dc.AbortDoc()
            raise
        finally:
            doc.close()
    finally:
        dc.DeleteDC()
