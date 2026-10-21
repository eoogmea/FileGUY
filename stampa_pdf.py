"""Stampa di un PDF senza Ghostscript, su Windows e su Linux.

Windows: le pagine vengono trasformate in immagini con PyMuPDF e mandate alla stampante
con le funzioni di stampa di Windows (pywin32), librerie già usate dal programma.
Linux e macOS: il PDF viene mandato a CUPS con il comando lp (pacchetto cups-client, presente
in tutte le distribuzioni con la stampa configurata), o lpr se lp manca. CUPS stampa i PDF da solo.
"""
import os
import shutil
import subprocess

from piattaforma import WINDOWS, ambiente_esterno

# Risoluzione massima con cui vengono trasformate le pagine (più alta = più nitida ma più lenta)
DPI_MASSIMI = 300


def stampa_pdf(percorso, stampante=None, file_uscita=None):
    """Stampa il PDF sulla stampante indicata (se manca o è vuota, quella predefinita).

    Le pagine sono stampate a grandezza reale e rimpicciolite solo se non entrano nel foglio;
    se l'orientamento della pagina è diverso da quello del foglio la pagina viene ruotata.
    file_uscita serve solo su Windows con stampanti virtuali come "Microsoft Print to PDF",
    per non far comparire la finestra "Salva con nome".
    """
    if WINDOWS:
        _stampa_windows(percorso, stampante or None, file_uscita)
    else:
        _stampa_cups(percorso, stampante or None)


def elenca_stampanti():
    """Nomi delle stampanti installate (lista vuota se non si riesce a leggerle)."""
    if WINDOWS:
        import win32print
        tipi = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        return sorted(p[2] for p in win32print.EnumPrinters(tipi))
    if shutil.which("lpstat") is None:
        return []
    try:
        uscita = subprocess.run(["lpstat", "-e"], capture_output=True, text=True, timeout=10,
                                env=ambiente_esterno()).stdout
    except (OSError, subprocess.SubprocessError):
        return []
    return sorted(riga.split()[0] for riga in uscita.splitlines() if riga.strip())


def _stampa_cups(percorso, stampante):
    # print-scaling=auto: grandezza reale, rimpicciolita solo se non entra nel foglio (come su Windows);
    # CUPS ruota da solo le pagine orizzontali. Le opzioni sconosciute a una stampante vengono ignorate.
    if shutil.which("lp"):
        comando = ["lp", "-o", "print-scaling=auto"]
        if stampante:
            comando += ["-d", stampante]
    elif shutil.which("lpr"):
        comando = ["lpr", "-o", "print-scaling=auto"]
        if stampante:
            comando += ["-P", stampante]
    else:
        raise RuntimeError("Comando di stampa non trovato (lp o lpr): installa CUPS (pacchetto cups-client)")
    risultato = subprocess.run(comando + ["--", percorso], capture_output=True, text=True, timeout=120,
                               env=ambiente_esterno())
    if risultato.returncode != 0:
        messaggio = (risultato.stderr or risultato.stdout).strip()
        raise RuntimeError("Stampa non riuscita (" + " ".join(comando) + "): " + messaggio)


def _stampa_windows(percorso, stampante, file_uscita):
    import pymupdf
    import win32con
    import win32print
    import win32ui
    from PIL import Image, ImageWin

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

        doc = pymupdf.open(percorso)
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
