"""Prove che hanno senso solo su un Windows vero (nessuna simulazione): il gemello di test_linux.py.

Uso:  python test\\test_windows.py
Su Linux non fa nulla. Lo lancia test\\windows\\prepara_e_prova.bat insieme agli altri test.
Con la variabile FILEGUY_PROVA (la cartella dei programmi, es. C:\\prova) controlla anche il suo config.dat.

1. Font: per ogni font del .gui quale file viene usato (cartella Fonts, Windows o sostituto),
   PDF con tutti i font e le varianti, font davvero incorporati, immagine da guardare.
2. Editor: i font della cartella Fonts arrivano a Tk (registrati solo per il programma) e quelli
   a larghezza fissa lo sono davvero.
3. config.dat cifrato con DPAPI (utente e tutto il PC) e il config.dat della cartella di prova.
4. Testi di Linux (UTF-8, a capo \\n) e di Windows (cp1252) letti su Windows.
5. Stampa: stampanti installate, stampa vera su "Microsoft Print to PDF" (su un file), programma che apre i
   PDF; stampa sulla stampante vera solo con FILEGUY_STAMPA_VERA=1.
"""
import os
import platform
import re
import shutil
import sys
import tempfile
import traceback

CARTELLA_PROGRAMMA = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CARTELLA_RUN = os.path.join(CARTELLA_PROGRAMMA, "test", "_run", "windows")
sys.path.insert(0, CARTELLA_PROGRAMMA)
sys.path.insert(0, os.path.join(CARTELLA_PROGRAMMA, "test"))
os.chdir(CARTELLA_PROGRAMMA)

if os.name != "nt":
    print("Queste prove servono su Windows: qui non c'è niente da provare.")
    sys.exit(0)

import modello_gui
import pymupdf

risultati = []
note = []
shutil.rmtree(CARTELLA_RUN, ignore_errors=True)
os.makedirs(CARTELLA_RUN)


def prova(funzione):
    try:
        funzione()
        risultati.append((funzione.__name__, None))
    except Exception:
        risultati.append((funzione.__name__, traceback.format_exc(limit=-3).strip()))
    return funzione


def nota(testo):
    note.append(testo)


def servizio():
    import esegui_test
    return esegui_test.carica_programma("SingleGUY")


@prova
def sistema():
    nota("Sistema: %s %s, Python %s" % (platform.system(), platform.version(), sys.version.split()[0]))


@prova
def file_dei_font():
    """Ogni font di ogni codice deve trovare un file: dalla cartella Fonts, da Windows o un sostituto."""
    import piattaforma
    modulo = servizio()
    cartella_fonts = os.path.join(CARTELLA_PROGRAMMA, "Fonts")
    cartelle_sistema = [os.path.normcase(c) for c in piattaforma.CARTELLE_FONT_SISTEMA]
    for lettera, (base, nome_word) in modulo.FONT_CODICI.items():
        provenienze = []
        for variante in ("", "Bold", "It", "BoldIt"):
            percorso = modulo.trova_font(base + variante)
            assert os.path.isfile(percorso), (base + variante, percorso)
            if os.path.normcase(os.path.dirname(percorso)) == os.path.normcase(cartella_fonts):
                origine = "Fonts" if os.path.basename(percorso) == modulo.FONT_PDF[base + variante][0] \
                    else "sostituto da Fonts"
            elif os.path.normcase(os.path.dirname(percorso)) in cartelle_sistema:
                origine = "Windows"
            else:
                origine = "altro"
            provenienze.append("%s=%s (%s)" % (variante or "normale", os.path.basename(percorso), origine))
        nota("Font %s %-16s %s" % (lettera, nome_word, "; ".join(provenienze)))
    # Quelli della cartella Fonts non devono mai dipendere dal sistema
    for base in ("Courier", "DejaVuSansMono", "IBMPlexMono", "LiberationMono", "Helvetica", "Times"):
        assert os.path.normcase(os.path.dirname(modulo.trova_font(base))) == os.path.normcase(cartella_fonts), base


@prova
def pdf_con_tutti_i_font():
    """Un PDF con ogni font e variante (grassetto, corsivo), accenti ed euro; i font devono esserci davvero."""
    righe, formati = [], []
    for lettera, (nome, _) in modello_gui.FONT.items():
        for grassetto, corsivo in ((False, False), (True, False), (False, True), (True, True)):
            testo = "%s %-16s %-9s Città perché € 0123456789 iiiii MMMMM" % (
                lettera, nome, ("grassetto " if grassetto else "") + ("corsivo" if corsivo else "") or "normale")
            righe.append(testo)
            formati.append([(lettera, grassetto, corsivo, 9, "n")] * len(testo))
    cartella = tempfile.mkdtemp(dir=CARTELLA_RUN)
    with open(os.path.join(cartella, "font.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(righe) + "\n")
    intestazione = dict(modello_gui.nuovo_modello(False), testo="font.txt", uscita="font.pdf")
    modello_gui.scrivi_gui(os.path.join(cartella, "font.gui"), intestazione, modello_gui.linee_formato(formati, False))
    with open(os.path.join(cartella, "font.gui")) as f:
        guida = f.readlines()
    servizio().pdf(guida, {"cartella_lavoro": cartella, "cartella_template": cartella, "cartella_output": cartella})
    documento = pymupdf.open(os.path.join(cartella, "font.pdf"))
    incorporati = {re.sub(r"^[A-Z]{6}\+", "", f[3]) for pagina in documento for f in pagina.get_fonts()}
    nota("Font incorporati nel PDF: " + ", ".join(sorted(incorporati)))
    # Su Windows ci sono anche i font di sistema: nessun sostituto per Consolas, Calibri...
    for atteso in ("Consolas", "Calibri", "Verdana", "Georgia", "SegoeUI", "DejaVuSansMono", "IBMPlexMono",
                   "LiberationMono", "CourierNew", "TimesNewRoman", "Helvetica"):
        assert any(atteso in n.replace(" ", "").replace("-", "") for n in incorporati), (atteso, incorporati)
    testo = "".join(p.get_text() for p in documento)
    assert "Città perché €" in testo, "accenti o euro persi"
    immagine = os.path.join(CARTELLA_RUN, "font_windows.png")
    documento[0].get_pixmap(dpi=110).save(immagine)
    nota("Da guardare: %s e %s" % (immagine, os.path.join(cartella, "font.pdf")))


@prova
def font_della_cartella_nell_editor():
    """registra_font: i font della cartella Fonts arrivano a Tk senza installarli in Windows."""
    import tkinter as tk
    import tkinter.font as tkfont
    import EditorModelli  # registra i font all'importazione
    finestra = tk.Tk()
    try:
        famiglie = set(tkfont.families())
        for lettera in modello_gui.FONT_FISSI:
            originale = modello_gui.FONT[lettera][1]
            famiglia = EditorModelli.famiglia_tk(lettera)
            assert famiglia in famiglie, "l'editor non ha un carattere per %s" % originale
            font = tkfont.Font(family=famiglia, size=11)
            fissa = font.measure("iiiiiiiiii") == font.measure("MMMMMMMMMM")
            nota("Editor: %-16s %s%s" % (originale, "" if famiglia == originale else "mostrato con %s, " % famiglia,
                                         "larghezza fissa" if fissa else "NON a larghezza fissa"))
            assert fissa, famiglia
        for famiglia in ("DejaVu Sans Mono", "IBM Plex Mono", "Liberation Mono", "Helvetica"):
            assert famiglia in famiglie, "%s non arriva all'editor" % famiglia
    finally:
        finestra.destroy()


@prova
def configurazione_cifrata_vera():
    import config_sicuro
    percorso = os.path.join(CARTELLA_RUN, "config.dat")
    dati = dict(config_sicuro.CAMPI, cartella_lavoro="C:\\lavoro", email_password="segreta è")
    for tutto_il_pc in (False, True):
        config_sicuro.salva_config(dati, percorso, tutto_il_pc=tutto_il_pc)
        with open(percorso, "rb") as f:
            contenuto = f.read()
        assert contenuto.startswith(b"INFASCFG\x01") and b"segreta" not in contenuto
        assert config_sicuro.carica_config(percorso) == dati
    nota("config.dat: DPAPI per l'utente e per tutto il PC: ok")
    # Il config.dat della cartella di prova (quello che usano gli .exe)
    prova_programmi = os.environ.get("FILEGUY_PROVA")
    if prova_programmi:
        percorso = os.path.join(prova_programmi, "config.dat")
        if not os.path.isfile(percorso):
            nota("config.dat di %s: non c'è (crealo con ConfigManager)" % prova_programmi)
            return
        config = config_sicuro.carica_config(percorso)
        for campo in ("cartella_lavoro", "cartella_template", "cartella_output"):
            nota("config.dat di %s: %-17s %s %s" % (prova_programmi, campo, config[campo],
                                                     "(c'è)" if os.path.isdir(config[campo]) else "(NON ESISTE)"))
            assert config[campo] == "" or os.path.isdir(config[campo]), (campo, config[campo])
        nota("config.dat di %s: stampante %r, log %r" % (prova_programmi, config["stampante"] or "predefinita",
                                                          config["file_log"] or "FileGUY.log accanto agli exe"))


@prova
def testi_scritti_su_linux_e_su_windows():
    import piattaforma
    for nome, dati in (("linux.txt", "Città perché\nTotale 1.234,56 €\n".encode("utf-8")),
                       ("windows.txt", "Città perché\r\nTotale 1.234,56 €\r\n".encode("cp1252"))):
        percorso = os.path.join(CARTELLA_RUN, nome)
        with open(percorso, "wb") as f:
            f.write(dati)
        assert piattaforma.leggi_righe(percorso) == ["Città perché\n", "Totale 1.234,56 €\n"], nome


@prova
def stampa_e_pdf_a_schermo():
    import winreg
    import stampa_pdf
    import win32print
    stampanti = stampa_pdf.elenca_stampanti()
    nota("Stampanti: %s" % ", ".join(stampanti))
    nota("Stampante predefinita: %s" % win32print.GetDefaultPrinter())
    # Programma che apre i PDF con l'azione "Crea e apri a schermo"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Explorer"
                            r"\FileExts\.pdf\UserChoice") as chiave:
            programma = winreg.QueryValueEx(chiave, "ProgId")[0]
    except OSError:
        try:
            programma = winreg.QueryValue(winreg.HKEY_CLASSES_ROOT, ".pdf")
        except OSError:
            programma = ""
    nota("I PDF si aprono con: %s" % (programma or "NESSUN PROGRAMMA (l'azione 'Crea e apri a schermo' non funziona)"))
    # Stampa vera, ma su un file
    if "Microsoft Print to PDF" in stampanti:
        documento = pymupdf.open()
        for n in range(2):
            documento.new_page().insert_text((72, 72), "Prova di stampa FileGUY, pagina %d" % (n + 1), fontsize=16)
        sorgente = os.path.join(CARTELLA_RUN, "da_stampare.pdf")
        documento.save(sorgente)
        uscita = os.path.join(CARTELLA_RUN, "stampato.pdf")
        stampa_pdf.stampa_pdf(sorgente, "Microsoft Print to PDF", uscita)
        assert pymupdf.open(uscita).page_count == 2
        nota("Stampa su Microsoft Print to PDF: ok (%s)" % uscita)
    else:
        nota("Microsoft Print to PDF non installata: stampa di prova saltata")
    if os.environ.get("FILEGUY_STAMPA_VERA") == "1":
        documento = pymupdf.open()
        documento.new_page().insert_text((72, 72), "Prova di stampa FileGUY da Windows", fontsize=16)
        percorso = os.path.join(CARTELLA_RUN, "stampa_vera.pdf")
        documento.save(percorso)
        stampa_pdf.stampa_pdf(percorso, os.environ.get("FILEGUY_STAMPANTE") or None)
        nota("Stampa vera mandata a %s" % (os.environ.get("FILEGUY_STAMPANTE") or win32print.GetDefaultPrinter()))


print("\n".join(note))
print()
for nome, errore in risultati:
    print("[%s] %s" % ("OK " if errore is None else "ERR", nome))
    if errore:
        print("       " + errore.replace("\n", "\n       "))
print("\n%d/%d prove OK (Windows)" % (sum(e is None for _, e in risultati), len(risultati)))
sys.exit(0 if all(e is None for _, e in risultati) else 1)
