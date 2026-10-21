"""Prove che hanno senso solo su un Linux vero (nessuna simulazione).

Uso:  python test/test_linux.py      (con uno schermo, o con xvfb-run se non c'è)
Su Windows non fa nulla. Lo lancia test/linux/prepara_e_prova.sh insieme agli altri test.

1. Font: per ogni font del .gui quale file viene usato (cartella Fonts, sistema o sostituto),
   PDF con tutti i font e le varianti, font davvero incorporati, immagine da guardare.
2. Editor: i font della cartella Fonts arrivano a Tk tramite fontconfig, e quelli a larghezza fissa
   lo sono davvero.
3. config.dat cifrato con la chiave dell'utente e con /etc/machine-id, permessi dei file.
4. Testi di Windows (cp1252, a capo \\r\\n) letti su Linux.
5. Comandi esterni: xdg-open, lp/lpr, lpstat (stampa vera solo con FILEGUY_STAMPA_VERA=1).
"""
import os
import re
import shutil
import stat
import subprocess
import sys
import tempfile
import traceback

CARTELLA_PROGRAMMA = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CARTELLA_RUN = os.path.join(CARTELLA_PROGRAMMA, "test", "_run", "linux")
sys.path.insert(0, CARTELLA_PROGRAMMA)
sys.path.insert(0, os.path.join(CARTELLA_PROGRAMMA, "test"))
os.chdir(CARTELLA_PROGRAMMA)

if os.name == "nt":
    print("Queste prove servono su Linux: su Windows non c'è niente da provare.")
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
    try:
        with open("/etc/os-release") as f:
            nome = dict(r.rstrip().split("=", 1) for r in f if "=" in r).get("PRETTY_NAME", "?").strip('"')
    except OSError:
        nome = "?"
    nota("Sistema: %s, Python %s" % (nome, sys.version.split()[0]))
    nota("Schermo: DISPLAY=%r WAYLAND_DISPLAY=%r" % (os.environ.get("DISPLAY"), os.environ.get("WAYLAND_DISPLAY")))


@prova
def file_dei_font():
    """Ogni font di ogni codice deve trovare un file: dalla cartella Fonts, dal sistema o un sostituto."""
    modulo = servizio()
    cartella_fonts = os.path.join(CARTELLA_PROGRAMMA, "Fonts")
    for lettera, (base, nome_word) in modulo.FONT_CODICI.items():
        provenienze = []
        for variante in ("", "Bold", "It", "BoldIt"):
            percorso = modulo.trova_font(base + variante)
            assert os.path.isfile(percorso), (base + variante, percorso)
            atteso = modulo.FONT_PDF[base + variante][0]
            if os.path.dirname(percorso) == cartella_fonts:
                origine = "Fonts" if os.path.basename(percorso) == atteso else "sostituto da Fonts"
            else:
                origine = "sistema" if os.path.basename(percorso).lower() == modulo.FONT_PDF[base + variante][1].lower() \
                    else "sostituto"
            provenienze.append("%s=%s (%s)" % (variante or "normale", os.path.basename(percorso), origine))
        nota("Font %s %-16s %s" % (lettera, nome_word, "; ".join(provenienze)))
    # Quelli della cartella Fonts non devono mai dipendere dal sistema
    for base in ("Courier", "DejaVuSansMono", "IBMPlexMono", "LiberationMono", "Helvetica", "Times"):
        assert os.path.dirname(modulo.trova_font(base)) == cartella_fonts, base


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
    for atteso in ("DejaVuSansMono", "IBMPlexMono", "LiberationMono", "CourierNew"):
        assert any(atteso in n.replace(" ", "").replace("-", "") for n in incorporati), (atteso, incorporati)
    testo = "".join(p.get_text() for p in documento)
    assert "Città perché €" in testo, "accenti o euro persi"
    immagine = os.path.join(CARTELLA_RUN, "font_linux.png")
    documento[0].get_pixmap(dpi=110).save(immagine)
    nota("Da guardare: %s e %s" % (immagine, os.path.join(cartella, "font.pdf")))


@prova
def font_della_cartella_nell_editor():
    """registra_font: i font della cartella Fonts arrivano a Tk tramite fontconfig, senza installarli."""
    import tkinter as tk
    import tkinter.font as tkfont
    sistema = subprocess.run(["fc-list", ":", "family"], capture_output=True, text=True).stdout \
        if shutil.which("fc-list") else ""
    import EditorModelli  # registra i font all'importazione
    finestra = tk.Tk()
    try:
        famiglie = set(tkfont.families())
        for lettera in modello_gui.FONT_FISSI:
            # Il carattere che l'editor usa davvero: se il font manca (Consolas) quello che lo sostituisce nel PDF
            originale = modello_gui.FONT[lettera][1]
            famiglia = EditorModelli.famiglia_tk(lettera)
            assert famiglia in famiglie, "l'editor non ha un carattere per %s" % originale
            nel_sistema = famiglia in sistema
            font = tkfont.Font(family=famiglia, size=11)
            fissa = font.measure("iiiiiiiiii") == font.measure("MMMMMMMMMM")
            nota("Editor: %-16s %s%s, %s" % (originale, "" if famiglia == originale else "mostrato con %s, " % famiglia,
                                             "installato nel sistema" if nel_sistema else "dalla cartella Fonts",
                                             "larghezza fissa" if fissa else "NON a larghezza fissa"))
            assert fissa, famiglia
        # Questi sono nella cartella Fonts: devono esserci sempre
        for famiglia in ("DejaVu Sans Mono", "IBM Plex Mono", "Liberation Mono"):
            assert famiglia in famiglie, "%s non arriva all'editor" % famiglia
    finally:
        finestra.destroy()


@prova
def configurazione_cifrata_vera():
    import config_sicuro
    vecchio = os.environ.get("XDG_CONFIG_HOME")
    os.environ["XDG_CONFIG_HOME"] = tempfile.mkdtemp(dir=CARTELLA_RUN)
    percorso = os.path.join(CARTELLA_RUN, "config.dat")
    dati = dict(config_sicuro.CAMPI, cartella_lavoro="/srv/lavoro", email_password="segreta è")
    try:
        config_sicuro.salva_config(dati, percorso)
        assert config_sicuro.carica_config(percorso) == dati
        with open(percorso, "rb") as f:
            assert b"segreta" not in f.read()
        chiave = config_sicuro.file_chiave()
        assert stat.S_IMODE(os.stat(chiave).st_mode) == 0o600, oct(os.stat(chiave).st_mode)
        assert stat.S_IMODE(os.stat(os.path.dirname(chiave)).st_mode) == 0o700
        assert stat.S_IMODE(os.stat(percorso).st_mode) == 0o600
        # Tutti gli utenti del PC: chiave dal vero /etc/machine-id
        config_sicuro.salva_config(dati, percorso, tutto_il_pc=True)
        assert config_sicuro.carica_config(percorso) == dati
        assert stat.S_IMODE(os.stat(percorso).st_mode) == 0o644
        nota("config.dat: chiave in %s (600), modo 'tutto il PC' con /etc/machine-id: ok" % chiave)
    finally:
        if vecchio is None:
            os.environ.pop("XDG_CONFIG_HOME", None)
        else:
            os.environ["XDG_CONFIG_HOME"] = vecchio


@prova
def testi_scritti_su_windows():
    import piattaforma
    percorso = os.path.join(CARTELLA_RUN, "windows.txt")
    with open(percorso, "wb") as f:
        f.write("Città perché\r\nTotale 1.234,56 €\r\n".encode("cp1252"))
    assert piattaforma.leggi_righe(percorso) == ["Città perché\n", "Totale 1.234,56 €\n"]


@prova
def comandi_esterni():
    import stampa_pdf
    for comando in ("xdg-open", "lp", "lpr", "lpstat", "fc-list"):
        nota("Comando %-8s %s" % (comando, shutil.which(comando) or "MANCA"))
    stampanti = stampa_pdf.elenca_stampanti()
    nota("Stampanti CUPS: %s" % (", ".join(stampanti) or "nessuna"))
    if os.environ.get("FILEGUY_STAMPA_VERA") == "1":
        documento = pymupdf.open()
        documento.new_page().insert_text((72, 72), "Prova di stampa FileGUY da Linux", fontsize=16)
        percorso = os.path.join(CARTELLA_RUN, "stampa.pdf")
        documento.save(percorso)
        stampa_pdf.stampa_pdf(percorso, os.environ.get("FILEGUY_STAMPANTE") or None)
        nota("Stampa vera mandata a %s" % (os.environ.get("FILEGUY_STAMPANTE") or "stampante predefinita"))
    assert shutil.which("lp") or shutil.which("lpr"), "manca il comando di stampa: installa cups-client"


print("\n".join(note))
print()
for nome, errore in risultati:
    print("[%s] %s" % ("OK " if errore is None else "ERR", nome))
    if errore:
        print("       " + errore.replace("\n", "\n       "))
print("\n%d/%d prove OK (Linux)" % (sum(e is None for _, e in risultati), len(risultati)))
sys.exit(0 if all(e is None for _, e in risultati) else 1)
