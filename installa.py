"""Installa tutto quello che serve a FileGUY, su Windows e su Linux.

Uso (dalla cartella del programma; di solito lo avviano installa.bat o installa.sh):
    python installa.py              pacchetti del sistema (solo Linux), ambiente Python e librerie
    python installa.py --compila    in più crea gli eseguibili in dist/ (con Fonts e guide accanto)
    python installa.py --controlla  controlla soltanto che tutto sia installato

1. Linux: i pacchetti del sistema che pip non installa (tkinter per le finestre, CUPS per la
   stampa, xdg-utils per aprire i PDF, fontconfig per i font), con apt, dnf, zypper o pacman.
   Serve la password di sudo.
2. L'ambiente Python .venv-windows o .venv-linux nella cartella del programma, con le librerie di
   requirements.txt: così non si toccano le librerie del sistema.
3. Il controllo che ogni libreria si importi e che le finestre si possano aprire.
4. Con --compila: SingleGUY, back, ConfigManager, EditorModelli ed EditorExcel con PyInstaller.

Gli eseguibili creati contengono già tutte le librerie: sul PC dove si usano non serve installare
niente, basta copiare la cartella dist (con config.dat creato da ConfigManager).
Lo script non si ferma con un traceback: ogni passo dice cosa è andato storto e come rimediare.
"""
import os
import shutil
import subprocess
import sys

CARTELLA = os.path.dirname(os.path.abspath(__file__))
WINDOWS = os.name == "nt"
AMBIENTE = os.path.join(CARTELLA, ".venv-windows" if WINDOWS else ".venv-linux")
PYTHON_AMBIENTE = os.path.join(AMBIENTE, "Scripts", "python.exe") if WINDOWS else os.path.join(AMBIENTE, "bin",
                                                                                               "python")
VERSIONE_MINIMA = (3, 9)

# Pacchetti del sistema su Linux: gestore -> (comando prima dell'installazione, comando di installazione)
PACCHETTI_LINUX = {
    "apt-get": (["apt-get", "update"],
                ["apt-get", "install", "-y", "python3", "python3-venv", "python3-pip", "python3-tk", "cups-client",
                 "xdg-utils", "fontconfig"]),
    "dnf": (None, ["dnf", "install", "-y", "python3", "python3-pip", "python3-tkinter", "cups-client", "xdg-utils",
                   "fontconfig"]),
    "zypper": (None, ["zypper", "install", "-y", "python3", "python3-pip", "python3-tk", "cups-client",
                      "xdg-utils", "fontconfig"]),
    "pacman": (None, ["pacman", "-S", "--needed", "--noconfirm", "python", "python-pip", "tk", "cups", "xdg-utils",
                      "fontconfig"]),
}

# Libreria da importare -> pacchetto di requirements.txt (per dire cosa manca)
LIBRERIE = {
    "PyPDF2": "PyPDF2", "fpdf": "fpdf", "docx": "python-docx", "openpyxl": "openpyxl", "barcode": "python-barcode",
    "PIL": "pillow", "pymupdf": "PyMuPDF", "requests": "requests", "telepot": "telepot",
}
if WINDOWS:
    LIBRERIE["win32print"] = "pywin32"
else:
    LIBRERIE["cryptography"] = "cryptography"

PROGRAMMI = ["app.spec", "back.spec", "ConfigManager.spec", "EditorModelli.spec", "EditorExcel.spec"]


class Errore(Exception):
    """Un passo non riuscito: il messaggio dice cosa fare."""


def titolo(testo):
    print("\n=== " + testo + " ===", flush=True)


def esegui(comando, sudo=False, descrizione=None, **opzioni):
    """Esegue un comando mostrando l'uscita; Errore se non riesce."""
    if sudo and hasattr(os, "geteuid") and os.geteuid() != 0:
        comando = ["sudo"] + comando
    print("> " + (descrizione or " ".join(comando)), flush=True)
    try:
        esito = subprocess.call(comando, cwd=CARTELLA, **opzioni)
    except OSError as e:
        raise Errore("impossibile avviare %s: %s" % (comando[0], e))
    if esito != 0:
        raise Errore("il comando non è riuscito (codice %d): %s" % (esito, " ".join(comando)))


def controlla_python():
    if sys.version_info < VERSIONE_MINIMA:
        raise Errore("serve Python %d.%d o più recente (questo è %s)" % (VERSIONE_MINIMA + (sys.version.split()[0],)))
    print("Python %s: %s" % (sys.version.split()[0], sys.executable))


def pacchetti_linux():
    titolo("Pacchetti del sistema (serve la password di sudo)")
    for gestore, (prima, installa) in PACCHETTI_LINUX.items():
        if shutil.which(gestore):
            if prima:
                esegui(prima, sudo=True)
            esegui(installa, sudo=True)
            return
    print("Gestore dei pacchetti non riconosciuto: installa a mano python3, il modulo tkinter, cups-client, "
          "xdg-utils e fontconfig")


def crea_ambiente():
    titolo("Ambiente Python " + os.path.basename(AMBIENTE))
    if not os.path.isfile(PYTHON_AMBIENTE):
        try:
            esegui([sys.executable, "-m", "venv", AMBIENTE])
        except Errore:
            raise Errore("non si riesce a creare l'ambiente Python" +
                         ("" if WINDOWS else " (su Debian/Ubuntu serve il pacchetto python3-venv)"))
    else:
        print("C'è già: " + AMBIENTE)
    esegui([PYTHON_AMBIENTE, "-m", "pip", "install", "--upgrade", "pip"])
    try:
        esegui([PYTHON_AMBIENTE, "-m", "pip", "install", "-r", os.path.join(CARTELLA, "requirements.txt")])
    except Errore:
        raise Errore("l'installazione delle librerie non è riuscita: controlla la connessione a internet e "
                     "rilancia (i messaggi di pip sono qui sopra)")


def controlla(python):
    """Controlla con il Python indicato che si importino tutte le librerie e tkinter."""
    titolo("Controllo delle librerie")
    script = ("import importlib, sys\n"
              "mancano = []\n"
              "for nome, pacchetto in %r.items():\n"
              "    try:\n"
              "        importlib.import_module(nome)\n"
              "    except Exception as e:\n"
              "        mancano.append('%%s (%%s): %%s' %% (pacchetto, nome, e))\n"
              "try:\n"
              "    import tkinter\n"
              "except Exception as e:\n"
              "    mancano.append('tkinter (finestre): %%s' %% e)\n"
              "print('\\n'.join(mancano) if mancano else 'Tutte le librerie ci sono')\n"
              "sys.exit(1 if mancano else 0)\n") % LIBRERIE
    try:
        esegui([python, "-c", script], descrizione=python + " (importa ogni libreria)")
    except Errore:
        raise Errore("mancano delle librerie (sopra l'elenco): rilancia l'installazione" +
                     ("" if WINDOWS else "; per tkinter serve il pacchetto del sistema (python3-tk)"))
    for programma in ("lp", "xdg-open", "fc-list") if not WINDOWS else ():
        print("%-9s %s" % (programma, "c'è" if shutil.which(programma) else
                           "MANCA (stampa, apertura dei PDF o font potrebbero non funzionare)"))


def compila():
    titolo("Eseguibili (PyInstaller)")
    dist = os.path.join(CARTELLA, "dist")
    for spec in PROGRAMMI:
        esegui([PYTHON_AMBIENTE, "-m", "PyInstaller", "--noconfirm", spec])
    # Accanto agli eseguibili: i font del PDF e le guide
    shutil.copytree(os.path.join(CARTELLA, "Fonts"), os.path.join(dist, "Fonts"), dirs_exist_ok=True)
    for guida in ("Guida_EditorModelli.html", "Guida_file_gui.html", "guida_whatsapp.txt"):
        if os.path.isfile(os.path.join(CARTELLA, guida)):
            shutil.copy(os.path.join(CARTELLA, guida), dist)
    print("\nEseguibili in " + dist + ": copia la cartella dove deve girare il servizio, poi crea config.dat "
          "con ConfigManager.")


def principale(argomenti):
    solo_controllo = "--controlla" in argomenti
    try:
        controlla_python()
        if solo_controllo:
            controlla(PYTHON_AMBIENTE if os.path.isfile(PYTHON_AMBIENTE) else sys.executable)
            return 0
        if not WINDOWS:
            pacchetti_linux()
        crea_ambiente()
        controlla(PYTHON_AMBIENTE)
        if "--compila" in argomenti:
            compila()
    except Errore as e:
        print("\nERRORE: %s" % e, file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("\nInterrotto.", file=sys.stderr)
        return 1
    attiva = AMBIENTE + (r"\Scripts\activate" if WINDOWS else "/bin/activate")
    print("\nTUTTO PRONTO. Per avviare i programmi con queste librerie:")
    print("    " + PYTHON_AMBIENTE + " ConfigManager.py   (poi EditorModelli.py, back.py...)")
    print("oppure attiva l'ambiente:  " + ("" if WINDOWS else ". ") + attiva)
    return 0


if __name__ == "__main__":
    sys.exit(principale(sys.argv[1:]))
