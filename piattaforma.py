"""Funzioni che cambiano tra Windows e Linux (e macOS).

- lettura dei file di testo: quelli creati su Windows sono in cp1252, quelli di Linux in UTF-8
- apertura di un documento con il programma predefinito (il PDF a schermo)
- ricerca dei font installati nel sistema
- ambiente per i programmi esterni avviati dall'eseguibile creato con PyInstaller
"""
import os
import subprocess
import sys

WINDOWS = os.name == "nt"
MACOS = sys.platform == "darwin"

if WINDOWS:
    CARTELLE_FONT_SISTEMA = [
        os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts"),
    ]
elif MACOS:
    CARTELLE_FONT_SISTEMA = [os.path.expanduser("~/Library/Fonts"), "/Library/Fonts", "/System/Library/Fonts"]
else:
    # Cartelle standard di fontconfig, uguali in tutte le distribuzioni
    CARTELLE_FONT_SISTEMA = [
        os.path.join(os.environ.get("XDG_DATA_HOME") or os.path.expanduser("~/.local/share"), "fonts"),
        os.path.expanduser("~/.fonts"),
        "/usr/local/share/fonts",
        "/usr/share/fonts",
    ]


def console_sicura():
    """Nei programmi senza console (gli editor e ConfigManager creati con PyInstaller) sys.stdout e
    sys.stderr sono None: chi ci scrive fallisce, per esempio PyPDF2 quando segnala un difetto di
    un template ("Xref table not zero-indexed"). Li manda a un file vuoto, così i messaggi si perdono
    ma il programma continua."""
    for nome in ("stdout", "stderr"):
        if getattr(sys, nome) is None:
            setattr(sys, nome, open(os.devnull, "w", encoding="utf-8"))


def registra_font(cartella):
    """Rende i font della cartella (es. Fonts del programma) disponibili alle finestre del programma,
    senza installarli nel sistema: così l'editor mostra il testo con gli stessi caratteri del PDF.
    Restituisce quanti file sono stati registrati (0 se non si può: il programma va avanti lo stesso)."""
    if not os.path.isdir(cartella):
        return 0
    file_font = [os.path.join(cartella, f) for f in os.listdir(cartella) if f.lower().endswith((".ttf", ".otf"))]
    try:
        if WINDOWS:
            import ctypes
            # FR_PRIVATE: i font valgono solo per questo programma e spariscono quando si chiude
            aggiungi = ctypes.windll.gdi32.AddFontResourceExW
            return sum(1 for f in file_font if aggiungi(f, 0x10, None) > 0)
        if not MACOS:
            import ctypes
            import ctypes.util
            # fontconfig (usato da Tk su Linux): aggiunge la cartella per questo programma
            libreria = ctypes.util.find_library("fontconfig")
            if libreria and ctypes.CDLL(libreria).FcConfigAppFontAddDir(None, cartella.encode()):
                return len(file_font)
    except (OSError, AttributeError):
        pass
    return 0


def cartella_utente():
    """Cartella delle impostazioni dell'utente (modelli recenti, zoom...).

    Windows: %APPDATA%\\FileGUY. Linux: ~/.config/fileguy (o $XDG_CONFIG_HOME/fileguy), dove
    sta anche la chiave di config.dat. macOS: ~/Library/Application Support/FileGUY.
    """
    if WINDOWS:
        return os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), "FileGUY")
    if MACOS:
        return os.path.expanduser("~/Library/Application Support/FileGUY")
    return os.path.join(os.environ.get("XDG_CONFIG_HOME") or os.path.expanduser("~/.config"), "fileguy")


def leggi_testo(percorso):
    """Legge un file di testo creato su Windows o su Linux.

    Prova UTF-8 e, se non è valido, la codifica di Windows (cp1252), che era quella usata finora.
    Gli a capo diventano tutti \\n, come con open() in modalità testo.
    """
    with open(percorso, "rb") as f:
        dati = f.read()
    try:
        testo = dati.decode("utf-8-sig")
    except UnicodeDecodeError:
        testo = dati.decode("cp1252", errors="replace")
    return testo.replace("\r\n", "\n").replace("\r", "\n")


def leggi_righe(percorso):
    """Come f.readlines(): le righe con il \\n finale (solo \\n separa le righe, non il salto pagina)."""
    parti = leggi_testo(percorso).split("\n")
    righe = [parte + "\n" for parte in parti[:-1]]
    if parti[-1] != "":
        righe.append(parti[-1])
    return righe


def ambiente_esterno():
    """Variabili d'ambiente per avviare un programma esterno.

    L'eseguibile di PyInstaller su Linux mette le sue librerie in LD_LIBRARY_PATH: i programmi
    del sistema (xdg-open, lp, il visualizzatore PDF...) potrebbero non partire, quindi si
    rimette il valore originale.
    """
    ambiente = dict(os.environ)
    if getattr(sys, "frozen", False) and not WINDOWS:
        for variabile in ("LD_LIBRARY_PATH", "DYLD_LIBRARY_PATH"):
            originale = ambiente.pop(variabile + "_ORIG", None)
            if originale is not None:
                ambiente[variabile] = originale
            else:
                ambiente.pop(variabile, None)
    return ambiente


def apri_file(percorso):
    """Apre il file con il programma predefinito del sistema (es. il PDF nel visualizzatore)."""
    if WINDOWS:
        os.startfile(percorso)
        return
    if not MACOS and not (os.environ.get("DISPLAY") or os.environ.get("WAYLAND_DISPLAY")):
        raise RuntimeError("Nessuno schermo disponibile (DISPLAY/WAYLAND_DISPLAY non impostati): "
                           "impossibile aprire " + percorso)
    comando = "open" if MACOS else "xdg-open"
    try:
        # Non aspetta: il visualizzatore resta aperto per conto suo
        subprocess.Popen([comando, percorso], stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                         stderr=subprocess.DEVNULL, start_new_session=True, env=ambiente_esterno())
    except FileNotFoundError:
        raise RuntimeError(comando + " non trovato: installa il pacchetto xdg-utils della distribuzione")


_indice_font = None


def cerca_font(nome_file):
    """Cerca un file di font nelle cartelle del sistema; restituisce il percorso o None.

    Su Linux i font possono stare in sottocartelle e il nome può avere le maiuscole
    (es. i font Microsoft di ttf-mscorefonts si chiamano Verdana.ttf o verdana.ttf).
    """
    global _indice_font
    for cartella in CARTELLE_FONT_SISTEMA:
        percorso = os.path.join(cartella, nome_file)
        if os.path.isfile(percorso):
            return percorso
    if WINDOWS:
        return None
    if _indice_font is None:
        _indice_font = {}
        for cartella in CARTELLE_FONT_SISTEMA:
            for radice, _, files in os.walk(cartella):
                for f in files:
                    _indice_font.setdefault(f.lower(), os.path.join(radice, f))
    return _indice_font.get(nome_file.lower())
