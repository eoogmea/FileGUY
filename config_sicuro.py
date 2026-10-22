"""Configurazione cifrata di FileGUY.

Sostituisce il vecchio config.txt in chiaro con config.dat. Il file può essere letto solo dallo
stesso utente (o, se scelto, da tutti gli utenti dello stesso PC) che lo ha salvato.
Non serve nessuna password all'avvio, quindi il servizio può partire da solo.
Per creare o modificare la configurazione si usa ConfigManager.py.

Windows: cifrato con DPAPI (la chiave è protetta da Windows con la password dell'utente).
Linux e macOS: cifrato con AES-GCM (libreria cryptography). La chiave è
- solo questo utente: un file casuale in ~/.config/fileguy/config.key leggibile solo
  dall'utente (permessi 600, cartella 700)
- tutti gli utenti del PC: ricavata dall'identificativo della macchina (/etc/machine-id)
Un config.dat creato su Windows non si legge su Linux e viceversa: va ricreato con ConfigManager.
"""
import ctypes
import hashlib
import json
import os
import sys

# Il file sta nella cartella del programma (o dell'eseguibile creato con PyInstaller)
if getattr(sys, "frozen", False):
    CARTELLA_PROGRAMMA = os.path.dirname(sys.executable)
else:
    CARTELLA_PROGRAMMA = os.path.dirname(os.path.abspath(__file__))
FILE_CONFIG = os.path.join(CARTELLA_PROGRAMMA, "config.dat")
FILE_CONFIG_VECCHIO = os.path.join(CARTELLA_PROGRAMMA, "config.txt")

# Campi della configurazione con il loro valore predefinito
CAMPI = {
    "cartella_lavoro": "",
    "cartella_template": "",
    "cartella_output": "",
    "email_mittente": "",
    "email_password": "",
    "email_smtp_server": "smtp.gmail.com",
    "email_smtp_porta": "465",
    "telegram_token": "",
    "whatsapp_token": "",
    "whatsapp_phone_number_id": "",
    "whatsapp_versione_api": "v23.0",
    "whatsapp_lingua": "it",
    # Vuoto = ricavato dal token (serve per l'elenco e la creazione dei template)
    "whatsapp_account_business": "",
    # Vuoto = stampante predefinita del sistema
    "stampante": "",
    # Vuoto = FileGUY.log nella cartella del programma
    "file_log": "",
}

# Posizione dei campi nel vecchio config.txt (la riga 0 è l'intestazione)
RIGHE_CONFIG_VECCHIO = {
    1: "cartella_lavoro",
    2: "cartella_template",
    3: "cartella_output",
    4: "email_mittente",
    5: "email_password",
    6: "telegram_token",
}

# Versione 1: DPAPI di Windows. Versione 2: AES-GCM (Linux, macOS)
_INTESTAZIONE = b"INFASCFG\x01"
_INTESTAZIONE_POSIX = b"INFASCFG\x02"
# Entropia aggiuntiva: anche un altro programma dello stesso utente deve conoscerla per decifrare
# Valore storico (il programma si chiamava InfasService): se cambia, i config.dat esistenti non si leggono più
_ENTROPIA = b"InfasService-config-v1"

_CRYPTPROTECT_UI_FORBIDDEN = 0x01
_CRYPTPROTECT_LOCAL_MACHINE = 0x04

# Modo della chiave su Linux (byte dopo l'intestazione)
_SOLO_UTENTE = b"U"
_TUTTO_IL_PC = b"M"
_FILE_MACHINE_ID = ["/etc/machine-id", "/var/lib/dbus/machine-id"]


class ConfigError(Exception):
    pass


# DPAPI (crypt32.dll) tramite ctypes, così non servono librerie esterne
if os.name == "nt":
    from ctypes import wintypes

    class _DATA_BLOB(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]


def _blob(dati):
    buffer = ctypes.create_string_buffer(dati, len(dati))
    return _DATA_BLOB(len(dati), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_char))), buffer


def _dpapi(cifra, dati, flags=0):
    crypt32 = ctypes.windll.crypt32
    kernel32 = ctypes.windll.kernel32
    kernel32.LocalFree.argtypes = [ctypes.c_void_p]
    kernel32.LocalFree.restype = ctypes.c_void_p

    blob_in, _buf_in = _blob(dati)
    blob_entropia, _buf_entropia = _blob(_ENTROPIA)
    blob_out = _DATA_BLOB()
    flags |= _CRYPTPROTECT_UI_FORBIDDEN
    if cifra:
        ok = crypt32.CryptProtectData(ctypes.byref(blob_in), "FileGUY", ctypes.byref(blob_entropia),
                                      None, None, flags, ctypes.byref(blob_out))
    else:
        ok = crypt32.CryptUnprotectData(ctypes.byref(blob_in), None, ctypes.byref(blob_entropia),
                                        None, None, flags, ctypes.byref(blob_out))
    if not ok:
        raise ctypes.WinError()
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        kernel32.LocalFree(ctypes.cast(blob_out.pbData, ctypes.c_void_p))


# --- Linux e macOS -----------------------------------------------------------

def file_chiave():
    """Chiave dell'utente: ~/.config/fileguy/config.key (o $XDG_CONFIG_HOME)."""
    base = os.environ.get("XDG_CONFIG_HOME") or os.path.join(os.path.expanduser("~"), ".config")
    return os.path.join(base, "fileguy", "config.key")


def _chiave_utente(crea):
    percorso = file_chiave()
    if os.path.isfile(percorso):
        with open(percorso, "rb") as f:
            chiave = f.read()
        if len(chiave) != 32:
            raise ConfigError("Il file della chiave " + percorso + " è danneggiato")
        return chiave
    if not crea:
        raise ConfigError("Chiave " + percorso + " non trovata: config.dat è stato creato da un altro utente "
                          "o su un altro PC. Ricrealo con ConfigManager.py")
    os.makedirs(os.path.dirname(percorso), mode=0o700, exist_ok=True)
    chiave = os.urandom(32)
    # Creato direttamente con i permessi 600: non esiste mai un momento in cui altri possono leggerlo
    descrittore = os.open(percorso, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
    with os.fdopen(descrittore, "wb") as f:
        f.write(chiave)
    return chiave


def _chiave_pc():
    for percorso in _FILE_MACHINE_ID:
        try:
            with open(percorso, "rb") as f:
                identificativo = f.read().strip()
        except OSError:
            continue
        if identificativo:
            return hashlib.sha256(_ENTROPIA + b"|" + identificativo).digest()
    raise ConfigError("Identificativo della macchina non trovato (" + ", ".join(_FILE_MACHINE_ID)
                      + "): usa la configurazione per il solo utente corrente")


def _aesgcm(chiave):
    try:
        from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    except ImportError:
        raise ConfigError("Manca la libreria cryptography: installala con  pip install cryptography")
    return AESGCM(chiave)


def _cifra_posix(dati, tutto_il_pc):
    modo = _TUTTO_IL_PC if tutto_il_pc else _SOLO_UTENTE
    chiave = _chiave_pc() if tutto_il_pc else _chiave_utente(crea=True)
    nonce = os.urandom(12)
    intestazione = _INTESTAZIONE_POSIX + modo
    return intestazione + nonce + _aesgcm(chiave).encrypt(nonce, dati, intestazione)


def _decifra_posix(contenuto, percorso):
    intestazione = contenuto[:len(_INTESTAZIONE_POSIX) + 1]
    modo = intestazione[-1:]
    chiave = _chiave_pc() if modo == _TUTTO_IL_PC else _chiave_utente(crea=False)
    nonce = contenuto[len(intestazione):len(intestazione) + 12]
    try:
        from cryptography.exceptions import InvalidTag
        return _aesgcm(chiave).decrypt(nonce, contenuto[len(intestazione) + 12:], intestazione)
    except (InvalidTag, ValueError):
        raise ConfigError("Impossibile decifrare " + percorso + ": è stato creato da un altro utente "
                          "o su un altro PC. Ricrealo con ConfigManager.py")


# --- Lettura e scrittura -----------------------------------------------------

def config_esiste(percorso=FILE_CONFIG):
    return os.path.isfile(percorso)


def carica_config(percorso=FILE_CONFIG):
    """Legge e decifra config.dat, restituisce un dizionario con tutti i CAMPI."""
    if not os.path.isfile(percorso):
        raise ConfigError("Configurazione non trovata: " + percorso
                          + "\nCreala con ConfigManager.py")
    with open(percorso, "rb") as f:
        contenuto = f.read()
    if contenuto.startswith(_INTESTAZIONE):
        if os.name != "nt":
            raise ConfigError("Il file " + percorso + " è stato creato su Windows e non si può leggere qui: "
                              "ricrealo con ConfigManager.py")
        try:
            chiaro = _dpapi(False, contenuto[len(_INTESTAZIONE):])
        except OSError:
            raise ConfigError("Impossibile decifrare " + percorso + ": è stato creato da un altro utente "
                              "Windows o su un altro PC. Ricrealo con ConfigManager.py")
    elif contenuto.startswith(_INTESTAZIONE_POSIX):
        if os.name == "nt":
            raise ConfigError("Il file " + percorso + " è stato creato su Linux e non si può leggere qui: "
                              "ricrealo con ConfigManager.py")
        chiaro = _decifra_posix(contenuto, percorso)
    else:
        raise ConfigError("Il file " + percorso + " non è una configurazione valida")
    dati = json.loads(chiaro.decode("utf-8"))
    config = dict(CAMPI)
    config.update({k: str(v) for k, v in dati.items() if k in CAMPI})
    return config


def salva_config(config, percorso=FILE_CONFIG, tutto_il_pc=False):
    """Cifra e salva la configurazione.

    tutto_il_pc=False: solo l'utente corrente può leggerla.
    tutto_il_pc=True: qualsiasi utente di questo PC può leggerla (serve se il servizio
    gira con un utente diverso da chi la configura, es. come servizio di sistema).
    """
    dati = {k: str(config.get(k, v)) for k, v in CAMPI.items()}
    chiaro = json.dumps(dati).encode("utf-8")
    if os.name == "nt":
        flags = _CRYPTPROTECT_LOCAL_MACHINE if tutto_il_pc else 0
        contenuto = _INTESTAZIONE + _dpapi(True, chiaro, flags)
    else:
        contenuto = _cifra_posix(chiaro, tutto_il_pc)
    # Scrive su un file temporaneo e poi lo sostituisce, così non resta mai un file a metà
    temporaneo = percorso + ".tmp"
    with open(temporaneo, "wb") as f:
        f.write(contenuto)
    if os.name != "nt":
        # Solo il proprietario lo modifica; con "tutto il PC" deve poterlo leggere anche il servizio
        os.chmod(temporaneo, 0o644 if tutto_il_pc else 0o600)
    os.replace(temporaneo, percorso)


def importa_config_vecchio(percorso=FILE_CONFIG_VECCHIO):
    """Legge il vecchio config.txt in chiaro e lo converte nel nuovo formato."""
    from piattaforma import leggi_testo
    righe = leggi_testo(percorso).splitlines()
    config = dict(CAMPI)
    for indice, campo in RIGHE_CONFIG_VECCHIO.items():
        if indice < len(righe):
            config[campo] = righe[indice].strip()
    return config
