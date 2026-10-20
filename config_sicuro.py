"""Configurazione cifrata di InfasService.

Sostituisce il vecchio config.txt in chiaro con config.dat, cifrato con DPAPI di Windows:
il file può essere letto solo dallo stesso utente Windows (o, se scelto, dallo stesso PC)
che lo ha salvato. Non serve nessuna password all'avvio, quindi il servizio può partire da solo.
Per creare o modificare la configurazione si usa ConfigManager.py.
"""
import ctypes
import json
import os
import sys
from ctypes import wintypes

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

_INTESTAZIONE = b"INFASCFG\x01"
# Entropia aggiuntiva: anche un altro programma dello stesso utente deve conoscerla per decifrare
_ENTROPIA = b"InfasService-config-v1"

_CRYPTPROTECT_UI_FORBIDDEN = 0x01
_CRYPTPROTECT_LOCAL_MACHINE = 0x04


class ConfigError(Exception):
    pass


# DPAPI (crypt32.dll) tramite ctypes, così non servono librerie esterne
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
        ok = crypt32.CryptProtectData(ctypes.byref(blob_in), "InfasService", ctypes.byref(blob_entropia),
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


def config_esiste(percorso=FILE_CONFIG):
    return os.path.isfile(percorso)


def carica_config(percorso=FILE_CONFIG):
    """Legge e decifra config.dat, restituisce un dizionario con tutti i CAMPI."""
    if not os.path.isfile(percorso):
        raise ConfigError("Configurazione non trovata: " + percorso
                          + "\nCreala con ConfigManager.py")
    with open(percorso, "rb") as f:
        contenuto = f.read()
    if not contenuto.startswith(_INTESTAZIONE):
        raise ConfigError("Il file " + percorso + " non è una configurazione valida")
    try:
        dati = json.loads(_dpapi(False, contenuto[len(_INTESTAZIONE):]).decode("utf-8"))
    except OSError:
        raise ConfigError("Impossibile decifrare " + percorso + ": è stato creato da un altro utente "
                          "Windows o su un altro PC. Ricrealo con ConfigManager.py")
    config = dict(CAMPI)
    config.update({k: str(v) for k, v in dati.items() if k in CAMPI})
    return config


def salva_config(config, percorso=FILE_CONFIG, tutto_il_pc=False):
    """Cifra e salva la configurazione.

    tutto_il_pc=False: solo l'utente Windows corrente può leggerla.
    tutto_il_pc=True: qualsiasi utente di questo PC può leggerla (serve se il servizio
    gira con un utente diverso da chi la configura, es. come servizio di sistema).
    """
    dati = {k: str(config.get(k, v)) for k, v in CAMPI.items()}
    flags = _CRYPTPROTECT_LOCAL_MACHINE if tutto_il_pc else 0
    cifrato = _dpapi(True, json.dumps(dati).encode("utf-8"), flags)
    # Scrive su un file temporaneo e poi lo sostituisce, così non resta mai un file a metà
    temporaneo = percorso + ".tmp"
    with open(temporaneo, "wb") as f:
        f.write(_INTESTAZIONE + cifrato)
    os.replace(temporaneo, percorso)


def importa_config_vecchio(percorso=FILE_CONFIG_VECCHIO):
    """Legge il vecchio config.txt in chiaro e lo converte nel nuovo formato."""
    with open(percorso, "r") as f:
        righe = f.read().splitlines()
    config = dict(CAMPI)
    for indice, campo in RIGHE_CONFIG_VECCHIO.items():
        if indice < len(righe):
            config[campo] = righe[indice].strip()
    return config
