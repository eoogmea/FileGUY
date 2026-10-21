"""Registro (log) delle operazioni del servizio.

Ogni riga: data e ora | nome del file | azione | esito [| dettagli]
    21/10/2026 10:15:03 | fattura_123.txt | pdf fattura_123.pdf su template fat.pdf | OK
    21/10/2026 10:15:04 | fattura_123.txt | email a cliente@esempio.it | ERRORE | SMTPAuthenticationError: ...
Dopo una riga di ERRORE segue il dettaglio tecnico (traceback), rientrato di 4 spazi.

Il percorso si imposta in ConfigManager (campo "File di log"); vuoto = FileGUY.log
nella cartella del programma. Superati 5 MB il file viene rinominato in .1 (poi .2, .3)
e se ne inizia uno nuovo.
Scrivere il registro non blocca mai il programma: se il file non è scrivibile
l'errore viene solo mostrato sulla console.
"""
import datetime
import os
import sys
import traceback

from config_sicuro import CARTELLA_PROGRAMMA

FILE_PREDEFINITO = os.path.join(CARTELLA_PROGRAMMA, "FileGUY.log")
DIMENSIONE_MASSIMA = 5 * 1024 * 1024
COPIE = 3

# None = registro non attivo (es. l'anteprima di EditorModelli, che usa le funzioni del servizio)
_percorso = None


def configura(percorso=""):
    """Attiva il registro sul file indicato (vuoto = quello predefinito)."""
    global _percorso
    _percorso = percorso.strip() or FILE_PREDEFINITO


def _ruota():
    if not os.path.isfile(_percorso) or os.path.getsize(_percorso) < DIMENSIONE_MASSIMA:
        return
    for n in range(COPIE - 1, 0, -1):
        vecchio = "%s.%d" % (_percorso, n)
        if os.path.isfile(vecchio):
            os.replace(vecchio, "%s.%d" % (_percorso, n + 1))
    os.replace(_percorso, _percorso + ".1")


def _scrivi(righe):
    if _percorso is None:
        return
    try:
        cartella = os.path.dirname(os.path.abspath(_percorso))
        os.makedirs(cartella, exist_ok=True)
        _ruota()
        # Aperto e chiuso a ogni scrittura: back.py e SingleGUY scrivono sullo stesso file
        with open(_percorso, "a", encoding="utf-8") as f:
            f.write("".join(r + "\n" for r in righe))
    except Exception as e:
        # Qualsiasi problema (permessi, disco pieno, percorso non valido): il programma continua
        print("Impossibile scrivere il registro " + repr(_percorso) + ": " + str(e), file=sys.stderr)


def registra(nome_file, azione, esito="OK", dettagli=""):
    campi = [datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"), nome_file or "-", azione, esito]
    if dettagli:
        campi.append(dettagli)
    _scrivi([" | ".join(c.replace("\n", " ") for c in campi)])


def registra_errore(nome_file, azione, errore):
    """Registra un'eccezione: la riga di ERRORE e sotto il traceback."""
    registra(nome_file, azione, "ERRORE", type(errore).__name__ + ": " + str(errore))
    dettaglio = traceback.format_exception(type(errore), errore, errore.__traceback__)
    _scrivi(["    " + riga for riga in "".join(dettaglio).rstrip().split("\n")])
