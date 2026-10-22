import os
import sys
import time
import subprocess

# CONFIGURAZIONE CIFRATA
from config_sicuro import carica_config
from os import listdir
from os.path import isfile, join

# WINDOWS / LINUX E REGISTRO (LOG)
from piattaforma import WINDOWS, ambiente_esterno
import registro

# Cartella del programma (o dell'eseguibile creato con PyInstaller)
if getattr(sys, "frozen", False):
    CARTELLA_PROGRAMMA = os.path.dirname(sys.executable)
else:
    CARTELLA_PROGRAMMA = os.path.dirname(os.path.abspath(__file__))
# L'eseguibile creato con PyInstaller su Linux non ha estensione
NOME_ESEGUIBILE = "SingleGUY.exe" if WINDOWS else "SingleGUY"

#Controlla se è cambiato qualcosa nella cartella
#Se sì fa partire SingleGUY(.exe) o SingleGUY.py
def process(filesprec):
    # Legge la configurazione cifrata (config.dat, si modifica con ConfigManager.py)
    config = carica_config()
    registro.configura(config["file_log"])

    # Cartella con i file di testo e i .gui
    folder = config["cartella_lavoro"]

    # Legge i file in questa cartella
    onlyfiles = [f for f in listdir(folder) if isfile(join(folder, f))]

    if filesprec != onlyfiles:
        # Usa l'eseguibile se c'è, altrimenti lo script con lo stesso Python di back.py
        # (subprocess.call con shell=True non dà errore se il file manca, quindi va controllato prima)
        eseguibile = join(CARTELLA_PROGRAMMA, NOME_ESEGUIBILE)
        if isfile(eseguibile):
            subprocess.call([eseguibile], cwd=CARTELLA_PROGRAMMA, env=ambiente_esterno())
        else:
            subprocess.call([sys.executable, join(CARTELLA_PROGRAMMA, "SingleGUY.py")],
                            cwd=CARTELLA_PROGRAMMA)
    return onlyfiles

#Crea un array con i file precedenti nella cartella
filesprec = []
#Finché non è letta la configurazione il log va nel file predefinito
registro.configura()
registro.registra("", "avvio del servizio (back)")
ultimo_errore = None

#loop che avviene ogni secondo
try:
    while True:
        time.sleep(1)
        #Un errore (configurazione o cartella mancante...) non ferma il servizio: va nel log una volta
        #sola finché non cambia, e si riprova al giro dopo
        try:
            filesprec = process(filesprec)
            ultimo_errore = None
        except Exception as e:
            if str(e) != ultimo_errore:
                registro.registra_errore("", "controllo della cartella di lavoro", e)
                print("Errore: " + str(e), file=sys.stderr)
            ultimo_errore = str(e)
except KeyboardInterrupt:
    #Ctrl+C: il servizio si chiude senza traceback
    registro.registra("", "arresto del servizio (back)")
