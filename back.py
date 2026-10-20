import os
import sys
import time
import subprocess

# CONFIGURAZIONE CIFRATA
from config_sicuro import carica_config
from os import listdir
from os.path import isfile, join

# Cartella del programma (o dell'eseguibile creato con PyInstaller)
if getattr(sys, "frozen", False):
    CARTELLA_PROGRAMMA = os.path.dirname(sys.executable)
else:
    CARTELLA_PROGRAMMA = os.path.dirname(os.path.abspath(__file__))

#Controlla se è cambiato qualcosa nella cartella
#Se sì fa partire SingleGUY_editor.exe o SingleGUY_editor.py
def process(filesprec):
    # Legge la configurazione cifrata (config.dat, si modifica con ConfigManager.py)
    config = carica_config()

    # Cartella con i file di testo e i .gui
    folder = config["cartella_lavoro"]

    # Legge i file in questa cartella
    onlyfiles = [f for f in listdir(folder) if isfile(join(folder, f))]

    if filesprec != onlyfiles:
        # Usa l'eseguibile se c'è, altrimenti lo script con lo stesso Python di back.py
        # (subprocess.call con shell=True non dà errore se il file manca, quindi va controllato prima)
        eseguibile = join(CARTELLA_PROGRAMMA, "SingleGUY_editor.exe")
        if isfile(eseguibile):
            subprocess.call([eseguibile], cwd=CARTELLA_PROGRAMMA)
        else:
            subprocess.call([sys.executable, join(CARTELLA_PROGRAMMA, "SingleGUY_editor.py")],
                            cwd=CARTELLA_PROGRAMMA)
    return onlyfiles

#Crea un array con i file precedenti nella cartella
filesprec = []

#loop che avviene ogni secondo
while True:
    time.sleep(1)
    filesprec = process(filesprec)
