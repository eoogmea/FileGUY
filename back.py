import time
import subprocess
from os import listdir
from os.path import isfile, join

#Controlla se è cambiato qualcosa nella cartella
#Se sì fa partire App.exe o App.py
def process(filesprec):
    # Apre il file di configurazione mette il contenuto in config
    f = open("config.txt", "r")
    config = f.readlines()
    f.close()

    # Cartella con i file di testo e i .gui
    folder = config[1].replace("\n", "")

    # Legge i file in questa cartella
    onlyfiles = [f for f in listdir(folder) if isfile(join(folder, f))]

    if filesprec != onlyfiles:
        try:
            subprocess.call("App.exe", shell=True)
        except:
            subprocess.call("App.py", shell=True)
    return onlyfiles

#Crea un array con i file precedenti nella cartella
filesprec = []

#loop che avviene ogni secondo
while True:
    time.sleep(1)
    filesprec = process(filesprec)


