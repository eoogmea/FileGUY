import os
from os import listdir
from os.path import isfile, join
import requests


def main(name,config):
    name_gui = name.replace(".txt", "") + ".gui"
    path_gui = config[1].replace("\n", "") + "\\" + name_gui
    if not os.path.isfile(path_gui):
        return
    f = open(path_gui, "r")
    guida = f.readlines()
    f.close()
    telegram(guida, config)
f = open("config.txt", "r")
config = f.readlines()
f.close()
folder = config[1].replace("\n", "")
onlyfiles = [f for f in listdir(folder) if isfile(join(folder, f))]
for name in onlyfiles:
    main(name, config)
def telegram(guida, config):
    bot_token = config[6].replace("\n", "")
    # https://api.telegram.org/bot<bot_token>/getUpdates
    bot_message = guida[6].replace("\n", "")
    bot_chatID = guida[5].replace("\n", "").replace("telegram ", "")
    send_text = 'https://api.telegram.org/bot' + bot_token + '/sendMessage?chat_id=' + bot_chatID + '&text=' + bot_message
    response = requests.get(send_text)
    document = config[3].replace("\n", "") + "\\" + guida[2].replace("\n", "")
    send_text = 'https://api.telegram.org/bot' + bot_token + '/sendDocument?chat_id=' + bot_chatID + '&document=' + document
    response = requests.get(send_text)

