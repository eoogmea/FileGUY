import subprocess

# EMAIL
import email
import smtplib
import ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import os
from os import listdir
from os.path import isfile, join

import pandas as pd

import numpy as np

# INTERNET (TELEGRAM)
import requests
import telegram_upload
import telepot
#import cookie
from telepot.loop import MessageLoop
import pdb

# PDF
from PyPDF2 import PdfFileWriter, PdfFileReader
from fpdf import FPDF

# DOCX
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, Cm

# EXCEL
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

import time

bot_token = "1085822552:AAFeBpz7MPc1_Z-QtswOmOv25vLPzpDWr38"
TOKEN = bot_token
bot = telepot.Bot(TOKEN)
bot_chatID = 897989182
bot_message = "Ciao Fiocco"
send_text = 'https://api.telegram.org/bot' + str(bot_token) + '/sendMessage?chat_id=' + str(bot_chatID) + '&parse_mode=Markdown&text=' + bot_message
time.sleep(10)
while True:
    requests.get(send_text)
    requests.get(send_text)
    requests.get(send_text)
    bot.sendMessage(bot_chatID, bot_message)
    requests.get(send_text)
    requests.get(send_text)
    requests.get(send_text)