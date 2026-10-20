import os
from os import listdir
from os.path import isfile, join
# SUBPROCESS (EXE)
import subprocess

# CONFIGURAZIONE CIFRATA
from config_sicuro import carica_config

# EMAIL
import email
import smtplib
import ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd

# INTERNET (TELEGRAM)
import requests
import telepot

#import cookie

# WHATSAPP (API Cloud di Meta)
from whatsapp_meta import WhatsApp, WhatsAppError

# PDF
from PyPDF2 import PdfFileWriter, PdfFileReader
from fpdf import FPDF
# Stampa senza Ghostscript (PyMuPDF + stampante di Windows)
from stampa_pdf import stampa_pdf

# DOCX
from docx import Document
from docx.shared import RGBColor
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, Cm

# EXCEL
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

#BARCODE
import barcode
from barcode.writer import ImageWriter
import fitz

import time

def process(filesprec):
    # Legge la configurazione cifrata (config.dat, si modifica con ConfigManager.py)
    config = carica_config()

    # Cartella con i file di testo e i .gui
    folder = config["cartella_lavoro"]

    # Legge i file in questa cartella
    onlyfiles = [f for f in listdir(folder) if isfile(join(folder, f))]

    if filesprec != onlyfiles:
        start()
    return onlyfiles

def main(name, config):
    # Sostituisce al file .txt la terminazione .gui per vedere se esiste
    name_gui = name.replace(".txt", ".gui")
    path_gui = config["cartella_lavoro"] + "\\" + name_gui
    if name_gui[len(name_gui) - 3:len(name_gui)] != "gui":
        return
    if not os.path.isfile(path_gui):
        return
    # Apre il file .gui
    f = open(path_gui, "r")
    guida = f.readlines()
    f.close()
    if guida[1].replace("\n", "") == "exe":
        exe(guida, config)
        # Rimuove i file dopo l'uso
        os.remove(path_gui)
    elif guida[1].replace("\n", "") == "barcode":
        barcode_add(guida, config)
        os.remove(path_gui)
    elif guida[1].replace("\n", "") == "immagini":
        image_add(guida, config)
        os.remove(path_gui)

    if name[-3:] == "gui":
        return

    # Legge la seconda linea del .gui per decidere cosa fare
    if guida[1].replace("\n", "") == "pdf" or guida[1].replace("\n", "") == "file" \
            or guida[1].replace("\n", "") == "stampa":
        pdf(guida, config)
        if guida[5].replace("\n", "") == "email":
            email_send(guida, config)
        if guida[5].replace("\n", "") == "telegram":
            telegram(guida, config)
        if guida[5].replace("\n", "") == "whatsapp":
            whatsapp(guida, config)
    if guida[1].replace("\n", "") == "excel":
        excel(guida, config)
        if guida[5].replace("\n", "") == "email":
            email_send(guida, config)
        if guida[5].replace("\n", "") == "telegram":
            telegram(guida, config)
        if guida[5].replace("\n", "") == "whatsapp":
            whatsapp(guida, config)
    if guida[1].replace("\n", "") == "docx":
        docx(guida, config)
        if guida[5].replace("\n", "") == "email":
            email_send(guida, config)
        if guida[5].replace("\n", "") == "telegram":
            telegram(guida, config)
        if guida[5].replace("\n", "") == "whatsapp":
            whatsapp(guida, config)

    # Rimuove i file dopo l'uso
    os.remove(path_gui)
    os.remove(config["cartella_lavoro"] + "\\" + guida[0].replace("\n", ""))


def barcode_add(guida, config):
    numero_barcodes = int(guida[2].replace("\n", ""))
    for i in range(0, numero_barcodes):
        i *= 6
        codice = guida[3 + i].replace("\n", "")
        cab = barcode.get_barcode_class(guida[4 + i].replace("\n", ""))
        codicino = cab(codice, writer=ImageWriter())
        percorso = config["cartella_lavoro"] + "\\" + "barcode" + str(codice)
        codicino.save(percorso)

        # la posizione è data da due vertici opposti partendo da quello alto a sx
        x0, y0 = guida[5 + i].replace("\n", "").split(",")
        x1, y1 = guida[6 + i].replace("\n", "").split(",")
        image_rectangle = fitz.Rect(int(x0), int(y0), int(x1), int(y1))

        # retrieve the page of the PDF
        doc = fitz.open(guida[7 + i].replace("\n", ""))
        page = doc[int(guida[8 + i].replace("\n", ""))]

        # add the image
        page.insert_image(image_rectangle, filename=percorso + ".png")
        doc.save(config["cartella_lavoro"] + "\\" + guida[7 + i].replace("\n", "").split("\\")[-1], incremental=True, encryption=0)

def image_add(guida,config):
    numero_image = int(guida[2].replace("\n", ""))
    for i in range(0, numero_image):
        i *= 5
        # la posizione è data da due vertici opposti partendo da quello alto a sx
        percorso = guida[3 + i].replace("\n", "")
        x0, y0 = guida[4 + i].replace("\n", "").split(",")
        x1, y1 = guida[5 + i].replace("\n", "").split(",")
        image_rectangle = fitz.Rect(int(x0), int(y0), int(x0) + 100, int(y0) + 100)

        # retrieve the page of the PDF
        doc = fitz.open(guida[6 + i].replace("\n", ""))
        page = doc[int(guida[7 + i].replace("\n", ""))]

        # add the image
        page.insert_image(image_rectangle, filename= percorso)
        doc.save(config["cartella_lavoro"] + "\\" + guida[6 + i].replace("\n", "").split("\\")[-1], incremental=True, encryption=0)



def exe(guida, config) :

    path = config["cartella_lavoro"] + "\\" + guida[0].replace("\n", "")
    if not os.path.isfile(path):
        path = guida[0].replace("\n", "")
    if not os.path.isfile(path):
        return
    try:
        os.system(path)
    except:
        subprocess.call(path)

def telegram(guida, config):
    #Legge il token del bot
    #Per creare un bot si usa TheBotFather su telegram
    bot_token = config["telegram_token"]
    #Quando si manda un messaggio al bot da queso link se ne ottengono i dati
    # https://api.telegram.org/bot<bot_token>/getUpdates
    i = 0
    body = ""
    document = ""
    #argomenti della libreria telepot
    TOKEN = bot_token
    bot = telepot.Bot(TOKEN)

    #legge le ultime linee del file dopo "fine"
    #la prima linea è l'url del documento
    #la seconda e seguenti sono il testo del messaggio

    for x in guida:
        if i == 1 or i == 2:
            if i == 1:
                document = x.replace("\n", "")
            i += 1
        if i >= 3:
            body = body + " " + x
        if x[0:4] == "fine":
            i += 1
    bot_message = body
    bot_chatID = guida[6].replace("\n", "").replace("telegram ", "")
    doc_path = config["cartella_output"] + "\\" + document

    #manda il messaggio
    bot.sendMessage(bot_chatID, bot_message)
    #se esiste manda il documento
    if document != "":
        bot.sendDocument(chat_id=bot_chatID, document=open(doc_path, 'rb'))

def whatsapp(guida, config):
    #Usa le API Cloud di Meta: token e Phone number ID si impostano con ConfigManager.py
    #La settima linea del .gui è "whatsapp <numero> [<template> [<lingua>]]"
    #il numero va con il prefisso internazionale (es. 393331234567 o +39 333 1234567)
    #legge le ultime linee del file dopo "fine"
    #la prima linea è il nome del documento nella cartella di output (vuota = nessun documento)
    #la seconda e seguenti sono il testo del messaggio
    #Senza template il messaggio arriva solo se il destinatario ha scritto al numero
    #aziendale nelle ultime 24 ore (regola di Meta).
    #Con il template le linee di testo diventano i parametri {{1}}, {{2}}, ... del template
    #e il documento diventa l'intestazione (il template deve avere un'intestazione Documento)
    parti = guida[6].replace("\n", "").split()[1:]
    #il numero può essere scritto con spazi: prende tutte le parti fatte solo di cifre e +-./()
    numero = ""
    while len(parti) > 0 and parti[0].strip("+-./()0123456789") == "":
        numero += parti.pop(0)
    if numero == "":
        print("WhatsApp: numero del destinatario mancante nella linea 7 del .gui")
        return
    template = parti[0] if len(parti) > 0 else ""
    lingua = parti[1] if len(parti) > 1 else config["whatsapp_lingua"]

    #cerca la linea "fine" e prende le linee successive
    fine = len(guida)
    for n in range(linea_inizio_format_pdf, len(guida)):
        if guida[n][0:4] == "fine":
            fine = n
            break
    coda = [x.replace("\n", "") for x in guida[fine + 1:]]
    document = coda[0].strip() if len(coda) > 0 else ""
    righe = coda[1:]
    doc_path = config["cartella_output"] + "\\" + document if document != "" else None

    try:
        messenger = WhatsApp(config["whatsapp_token"], config["whatsapp_phone_number_id"],
                             config["whatsapp_versione_api"])
        if template != "":
            parametri = [r for r in righe if r.strip() != ""]
            messenger.invia_template(numero, template, lingua, parametri, doc_path)
        else:
            testo = "\n".join(righe).strip()
            #manda il messaggio
            if testo != "":
                messenger.invia_testo(numero, testo)
            #se esiste manda il documento
            if doc_path is not None:
                messenger.invia_documento(numero, doc_path)
    except (WhatsAppError, requests.RequestException, OSError) as e:
        print("Errore invio WhatsApp a " + numero + ": " + str(e))

def pdf(guida, config):

    #Apre il documento di testo
    f = open(config["cartella_lavoro"] + "\\" + guida[0].replace("\n", ""), 'r')
    content = f.readlines()
    f.close()

    #Immagazzina la rotazione(eventuale) da applicare alla pagina
    rotazione = int(guida[4])

    #Dichiara il documento
    pdf = FPDF()

    #Aggiunge i Font
    #Se non sono nella cartella dell'eseguibile li va a cercare in Windows/Font
    try:
        pdf.add_font('Times', '', 'Font\\times.ttf', uni=True)
        pdf.add_font('TimesBold', '', 'Font\\timesbd.ttf', uni=True)
        pdf.add_font('TimesBoldIt', '', 'Font\\timesbi.ttf', uni=True)
        pdf.add_font('TimesIt', '', 'Font\\timesi.ttf', uni=True)
    except:
        pdf.add_font('Times', '', 'C:\Windows\Fonts\\times.ttf', uni=True)
        pdf.add_font('TimesBold', '', 'C:\Windows\Fonts\\timesbd.ttf', uni=True)
        pdf.add_font('TimesBoldIt', '', 'C:\Windows\Fonts\\timesbi.ttf', uni=True)
        pdf.add_font('TimesIt', '', 'C:\Windows\Fonts\\timesi.ttf', uni=True)
    try:
        pdf.add_font('Helvetica', '', 'Font\\Helvetica.ttf', uni=True)
        pdf.add_font('HelveticaBold', '', 'Font\\Helvetica-Bold.ttf', uni=True)
        pdf.add_font('HelveticaBoldIt', '', 'Font\\Helvetica-BoldOblique.ttf', uni=True)
        pdf.add_font('HelveticaIt', '', 'Font\\Helvetica-Oblique.ttf', uni=True)
    except:
        pdf.add_font('Times', '', 'C:\Windows\Fonts\\times.ttf', uni=True)
        pdf.add_font('TimesBold', '', 'C:\Windows\Fonts\\timesbd.ttf', uni=True)
        pdf.add_font('TimesBoldIt', '', 'C:\Windows\Fonts\\timesbi.ttf', uni=True)
        pdf.add_font('TimesIt', '', 'C:\Windows\Fonts\\timesi.ttf', uni=True)
    try:
        pdf.add_font('Courier', '', 'Font\\cour.ttf', uni=True)
        pdf.add_font('CourierBold', '', 'Font\\courbd.ttf', uni=True)
        pdf.add_font('CourierBoldIt', '', 'Font\\courbi.ttf', uni=True)
        pdf.add_font('CourierIt', '', 'Font\\couri.ttf', uni=True)
    except:
        pdf.add_font('Courier', '', 'C:\Windows\Fonts\\cour.ttf', uni=True)
        pdf.add_font('CourierBold', '', 'C:\Windows\Fonts\\courbd.ttf', uni=True)
        pdf.add_font('CourierBoldIt', '', 'C:\Windows\Fonts\\courbi.ttf', uni=True)
        pdf.add_font('CourierIt', '', 'C:\Windows\Fonts\\couri.ttf', uni=True)

    #Aggiunge la prima pagina al pdf
    pdf.alias_nb_pages()
    pdf.add_page()

    i = linea_inizio_format_pdf
    index = 0
    testopasprec = ""
    testopass = ""
    codice = "tnn10"
    #For praticamente identico a quello della funzione word
    #vedi nota lì
    for x in content:
        arrcodici = []
        riga = x
        car1prec = 0
        if guida[i][0:4] == "fine":
            i = linea_inizio_format_pdf
            index = 0

        codicecompleto = guida[i].replace("\n", "")[4:]
        if guida[i][10:12] != "":
            lung = len(guida[i].replace("\n", "")) - 3
            times = int(lung/12)
            for h in range(0, times):
                primochar = 12*h
                ultimochar = 12*h + 11
                arrcodici.append(codicecompleto[primochar: ultimochar])
        else:
            arrcodici.append(codicecompleto)

        index += 1
        testo_tot = ""
        def controllacodice(codice):
                font = ""
                if codice[0:1].lower()  == "h":
                    if codice[1:2].lower()  == "s":
                        if codice[2:3].lower()  == "s":
                            font = "HelveticaBoldIt"
                        else:
                            font = "HelveticaBold"
                    else:
                        if codice[2:3].lower()  == "s":
                            font = "HelveticaIt"
                        else:
                            font = "Helvetica"
                if codice[0:1].lower()  == "t":
                    if codice[1:2].lower()  == "s":
                        if codice[2:3].lower()  == "s":
                            font = "TimesBoldIt"
                        else:
                            font = "TimesBold"
                    else:
                        if codice[2:3] == "s":
                            font = "TimesIt"
                        else:
                            font = "Times"
                if codice[0:1].lower() == "c":
                    if codice[1:2].lower()  == "s":
                        if codice[2:3].lower()  == "s":
                            font = "CourierBoldIt"
                        else:
                            font = "CourierBold"
                    else:
                        if codice[2:3] == "s":
                            font = "CourierIt"
                        else:
                            font = "Courier"
                size = codice[-2:]
                array = [font, size]
                return array

        if arrcodici[0] != "" and arrcodici[0] != " ":
            for codici in arrcodici:
                codice2 = codici[0:5]
                if codici[6:] != "":
                    car1 = int(codici[6:9])
                    car2 = int(codici[9:])
                    testopass = riga[car1: car2]
                    testo_tot += testopass
                    if codici == arrcodici[0]:
                        testopasprec = riga[0: car1]
                        testo_tot += testopasprec
                    else:
                        testopasprec = riga[car1prec: car1]
                        testo_tot += testopasprec
                    car1prec = car2
                else:
                    testopass = riga
                    testo_tot += testopass
                if testopasprec != "":
                    testoarr = controllacodice(codice)
                    codice = guida[i][4:9]
                    testo = testopasprec.replace("\n", "")
                    pdf.set_font(family=testoarr[0], size=int(testoarr[1]))
                    pdf.cell(len(testo) + 6, 6, testo, border=0, ln=0)
                testoarr2 = controllacodice(codice2)
                testo = testopass.replace("\n", "")
                pdf.set_font(family=testoarr2[0], size=int(testoarr2[1]))
                if codici == arrcodici[len(arrcodici) - 1]:
                    pdf.cell(len(testo) + 6, 6, testo, border=0, ln=1)
                else:
                    pdf.cell(len(testo) + 6, 6, testo, border=0, ln=0)
        else:
            testo = x.replace("\n", "")
            pdf.cell(0, 6, testo, border=0, ln=1)
        i += 1

    #Salva il pdf temporaneo
    pdf.output(config["cartella_lavoro"] + "\\" + "systemgen.pdf")
    pdf.close()

    #Prepara i pdf overlay template e result
    nome_templ = guida[3].replace("\n", "")
    overlay_pdf_file_name = config["cartella_lavoro"] + "\\" + "systemgen.pdf"

    with open(overlay_pdf_file_name, 'rb') as Pdf_overlay:
        if nome_templ != "":
            pdf_template_file_name = config["cartella_template"] + "\\" + nome_templ
            result_pdf_file_name = config["cartella_output"] + "\\" + guida[2].replace("\n", "")

            # Open your overlay PDF that was created earlier

            overlay_pdf = PdfFileReader(Pdf_overlay)
            numpag = overlay_pdf.getNumPages()

            # Merge the overlay page onto the template page
            output_pdf = PdfFileWriter()
            for intero in range(0, numpag):
                pdf_template = PdfFileReader(open(pdf_template_file_name, 'rb'))
                template_page = pdf_template.getPage(0)
                template_page.rotateClockwise(rotazione)
                overlay_page = overlay_pdf.getPage(intero)
                template_page.mergePage(overlay_page)

                # Write the result to a new PDF file
                output_pdf.addPage(template_page)

            # Salva tutto nel result
            resultPdf = open(result_pdf_file_name, 'wb')
            output_pdf.write(resultPdf)

        else:
            result_pdf_file_name = config["cartella_output"] + "\\" + guida[2].replace("\n", "")
            output_pdf = PdfFileWriter()
            # Open your overlay PDF that was created earlier

            overlay_pdf = PdfFileReader(Pdf_overlay)
            numpag = overlay_pdf.getNumPages()

            for intero in range(0, numpag):
                over_page = overlay_pdf.getPage(0)
                over_page.rotateClockwise(rotazione)

                # Write the result to a new PDF file
                output_pdf.addPage(over_page)

            # Salva tutto nel result
            resultPdf = open(result_pdf_file_name, 'wb')
            output_pdf.write(resultPdf)

    if guida[1].replace("\n", "") == "pdf":
        os.startfile(result_pdf_file_name)
        resultPdf.close()
    if guida[1].replace("\n", "") == "stampa":
        #Chiude il file prima di stamparlo, così è scritto del tutto
        resultPdf.close()
        stampa_pdf(result_pdf_file_name)
    #Prova a eliminare il file temporaneo
    try:
        os.remove(config["cartella_lavoro"] + "\\" + "systemgen.pdf")
    except:
        print("Can't remove systemgen.pdf")

def docx(guida, config):
    #Apre il file di testo
    f = open(config["cartella_lavoro"] + "\\" + guida[0].replace("\n", ""), 'r')
    content = f.readlines()
    f.close()

    #dichiara il documento
    document = Document()
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(1.4)
        section.bottom_margin = Cm(1.4)
        section.left_margin = Cm(1.6)
        section.right_margin = Cm(1.6)
    par = document.add_paragraph()
    #dichiara lo stile
    style = document.styles.add_style("Custom", WD_STYLE_TYPE.CHARACTER)
    i = linea_inizio_format_pdf
    index = 0
    testo_tot = ""
    testopasprec = ""
    #for presente anche nella funzione pdf
    #per ogni linea del file di testo va a leggere la corrispondente linea del file guida
    #se è presente applica una formattazione
    #aggiunge linea per linea il testo al documento Word
    for x in content:
        riga = x
        car1prec = 0
        arrcodici = []
        if guida[i][0:4] == "fine":
            i = linea_inizio_format_pdf - 1
            index = 0
        else:
            codicecompleto = guida[i].replace("\n", "")[4:]
            if guida[i][11:13] != "":
                lung = len(guida[i].replace("\n", "")) - 3
                times = int(lung / 13)
                for h in range(0, times):
                    primochar = 13 * h
                    ultimochar = 13 * h + 12
                    arrcodici.append(codicecompleto[primochar: ultimochar])
            else:
                arrcodici.append(codicecompleto)
            linea = guida[i][0:3]
            codice = guida[i][4:10]
            index += 1

            def controllacodice(codice):
                font = ""
                bold = False
                italic = False
                if codice[0] == "t":
                    if codice[1] == "s":
                        if codice[2] == "s":
                            font = "TimesBoldIt"
                            bold = True
                            italic = True
                        else:
                            font = "TimesBold"
                            bold = True
                            italic = False
                    else:
                        if codice[2] == "s":
                            font = "TimesIt"
                            bold = False
                            italic = True
                        else:
                            font = "Times"
                            bold = False
                            italic = False

                if codice[0] == "h":
                    if codice[1] == "s":
                        if codice[2] == "s":
                            font = "HelveticaBoldIt"
                            bold = True
                            italic = True
                        else:
                            font = "HelveticaBold"
                            bold = True
                            italic = False
                    else:
                        if codice[2] == "s":
                            font = "HelveticaIt"
                            bold = False
                            italic = True
                        else:
                            font = "Helvetica"
                            bold = False
                            italic = False
                if codice[0].lower() == "c":
                    if codice[1] == "s":
                        if codice[2] == "s":
                            font = "CourierBoldIt"
                            bold = True
                            italic = True
                        else:
                            font = "CourierBold"
                            bold = True
                            italic = False
                    else:
                        if codice[2] == "s":
                            font = "CourierIt"
                            bold = False
                            italic = True
                        else:
                            font = "Courier"
                            bold = False
                            italic = False
                size = codice[3:5]
                colore = codice[5:6]
                array = [font, size, bold, italic, colore]
                return array

            if codice != "":
                    for codici in arrcodici:
                        codice2 = codici[0:6]
                        if codici[7:] != "":
                            car1 = int(codici[7:10])
                            car2 = int(codici[10:])
                            testopass = riga[car1: car2]
                            testo_tot += testopass
                            if codici == arrcodici[0]:
                                testopasprec = riga[0: car1]
                                testo_tot += testopasprec
                            else:
                                testopasprec = riga[car1prec: car1]
                                testo_tot += testopasprec
                            car1prec = car2
                        else:
                            testopass = riga
                            testo_tot += testopass
                        if testopasprec != "":
                            testoarr = controllacodice(codice)
                            codice = guida[i][5:10]
                            testo = testopasprec
                            run = par.add_run(testo)
                            font = run.font
                            font.name = testoarr[0]
                            font.size = Pt(int(testoarr[1]))
                            font.bold = testoarr[2]
                            font.italic = testoarr[3]
                            color = testoarr[4]
                            if color == "n":
                                font.color.rgb = RGBColor(0,0,0)
                            elif color == "r":
                                font.color.rgb = RGBColor(255,0,0)
                            elif color == "b":
                                font.color.rgb = RGBColor(0,0,255)
                            elif color == "b":
                                font.color.rgb = RGBColor(0,255,0)
                            else:
                                font.color.rgb = RGBColor(0, 0, 0)

                        testoarr2 = controllacodice(codice2)
                        testo = testopass
                        run = par.add_run(testo)
                        font = run.font
                        font.name = testoarr2[0]
                        font.size = Pt(int(testoarr2[1]))
                        font.bold = testoarr2[2]
                        font.italic = testoarr2[3]
                        lastfont = testoarr2[0]
                        lastsize = Pt(int(testoarr2[1]))
                        lastbold = testoarr2[2]
                        lastit = testoarr2[3]
                        color = testoarr2[4]
                        if color == "n":
                            font.color.rgb = RGBColor(0, 0, 0)
                        elif color == "r":
                            font.color.rgb = RGBColor(255, 0, 0)
                        elif color == "b":
                            font.color.rgb = RGBColor(0, 0, 255)
                        elif color == "b":
                            font.color.rgb = RGBColor(0, 255, 0)
                        else:
                            font.color.rgb = RGBColor(0, 0, 0)
            else:
                testo = riga
                run = par.add_run(testo)
                font = run.font
                font.name = lastfont
                font.size = lastsize
                font.bold = lastbold
                font.italic = lastit
            i += 1


    #Salva il documento
    document.save(config["cartella_output"] + "\\" + guida[2].replace("\n", ""))

def excel(guida, config):
    #Salva in content il contenuto del file di testo
    f = open(config["cartella_lavoro"] + "\\" + guida[0].replace("\n", ""), 'r')
    content = f.readlines()
    f.close()

    #Dichiara l'array delle colonne e il dizionario che le conterrà
    colonne = []
    d = {}

    primo_char = 0
    first_chk = 1
    second_chk = 1

    # Create new workbook
    wb = Workbook()
    sheet1 = wb.create_sheet('sheet1', 0)
    active = wb['sheet1']
    ws = wb.active
    foglio = []


    #legge linea per il linea il file e mette le righe in un array
    for line in content:
        riga = []
        num_col = 1
        i = 0
        primo_char = 0
        while i <= len(line):
            char = line[i: i + 1]
            if char == ";":
                res = isinstance(line[primo_char: i], str)
                if res == True:
                    riga.append(line[primo_char: i])
                else:
                    riga.append(float(line[primo_char: i]))
                primo_char = i + 1
                num_col += 1
            i += 1
        colonne.append(riga)

    #Trasforma le righe in colonne e le mette nel dizionario
    for x in range(len(colonne)):
        for y in range(len(colonne[x])):
            d["col" + str(y + 1)] = colonne[x][y]

        #Mette il dizionario in un data frame
        df = pd.DataFrame(data=d, index=[0])

        #trasforma il data frame in un file excel
        #Non funziona
        #df.to_excel('file.xlsx', header=False, index=False)

        #Aggiunge una line alla volta al file excel
        #Ne mette una ogni 3 perchè il dataframe ha un titolo poi una riga vuota poi la riga da passare

        for x in dataframe_to_rows(df):
            if first_chk == 1:
                first_chk = 0
            else:
                if second_chk == 1:
                    second_chk = 0
                else:
                    first_chk = 1
                    second_chk = 1
                    #active.append(x)
                    foglio.append(x)


    #Cancella la prima colonna che contiene l'indice
    sheet1.delete_cols(1, 1)

    #Prepara la variabile dimensioni
    dims = {}

    codici_num= [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 21]
    #Aggiusta la dimensione delle colonne
    i = linea_inizio_format_excel
    for col in range(0, len(foglio[0])):
            if guida[i] != "fine":
                if guida[i] != "":
                    try:
                        codice = int(guida[i].replace("\n", ""))
                    except:
                        codice = 0

                    def controllacodiceexcel(codice):
                        format = ""
                        if codice == 0:
                            format = u''
                        if codice == 1:
                            format = u'0'
                        if codice == 2:
                            format = '0, 00'
                        if codice == 3:
                            format = u'#.##0'
                        if codice == 4:
                            format = u'#.##0,00'
                        if codice == 5:
                            format = "#.##0 \"€\";-#.##0 \"€\""
                        if codice == 6:
                            format = "#.##0 \"€\";[Red]-#.##0 \"€\""
                        if codice == 7:
                            format = "#.##0 \"€\";[Red]-#.##0 \"€\""
                        if codice == 8:
                            format = u'#.##0,00 \"€\";[Red]-#.##0,00 \"€\"'
                        if codice == 9:
                            format = "0%"
                        if codice == 10:
                            format = "0,00%"
                        if codice == 11:
                            format = "0,00E+00"
                        if codice == 12:
                            format = u'dd/mm/yyyy'
                        if codice == 13:
                            format = "gg-mmm-yyyy"
                        if codice == 14:
                            format = "gg-mmm"
                        if codice == 15:
                            format = "mmm yyyy"
                        if codice == 16:
                            format = "h:mm AM/PM"
                        if codice == 17:
                            format = "h:mm:ss AM/PM"
                        if codice == 18:
                            format = "hh:mm"
                        if codice == 19:
                            format = "hh:mm:ss"
                        if codice == 20:
                            format = "gg/mm/yyyy hh:mm"
                        if codice == 21:
                            format = "##0,0E+0"
                        if codice == 22:
                            format = "mm:ss"
                        if codice == 23:
                            format = "@"
                        return format

                    formato = controllacodiceexcel(codice)

                    for f in range(0, len(foglio)):
                        res = isinstance(foglio[f][col], str)
                        if codice in codici_num and f != 0 and foglio[f][col].isnumeric():
                            if foglio[f][col].isspace() :
                                foglio[f][col] = 0
                            ws.cell(row=f + 1, column=col + 1).value = float(foglio[f][col])
                            ws.cell(row=f + 1, column=col + 1).number_format = formato
                        else:
                            ws.cell(row=f + 1, column=col + 1).value = str(foglio[f][col])
                            ws.cell(row=f + 1, column=col + 1).number_format = str(formato)

                        #if formato != "":
                        #    ws.cell(row=f + 1, column=col + 1).number_format = str(formato)

            else:
               sheet1.delete_cols(1, 1)
               column_widths = []
               for row in sheet1.iter_rows():
                   for i, cell in enumerate(row):
                       try:
                           column_widths[i] = max(column_widths[i], len(str(cell.value)))
                       except IndexError:
                           column_widths.append(len(str(cell.value)))
               for i, column_width in enumerate(column_widths):
                   sheet1.column_dimensions[get_column_letter(i + 1)].width = column_width


               wb.save(config["cartella_output"] + "\\" + guida[2].replace("\n", ""))
               return
            i += 1
    sheet1.delete_cols(1, 1)
    column_widths = []
    for row in sheet1.iter_rows():
        for i, cell in enumerate(row):
            try:
                column_widths[i] = max(column_widths[i], len(str(cell.value)))
            except IndexError:
                column_widths.append(len(str(cell.value)))
    for i, column_width in enumerate(column_widths):
        sheet1.column_dimensions[get_column_letter(i + 1)].width = column_width

    # Salva il file Excel

    wb.save(config["cartella_output"] + "\\" + guida[2].replace("\n", ""))

def email_send(guida, config):
    i = 0
    body = ""
    subject = ""
    receiver_email = ""
    filename2 = ""
    #legge le linee dopo fine
    #la prima è il ricevente
    #la seconda è l'oggetto
    #la terza è un eventuale secondo allegato
    #Quarta e seguenti sono per il messaggio
    for x in guida:
        if i == 1 or i == 2:
            if i == 1:
                receiver_email = x.replace("\n", "")
            i += 1
        if i == 3 or i == 4:
            if i == 3:
                subject = x.replace("\n", "")
            i += 1
        if i == 5 or i == 6:
            if i == 5:
                filename2 = x.replace("\n", "")
            i += 1
        if i >= 7:
            body = body + " " + x
            i += 1
        if x[0:4] == "fine":
            i += 1

    port = int(config["email_smtp_porta"])  # For SSL

    #passord della mail
    password = config["email_password"]

    # Create a secure SSL context
    context = ssl.create_default_context()

    #L'email che spedisce
    sender_email = config["email_mittente"]

    #L'allegato generato dalla funzione pdf
    filename = config["cartella_output"] + "\\" + guida[2].replace("\n", "")

    if receiver_email == "":
        return

    #Forma il messaggio
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    message["Bcc"] = receiver_email  # Recommended for mass emails

    # Add body to email
    message.attach(MIMEText(body, "plain"))

    # Open PDF file in binary mode
    with open(filename, "rb") as attachment:
        # Add file as application/octet-stream
        # Email client can usually download this automatically as attachment
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    # Encode file in ASCII characters to send by email
    encoders.encode_base64(part)
    # Add header as key/value pair to attachment part
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {filename}",
    )

    # Add attachment to message and convert message to string
    message.attach(part)

    if filename2 != "":
        with open(filename2, "rb") as attachment:
            # Add file as application/octet-stream
            # Email client can usually download this automatically as attachment
            part2 = MIMEBase("application", "octet-stream")
            part2.set_payload(attachment.read())
        # Encode file in ASCII characters to send by email
        encoders.encode_base64(part2)
        # Add header as key/value pair to attachment part
        part2.add_header(
            "Content-Disposition",
            f"attachment; filename= {filename2}",
        )

        # Add attachment to message and convert message to string
        message.attach(part2)

    #il messaggio che sarà inviato
    text = message.as_string()

    #invia la mail
    with smtplib.SMTP_SSL(config["email_smtp_server"], port, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, text)

def start():
    # Legge la configurazione cifrata (config.dat, si modifica con ConfigManager.py)
    config = carica_config()

    # Cartella con i file di testo e i .gui
    folder = config["cartella_lavoro"]

    # Legge i file in questa cartella
    onlyfiles = [f for f in listdir(folder) if isfile(join(folder, f))]

    # Per ogni file passa il suo nome alla main
    for name in onlyfiles:
        main(name, config)

# file .gui global
linea_inizio_format_pdf = 7
linea_inizio_format_excel = 7

#Crea un array con i file precedenti nella cartella
filesprec = []

process(filesprec)

#loop che avviene ogni secondo
while True:
    time.sleep(1)
    filesprec = process(filesprec)
