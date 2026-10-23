import datetime
import os
import sys
import tempfile
from os import listdir
from os.path import isfile, join
# SUBPROCESS (EXE)
import subprocess

# CONFIGURAZIONE CIFRATA
from config_sicuro import carica_config

# WINDOWS / LINUX E REGISTRO (LOG)
from piattaforma import WINDOWS, ambiente_esterno, apri_file, cerca_font, leggi_righe
import registro

# EMAIL
import email
import html
import math
import re
import mimetypes
import smtplib
import ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# INTERNET (TELEGRAM)
import requests
import telepot

#import cookie

# WHATSAPP (API Cloud di Meta)
from whatsapp_meta import WhatsApp, WhatsAppError, leggi_destinatario, leggi_riga_template

# PDF
from PyPDF2 import PdfFileWriter, PdfFileReader
from fpdf import FPDF
import fpdf.fpdf
# Stampa senza Ghostscript (stampante di Windows o CUPS su Linux)
from stampa_pdf import stampa_pdf

# DOCX
from docx import Document
from docx.shared import RGBColor
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, Cm

# EXCEL
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, PatternFill
# Righe speciali, colonne e formati Excel del .gui, gli stessi dell'editor dei modelli Excel
from modello_gui import (ALLINEAMENTI, FORMATI_EXCEL, PREDEFINITO_EXCEL, PREFISSI_EMAIL, SFONDI, campi_excel,
                         colonna_excel_vuota, immagine_barcode, leggi_excel, righe_excel, stile_riga_vuoto,
                         valore_excel)

#BARCODE (l'immagine la crea modello_gui.immagine_barcode, la stessa dell'editor)
import pymupdf

def main(name, config):
    # Sostituisce al file .txt la terminazione .gui per vedere se esiste
    name_gui = name.replace(".txt", ".gui")
    path_gui = config["cartella_lavoro"] + os.sep + name_gui
    if name_gui[len(name_gui) - 3:len(name_gui)] != "gui":
        return
    if not os.path.isfile(path_gui):
        return
    # Apre il file .gui (creato su Windows o su Linux)
    guida = leggi_righe(path_gui)
    if len(guida) < 2:
        raise ValueError("il .gui ha %d righe: mancano almeno il file di testo e l'azione" % len(guida))
    azione = guida[1].replace("\n", "")
    if azione in ("exe", "barcode", "immagini"):
        funzione = {"exe": exe, "barcode": barcode_add, "immagini": image_add}[azione]
        descrizione = "exe " + guida[0].strip() if azione == "exe" else azione
        passo(name_gui, descrizione, funzione, guida, config)
        # Rimuove i file dopo l'uso
        os.remove(path_gui)
        return

    if name[-3:] == "gui":
        # Un .gui da solo si elabora solo se non crea documenti (linea 3 vuota) e il suo testo non c'è:
        # è un invio senza documento (email, Telegram, WhatsApp). Altrimenti aspetta il .txt
        testo = guida[0].strip()
        if not (len(guida) >= 7 and guida[2].strip() == "" and
                not os.path.isfile(config["cartella_lavoro"] + os.sep + name[:-4] + ".txt") and
                (testo == "" or not os.path.isfile(config["cartella_lavoro"] + os.sep + testo))):
            return
    if len(guida) < 7:
        raise ValueError("il .gui ha %d righe: ne servono almeno 7 (testo, azione, documento, template, "
                         "rotazione, invio, destinatario) e poi i formati fino a \"fine\"" % len(guida))

    # Legge la seconda linea del .gui per decidere cosa fare
    uscita = guida[2].replace("\n", "")
    invio = guida[5].replace("\n", "")
    if uscita.strip() == "":
        # Nessun documento da creare: solo l'invio (email, Telegram, WhatsApp)
        # Senza invio il .gui è sbagliato: va nel log (start()) e i file restano nella cartella di lavoro
        if invio not in ("email", "telegram", "whatsapp"):
            raise ValueError("il .gui non crea nessun documento (linea 3 vuota) e non ha nessun invio (linea 6)")
    elif azione == "pdf" or azione == "file" or azione == "stampa":
        template = guida[3].replace("\n", "")
        risultato = passo(name, "pdf " + uscita + (" su template " + template if template != "" else ""),
                          pdf, guida, config)
        # Il documento è già creato: se non si apre o non si stampa l'errore va nel log e si prosegue
        if azione == "pdf":
            passo_secondario(name, "visualizzazione a schermo di " + uscita, apri_file, risultato)
        if azione == "stampa":
            stampante = config.get("stampante", "")
            passo_secondario(name, "stampa di " + uscita + " su " + (stampante or "stampante predefinita"),
                             stampa_pdf, risultato, stampante)
    elif azione == "excel":
        passo(name, "excel " + uscita, excel, guida, config)
    elif azione == "docx":
        passo(name, "word " + uscita, docx, guida, config)
    else:
        registro.registra(name, "azione sconosciuta nel .gui: " + azione, "ERRORE")
        azione = ""

    # Invio del documento creato: un errore va nel log ma non blocca
    if azione != "" and invio in ("email", "telegram", "whatsapp"):
        if invio == "email":
            descrizione = "email a " + (dati_email(guida)[0] or "?")
        else:
            # "whatsapp <numero>" anche se nel .gui c'è solo il numero
            destinatario = guida[6].strip()
            if destinatario.split()[0:1] != [invio]:
                destinatario = invio + " " + destinatario
            descrizione = destinatario
        funzione = {"email": email_send, "telegram": telegram, "whatsapp": whatsapp}[invio]
        passo_secondario(name, "invio " + descrizione, funzione, guida, config)

    # Rimuove i file dopo l'uso (un .gui di solo invio può non avere il testo)
    os.remove(path_gui)
    testo = guida[0].strip()
    if testo != "" and os.path.isfile(config["cartella_lavoro"] + os.sep + testo):
        os.remove(config["cartella_lavoro"] + os.sep + testo)


def passo(nome_file, descrizione, funzione, *argomenti):
    """Esegue un'operazione e la scrive nel log con il suo esito.

    Se non riesce l'errore va nel log e viene rilanciato: il .txt e il .gui restano nella
    cartella di lavoro e start() passa al file successivo.
    """
    try:
        risultato = funzione(*argomenti)
    except Exception as e:
        registro.registra_errore(nome_file, descrizione, e)
        try:
            e.registrato = True
        except AttributeError:
            pass
        raise
    registro.registra(nome_file, descrizione)
    return risultato


def passo_secondario(nome_file, descrizione, funzione, *argomenti):
    """Come passo(), ma un errore non ferma l'elaborazione: resta nel log e sulla console."""
    try:
        passo(nome_file, descrizione, funzione, *argomenti)
    except Exception as e:
        segnala_errore(nome_file, descrizione, e)


def segnala_errore(nome_file, descrizione, errore):
    """Una riga sulla console per un errore già scritto nel log (lì c'è anche il traceback)."""
    print("Errore con il file %s: %s: %s: %s (dettagli nel log)" % (nome_file, descrizione,
                                                                   type(errore).__name__, errore), file=sys.stderr)


def coda_gui(guida):
    """Le linee del .gui dopo "fine" (dati dell'invio), senza a capo."""
    for n in range(linea_inizio_format_pdf, len(guida)):
        if guida[n][0:4] == "fine":
            return [x.replace("\n", "") for x in guida[n + 1:]]
    return []


def nome_file(percorso):
    """Solo il nome del file, sia da un percorso di Windows (\\) sia di Linux (/)."""
    return percorso.replace("\\", "/").split("/")[-1]


def barcode_add(guida, config):
    numero_barcodes = int(guida[2].replace("\n", ""))
    for i in range(0, numero_barcodes):
        i *= 6
        codice = guida[3 + i].replace("\n", "")
        immagine = immagine_barcode(codice, guida[4 + i].replace("\n", ""))

        # la posizione è data da due vertici opposti partendo da quello alto a sx
        x0, y0 = guida[5 + i].replace("\n", "").split(",")
        x1, y1 = guida[6 + i].replace("\n", "").split(",")
        image_rectangle = pymupdf.Rect(int(x0), int(y0), int(x1), int(y1))

        # retrieve the page of the PDF
        doc = pymupdf.open(guida[7 + i].replace("\n", ""))
        page = doc[int(guida[8 + i].replace("\n", ""))]

        # add the image
        page.insert_image(image_rectangle, stream=immagine)
        doc.save(config["cartella_lavoro"] + os.sep + nome_file(guida[7 + i].replace("\n", "")), incremental=True, encryption=0)
        doc.close()

def image_add(guida,config):
    numero_image = int(guida[2].replace("\n", ""))
    for i in range(0, numero_image):
        i *= 5
        # la posizione è data da due vertici opposti partendo da quello alto a sx
        percorso = guida[3 + i].replace("\n", "")
        x0, y0 = guida[4 + i].replace("\n", "").split(",")
        x1, y1 = guida[5 + i].replace("\n", "").split(",")
        image_rectangle = pymupdf.Rect(int(x0), int(y0), int(x1), int(y1))

        # retrieve the page of the PDF
        doc = pymupdf.open(guida[6 + i].replace("\n", ""))
        page = doc[int(guida[7 + i].replace("\n", ""))]

        # add the image
        page.insert_image(image_rectangle, filename= percorso)
        doc.save(config["cartella_lavoro"] + os.sep + nome_file(guida[6 + i].replace("\n", "")), incremental=True, encryption=0)
        doc.close()


def exe(guida, config) :

    path = config["cartella_lavoro"] + os.sep + guida[0].replace("\n", "")
    if not os.path.isfile(path):
        path = guida[0].replace("\n", "")
    if not os.path.isfile(path):
        raise FileNotFoundError("Programma non trovato: " + guida[0].replace("\n", ""))
    if WINDOWS:
        #Tra virgolette, così funzionano anche i percorsi con spazi
        os.system('"' + path + '"')
    else:
        #Su Linux il file deve essere eseguibile (chmod +x)
        subprocess.call([os.path.abspath(path)], env=ambiente_esterno())


def percorso_documento(documento, config):
    """Il file da mandare con Telegram o WhatsApp: il nome di un documento creato (nella cartella di
    output) oppure il percorso completo di un file che c'è già sul computer."""
    if os.path.isabs(documento):
        return documento
    return config["cartella_output"] + os.sep + documento


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
    #la prima linea è il documento: il nome nella cartella di output o il percorso completo di un
    #file già sul computer (vuota = nessun documento)
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
    doc_path = percorso_documento(document, config)

    #manda il messaggio
    bot.sendMessage(bot_chatID, bot_message)
    #se esiste manda il documento
    if document != "":
        bot.sendDocument(chat_id=bot_chatID, document=open(doc_path, 'rb'))

def whatsapp(guida, config):
    #Usa le API Cloud di Meta: token e Phone number ID si impostano con ConfigManager.py
    #La settima linea del .gui è "whatsapp <numero>" (la parola whatsapp si può omettere)
    #il numero va con il prefisso internazionale (es. 393331234567 o +39 333 1234567)
    #legge le ultime linee del file dopo "fine"
    #la prima linea è il nome del documento nella cartella di output o il percorso completo di un
    #file già sul computer (vuota = nessun documento)
    #Messaggio: la seconda e seguenti sono il testo del messaggio.
    #Senza template il messaggio arriva solo se il destinatario ha scritto al numero
    #aziendale nelle ultime 24 ore (regola di Meta).
    #Template: la seconda è "template: <nome> [<lingua>]", le seguenti (non vuote) sono i
    #parametri {{1}}, {{2}}, ... e il documento diventa l'intestazione (il template deve avere
    #un'intestazione Documento). Nei .gui di prima nome e lingua erano nella linea 7, dopo il numero.
    numero, template, lingua = leggi_destinatario(guida[6].replace("\n", ""))
    if numero == "":
        raise WhatsAppError("numero del destinatario mancante nella linea 7 del .gui")

    #prende le linee dopo "fine"
    coda = coda_gui(guida)
    document = coda[0].strip() if len(coda) > 0 else ""
    righe = coda[1:]
    riga_template = leggi_riga_template(righe[0]) if righe else None
    if riga_template is not None:
        template, lingua = riga_template
        righe = righe[1:]
        if template == "":
            raise WhatsAppError("nome del template mancante nella riga \"template:\"")
    lingua = lingua or config["whatsapp_lingua"]
    doc_path = percorso_documento(document, config) if document != "" else None

    #Gli errori (WhatsAppError, rete, file mancante) vanno nel log tramite passo_secondario()
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


def con_fine(guida):
    """Le linee del .gui con "fine" dopo i formati anche se manca (altrimenti si leggerebbe oltre)."""
    if any(linea[0:4] == "fine" for linea in guida[linea_inizio_format_pdf:]):
        return guida
    return list(guida) + ["fine\n"]


def pdf(guida, config):
    guida = con_fine(guida)

    #Apre il documento di testo
    content = leggi_righe(config["cartella_lavoro"] + os.sep + guida[0].replace("\n", ""))

    #Immagazzina la rotazione(eventuale) da applicare alla pagina
    rotazione = int(guida[4])

    #Dichiara il documento
    #Se il template è orizzontale anche il testo va su una pagina orizzontale,
    #altrimenti le prime righe finiscono fuori dal foglio
    orientamento = "P"
    nome_templ = guida[3].replace("\n", "")
    if nome_templ != "":
        with pymupdf.open(config["cartella_template"] + os.sep + nome_templ) as doc_templ:
            if doc_templ[0].mediabox.width > doc_templ[0].mediabox.height:
                orientamento = "L"
    pdf = FPDF(orientation=orientamento)

    #Imposta font, dimensione e colore. Il font viene aggiunto al pdf solo la prima volta che serve
    #(ogni font aggiunto finisce dentro il file). Il file del font lo cerca trova_font()
    def usa_font(nome_font, dimensione, colore="n"):
        if nome_font != "" and nome_font.lower() not in pdf.fonts:
            pdf.add_font(nome_font, '', trova_font(nome_font), uni=True)
        pdf.set_font(family=nome_font, size=dimensione)
        pdf.set_text_color(*COLORI_TESTO.get(colore, (0, 0, 0)))

    #Aggiunge la prima pagina al pdf
    pdf.alias_nb_pages()
    pdf.add_page()
    #Font predefinito per le righe senza formato (stesso del codice iniziale "tnn10")
    usa_font("Times", 10)

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
        #Il testo prima del formato va azzerato a ogni riga, altrimenti si ripete nelle righe dopo
        testopasprec = ""
        if guida[i][0:4] == "fine":
            i = linea_inizio_format_pdf
            index = 0

        codicecompleto = guida[i].replace("\n", "")[4:]
        #Il codice ha 5 caratteri (es. tsn12) o 6 con il colore (es. tsn12r)
        lunghezza_codice = 6 if codicecompleto[5:6] in COLORI_TESTO else 5
        if len(codicecompleto.rstrip()) > lunghezza_codice:
            #Più formati sulla riga: "<codice> <da> <a>" separati da un carattere
            passo = lunghezza_codice + 7
            times = int((len(codicecompleto) + 1) / passo)
            for h in range(0, times):
                primochar = passo*h
                ultimochar = passo*h + lunghezza_codice + 6
                arrcodici.append(codicecompleto[primochar: ultimochar])
        else:
            arrcodici.append(codicecompleto)

        index += 1
        testo_tot = ""
        def controllacodice(codice):
                #Font (es. Times, TimesBold, TimesBoldIt, TimesIt), dimensione e colore
                font = ""
                if codice[0:1].lower() in FONT_CODICI:
                    font = FONT_CODICI[codice[0:1].lower()][0]
                    if codice[1:2].lower() == "s":
                        font += "Bold"
                    if codice[2:3].lower() == "s":
                        font += "It"
                size = codice[3:5]
                colore = codice[5:6] if codice[5:6] in COLORI_TESTO else "n"
                array = [font, size, colore]
                return array

        if arrcodici[0] != "" and arrcodici[0] != " ":
            for codici in arrcodici:
                codice2 = codici[0:lunghezza_codice]
                if codici[lunghezza_codice + 1:] != "":
                    car1 = int(codici[lunghezza_codice + 1:lunghezza_codice + 4])
                    car2 = int(codici[lunghezza_codice + 4:])
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
                    codice = guida[i][4:4 + lunghezza_codice]
                    testo = testopasprec.replace("\n", "")
                    usa_font(testoarr[0], int(testoarr[1]), testoarr[2])
                    #Larghezza della cella = larghezza reale del testo, così i pezzi non si sovrappongono
                    pdf.cell(pdf.get_string_width(testo), 6, testo, border=0, ln=0)
                testoarr2 = controllacodice(codice2)
                testo = testopass.replace("\n", "")
                usa_font(testoarr2[0], int(testoarr2[1]), testoarr2[2])
                if codici == arrcodici[len(arrcodici) - 1]:
                    #Il testo dopo l'ultimo pezzo formattato prende lo stesso formato
                    if codici[lunghezza_codice + 1:] != "":
                        testo += riga[car1prec:].replace("\n", "")
                    pdf.cell(0, 6, testo, border=0, ln=1)
                else:
                    pdf.cell(pdf.get_string_width(testo), 6, testo, border=0, ln=0)
        else:
            testo = x.replace("\n", "")
            pdf.cell(0, 6, testo, border=0, ln=1)
        #Senza linee di formato (subito "fine") si resta lì: le righe dopo "fine" sono i dati dell'invio
        if guida[i][0:4] != "fine":
            i += 1

    #Salva il pdf temporaneo
    pdf.output(config["cartella_lavoro"] + os.sep + "systemgen.pdf")
    pdf.close()

    #Prepara i pdf overlay template e result
    nome_templ = guida[3].replace("\n", "")
    overlay_pdf_file_name = config["cartella_lavoro"] + os.sep + "systemgen.pdf"
    with open(overlay_pdf_file_name, 'rb') as Pdf_overlay:
        if nome_templ != "":
            pdf_template_file_name = config["cartella_template"] + os.sep + nome_templ
            result_pdf_file_name = config["cartella_output"] + os.sep + guida[2].replace("\n", "")

            # Open your overlay PDF that was created earlier


            overlay_pdf = PdfFileReader(Pdf_overlay)
            numpag = overlay_pdf.getNumPages()

        # Merge the overlay page onto the template page
            output_pdf = PdfFileWriter()
            #Il template resta aperto finché il risultato non è scritto (PyPDF2 lo legge in quel momento)
            file_template = []
            for intero in range(0, numpag):
                file_template.append(open(pdf_template_file_name, 'rb'))
                pdf_template = PdfFileReader(file_template[-1])
                template_page = pdf_template.getPage(0)
                template_page.rotateClockwise(rotazione)
                overlay_page = overlay_pdf.getPage(intero)
                template_page.mergePage(overlay_page)

            # Write the result to a new PDF file
                output_pdf.addPage(template_page)

        #Salva tutto nel result
            with open(result_pdf_file_name, 'wb') as resultPdf:
                output_pdf.write(resultPdf)
            for f in file_template:
                f.close()

        else:
            result_pdf_file_name = config["cartella_output"] + os.sep + guida[2].replace("\n", "")
            output_pdf = PdfFileWriter()
            # Open your overlay PDF that was created earlier

            overlay_pdf = PdfFileReader(Pdf_overlay)
            numpag = overlay_pdf.getNumPages()

            for intero in range(0, numpag):
                over_page = overlay_pdf.getPage(intero)
                over_page.rotateClockwise(rotazione)

                # Write the result to a new PDF file
                output_pdf.addPage(over_page)

            # Salva tutto nel result
            with open(result_pdf_file_name, 'wb') as resultPdf:
                output_pdf.write(resultPdf)
    #Elimina il file temporaneo
    os.remove(config["cartella_lavoro"] + os.sep + "systemgen.pdf")
    #Il risultato è chiuso (scritto del tutto): main() lo apre a schermo o lo stampa
    return result_pdf_file_name


def docx(guida, config):
    guida = con_fine(guida)
    #Apre il file di testo
    content = leggi_righe(config["cartella_lavoro"] + os.sep + guida[0].replace("\n", ""))

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
    #Formato delle righe senza codice prima del primo formato
    lastfont = "Times New Roman"
    lastsize = Pt(10)
    lastbold = False
    lastit = False
    lastcolor = "n"
    #for presente anche nella funzione pdf
    #per ogni linea del file di testo va a leggere la corrispondente linea del file guida
    #se è presente applica una formattazione
    #aggiunge linea per linea il testo al documento Word
    for x in content:
        riga = x
        car1prec = 0
        arrcodici = []
        #Il testo prima del formato va azzerato a ogni riga, altrimenti si ripete nelle righe dopo
        testopasprec = ""
        #Finiti i formati riparte dal primo, come nel pdf
        if guida[i][0:4] == "fine":
            i = linea_inizio_format_pdf
            index = 0
        if guida[i][0:4] == "fine":
            #Il .gui non ha formati: la riga mantiene l'ultimo formato
            run = par.add_run(riga)
            run.font.name = lastfont
            run.font.size = lastsize
            run.font.bold = lastbold
            run.font.italic = lastit
            run.font.color.rgb = RGBColor(*COLORI_TESTO.get(lastcolor, (0, 0, 0)))
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
                #Nome del font come lo conosce Word (es. Times New Roman), dimensione, grassetto,
                #corsivo e colore
                font = None
                if codice[0:1].lower() in FONT_CODICI:
                    font = FONT_CODICI[codice[0:1].lower()][1]
                bold = codice[1:2].lower() == "s"
                italic = codice[2:3].lower() == "s"
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
                            codice = guida[i][4:10]
                            testo = testopasprec
                            run = par.add_run(testo)
                            font = run.font
                            font.name = testoarr[0]
                            font.size = Pt(int(testoarr[1]))
                            font.bold = testoarr[2]
                            font.italic = testoarr[3]
                            font.color.rgb = RGBColor(*COLORI_TESTO.get(testoarr[4], (0, 0, 0)))

                        testoarr2 = controllacodice(codice2)
                        testo = testopass
                        #Il testo dopo l'ultimo pezzo formattato prende lo stesso formato
                        if codici == arrcodici[len(arrcodici) - 1] and codici[7:] != "":
                            testo += riga[car1prec:]
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
                        lastcolor = testoarr2[4]
                        font.color.rgb = RGBColor(*COLORI_TESTO.get(lastcolor, (0, 0, 0)))
            else:
                testo = riga
                run = par.add_run(testo)
                font = run.font
                font.name = lastfont
                font.size = lastsize
                font.bold = lastbold
                font.italic = lastit
                font.color.rgb = RGBColor(*COLORI_TESTO.get(lastcolor, (0, 0, 0)))
            i += 1

    #La formattazione si prepara una volta per tutte nel modello .gui con EditorModelli.py
    #(prima qui si apriva una finestra per ogni documento, che bloccava il servizio)

    #Salva il documento
    document.save(config["cartella_output"] + os.sep + guida[2].replace("\n", ""))


def font_excel(formato):
    #Font di openpyxl da un formato (font, grassetto, corsivo, dimensione, colore) del .gui
    lettera, grassetto, corsivo, dimensione, colore = formato
    #Colore ARGB con FF davanti (opaco): con 6 cifre openpyxl scrive 00, che alcuni programmi
    #considerano trasparente
    return Font(name=FONT_CODICI[lettera][1], size=dimensione, bold=grassetto, italic=corsivo,
                color="FF%02X%02X%02X" % COLORI_TESTO.get(colore, (0, 0, 0)))


def larghezza_mostrata(testo, formato, carattere):
    """Larghezza di colonna (in caratteri di Excel) che serve a mostrare il testo con il suo formato
    e carattere. Si parte dal testo del file (es. "1.234,56", come lo mostra il formato), più i
    simboli che il formato aggiunge; i caratteri grandi o in grassetto occupano di più. Se la colonna
    è stretta Excel e LibreOffice mostrano ### al posto dei numeri."""
    lunghezza = len(testo.strip())
    if "#,##" in formato and "." not in testo:
        #I punti delle migliaia che il formato aggiunge (1234 -> 1.234)
        cifre = len(testo.strip().split(",")[0].lstrip("-"))
        lunghezza += max(0, (cifre - 1) // 3)
    if '"€"' in formato:
        lunghezza += 2
    if "AM/PM" in formato:
        lunghezza += 3
    if "E+" in formato:
        lunghezza = max(lunghezza, 9)
    if "mmm" in formato:
        lunghezza += 1
    _, grassetto, corsivo, dimensione, _ = carattere
    fattore = dimensione / 11 * (1.1 if grassetto else 1) * (1.05 if corsivo else 1)
    return math.ceil(lunghezza * max(1, fattore)) + 2


def excel(guida, config):
    #Salva in content il contenuto del file di testo
    content = leggi_righe(config["cartella_lavoro"] + os.sep + guida[0].replace("\n", ""))

    # Create new workbook
    wb = Workbook()
    sheet1 = wb.create_sheet('sheet1', 0)
    #Toglie il foglio vuoto "Sheet" che Workbook crea da solo
    wb.remove(wb['Sheet'])
    ws = wb.active

    #Ogni riga del testo è una riga del foglio, i campi sono separati da ";"
    #(Prima ogni riga passava da un data frame di pandas e se ne prendeva una ogni 3 righe prodotte:
    #una riga vuota ne produceva 2 e faceva slittare tutte le successive)
    foglio = [campi_excel(line) for line in content]

    #Dalla linea 7 del .gui: le righe speciali (titoli, totali...), la linea "colonna" e una linea
    #per ogni colonna fino a "fine" (vedi modello_gui.py). La prima linea delle colonne è per la
    #prima colonna (prima andava a una colonna indice che poi veniva cancellata, e tutti i formati
    #risultavano spostati di una colonna)
    stili, colonne = leggi_excel(guida[linea_inizio_format_excel:])
    #Riga del foglio (da 1) -> stile; se una riga è in più stili vale l'ultimo
    stile_della_riga = {}
    for stile in stili:
        for numero in righe_excel(stile["righe"], len(foglio)):
            stile_della_riga[numero] = stile

    num_colonne = max(len(r) for r in foglio) if len(foglio) > 0 else 0
    larghezze_fisse = {}
    #Larghezza adatta al testo più lungo di ogni colonna (quando non è scritta nel .gui)
    larghezze_testo = [0] * num_colonne
    for col in range(0, num_colonne):
        #Le colonne dopo "fine" hanno il formato generale
        colonna = colonne[col] if col < len(colonne) else colonna_excel_vuota()
        codice = colonna["formato"]
        formato = FORMATI_EXCEL[codice][1]
        if colonna["larghezza"]:
            larghezze_fisse[col] = colonna["larghezza"]

        for f in range(0, len(foglio)):
            valore = foglio[f][col] if col < len(foglio[f]) else ""
            stile = stile_della_riga.get(f + 1, stile_riga_vuoto())
            cella = ws.cell(row=f + 1, column=col + 1)
            #Le righe con "testo=si" (es. l'intestazione) restano testo, senza formato numerico
            if stile["testo"]:
                cella.value = valore
            else:
                cella.value = valore_excel(valore, codice)
                if formato != "":
                    cella.number_format = formato
            #Carattere, sfondo e allineamento della riga speciale valgono più di quelli della colonna
            carattere = stile["carattere"] or colonna["carattere"]
            if carattere is not None:
                cella.font = font_excel(carattere)
            if stile["sfondo"]:
                colore = "FF" + SFONDI[stile["sfondo"]][1]
                cella.fill = PatternFill(fill_type="solid", start_color=colore, end_color=colore)
            allineamento = stile["allineamento"] or colonna["allineamento"]
            if allineamento:
                cella.alignment = Alignment(horizontal=ALLINEAMENTI[allineamento])
            larghezze_testo[col] = max(larghezze_testo[col], larghezza_mostrata(
                valore, "" if stile["testo"] else formato, carattere or PREDEFINITO_EXCEL))

    #Larghezza delle colonne: quella scritta nel .gui, altrimenti adatta al testo più lungo
    for n, larghezza in enumerate(larghezze_testo):
        sheet1.column_dimensions[get_column_letter(n + 1)].width = larghezze_fisse.get(n, larghezza)

    # Salva il file Excel
    wb.save(config["cartella_output"] + os.sep + guida[2].replace("\n", ""))


def dati_email(guida):
    """Destinatario, oggetto, altri allegati (lista), immagini nel testo (lista) e testo di un invio email.

    Con l'indirizzo nella linea 7 del .gui, dopo "fine" ci sono l'oggetto, le eventuali righe
    "allegato: <percorso>" (una per allegato) e "immagine: <percorso>" (una per immagine da
    mostrare dopo il testo), una riga vuota e il testo del messaggio (anche su più righe).
    Con la linea 7 vuota (i .gui di prima) dopo "fine" ci sono l'indirizzo, l'oggetto,
    un eventuale secondo allegato e il testo del messaggio.
    """
    coda = coda_gui(guida)
    destinatario = guida[6].strip() if len(guida) > 6 else ""
    if destinatario != "":
        oggetto = coda[0] if len(coda) > 0 else ""
        testo = coda[1:]
        file = {"allegati": [], "immagini": []}
        while testo:
            prefisso = testo[0].strip().lower().split(":", 1)[0] + ":"
            if prefisso not in PREFISSI_EMAIL:
                break
            file[PREFISSI_EMAIL[prefisso]].append(testo.pop(0).strip()[len(prefisso):].strip())
        # La riga vuota tra oggetto (o allegati) e testo
        if testo and testo[0].strip() == "":
            testo = testo[1:]
        return (destinatario, oggetto, [a for a in file["allegati"] if a], [i for i in file["immagini"] if i],
                "\n".join(testo))
    coda += [""] * (3 - len(coda))
    return coda[0].strip(), coda[1], [coda[2].strip()] if coda[2].strip() else [], [], "\n".join(coda[3:])


def allegato_email(percorso):
    with open(percorso, "rb") as attachment:
        # Add file as application/octet-stream
        # Email client can usually download this automatically as attachment
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    # Encode file in ASCII characters to send by email
    encoders.encode_base64(part)
    # Add header as key/value pair to attachment part
    part.add_header(
        "Content-Disposition",
        #Solo il nome del file, non tutto il percorso
        "attachment", filename=os.path.basename(percorso),
    )
    return part


# Larghezza massima (in pixel) con cui l'email mostra le immagini del testo e quella della firma
LARGHEZZA_IMMAGINE_EMAIL = 600
LARGHEZZA_FIRMA_EMAIL = 300


def immagine_email(percorso, cid, larghezza_massima):
    """La parte MIME di un'immagine da mostrare nel testo dell'email e il tag <img> che la mostra."""
    tipo = mimetypes.guess_type(percorso)[0] or ""
    if not tipo.startswith("image/"):
        raise ValueError("non è un'immagine (png, jpg, gif): " + percorso)
    with open(percorso, "rb") as f:
        dati = f.read()
    parte = MIMEImage(dati, _subtype=tipo.split("/")[1])
    parte.add_header("Content-ID", "<" + cid + ">")
    parte.add_header("Content-Disposition", "inline", filename=os.path.basename(percorso))
    # Le immagini grandi (es. una firma scansionata) si mostrano più piccole
    larghezza = ""
    try:
        from PIL import Image
        with Image.open(percorso) as immagine:
            if immagine.width > larghezza_massima:
                larghezza = ' width="%d"' % larghezza_massima
    except Exception:
        pass
    return parte, '<img src="cid:%s"%s alt="" style="max-width:100%%">' % (cid, larghezza)


def testo_da_html(codice):
    """La versione solo testo di un pezzo di HTML (la firma scritta in HTML): a capo al posto di
    <br>, </p>, </div>, </tr>, senza gli altri tag."""
    codice = re.sub(r"\s*\n\s*", " ", codice)
    codice = re.sub(r"(?i)<br\s*/?>|</?(p|div|tr|li|table|h[1-6])\b[^>]*>", "\n", codice)
    codice = re.sub(r"(?is)<(style|script)\b.*?</\1\s*>|<[^>]*>", "", codice)
    righe = [" ".join(html.unescape(r).split()) for r in codice.split("\n")]
    return re.sub(r"\n{3,}", "\n\n", "\n".join(righe)).strip()


def corpo_email(testo, immagini, firma, immagine_firma, firma_html=False):
    """Il testo dell'email: solo testo se non ci sono immagini né firma, altrimenti HTML con le
    immagini dentro il messaggio (non come allegati) più la versione solo testo per i programmi
    che non mostrano l'HTML. firma_html=True: la firma è già HTML e va così com'è."""
    firma = firma.strip()
    immagine_firma = immagine_firma.strip()
    if not immagini and not firma and not immagine_firma:
        return MIMEText(testo, "plain")

    def paragrafo(t):
        return html.escape(t).replace("\n", "<br>\n")
    parti = []
    pezzi = ['<div style="font-family:Arial,sans-serif;font-size:10pt">', paragrafo(testo)]
    for n, percorso in enumerate(immagini):
        parte, tag = immagine_email(percorso, "immagine%d" % (n + 1), LARGHEZZA_IMMAGINE_EMAIL)
        parti.append(parte)
        pezzi.append("<p>" + tag + "</p>")
    if firma or immagine_firma:
        # In un <div>: la firma in HTML può contenere tabelle e paragrafi, che in un <p> non vanno
        pezzi.append('<div style="margin-top:16px">' + (firma if firma_html else paragrafo(firma)))
        if immagine_firma:
            parte, tag = immagine_email(immagine_firma, "firma", LARGHEZZA_FIRMA_EMAIL)
            parti.append(parte)
            pezzi.append(("<br>" if firma else "") + tag)
        pezzi.append("</div>")
    pezzi.append("</div>")
    firma_testo = testo_da_html(firma) if firma_html else firma
    alternative = MIMEMultipart("alternative")
    alternative.attach(MIMEText(testo + ("\n\n" + firma_testo if firma_testo else ""), "plain"))
    alternative.attach(MIMEText("\n".join(pezzi), "html"))
    if not parti:
        return alternative
    # Le immagini stanno accanto all'HTML che le mostra (multipart/related)
    correlate = MIMEMultipart("related")
    correlate.attach(alternative)
    for parte in parti:
        correlate.attach(parte)
    return correlate


def email_send(guida, config):
    receiver_email, subject, altri_allegati, immagini, body = dati_email(guida)

    port = int(config["email_smtp_porta"])  # For SSL

    #passord della mail
    password = config["email_password"]

    # Create a secure SSL context
    context = ssl.create_default_context()

    #L'email che spedisce
    sender_email = config["email_mittente"]

    #L'allegato generato dalla funzione pdf (nessuno se il .gui non crea documenti)
    documento = guida[2].replace("\n", "").strip()
    allegati = ([config["cartella_output"] + os.sep + documento] if documento else []) + altri_allegati

    if receiver_email == "":
        raise ValueError("indirizzo del destinatario mancante (campo Destinatario, linea 7 del .gui)")

    #Forma il messaggio
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    message["Bcc"] = receiver_email  # Recommended for mass emails

    # Il testo, con le immagini e la firma impostata in ConfigManager. Se l'immagine della firma non
    # c'è più l'email parte lo stesso (senza): altrimenti si fermerebbero tutte le email
    immagine_firma = config.get("email_firma_immagine", "").strip()
    if immagine_firma and not os.path.isfile(immagine_firma):
        registro.registra(guida[0].strip(), "immagine della firma non trovata, email mandata senza: " +
                          immagine_firma, "AVVISO")
        immagine_firma = ""
    message.attach(corpo_email(body, immagini, config.get("email_firma", ""), immagine_firma,
                               config.get("email_firma_html", "") == "1"))

    # Il documento creato e poi gli altri allegati
    for percorso in allegati:
        message.attach(allegato_email(percorso))

    #il messaggio che sarà inviato
    text = message.as_string()

    #invia la mail
    with smtplib.SMTP_SSL(config["email_smtp_server"], port, context=context) as server:
        try:
            server.login(sender_email, password)
        except (smtplib.SMTPAuthenticationError, smtplib.SMTPServerDisconnected) as e:
            raise RuntimeError("il server di posta " + config["email_smtp_server"] + " non accetta mittente e "
                               "password (con Gmail serve una password per le app di 16 caratteri, non quella "
                               "dell'account): correggili in ConfigManager. Risposta: " + str(e)) from e
        server.sendmail(sender_email, receiver_email, text)


def start():
    """Elabora i file della cartella di lavoro. Non solleva eccezioni: ogni errore va nel log (con
    il traceback) e sulla console in una riga. Restituisce 0, oppure 1 se non si è potuto iniziare
    (configurazione o cartella di lavoro mancanti)."""
    # Legge la configurazione cifrata (config.dat, si modifica con ConfigManager.py)
    try:
        config = carica_config()
        registro.configura(config.get("file_log", ""))
        # Cartella con i file di testo e i .gui
        folder = config["cartella_lavoro"]
    except Exception as e:
        # Senza configurazione il log va nel file predefinito
        registro.configura()
        registro.registra_errore("config.dat", "lettura della configurazione", e)
        segnala_errore("config.dat", "lettura della configurazione", e)
        return 1

    # Legge i file in questa cartella
    try:
        onlyfiles = [f for f in listdir(folder) if isfile(join(folder, f))]
    except OSError as e:
        registro.registra_errore(folder, "lettura della cartella di lavoro", e)
        segnala_errore(folder, "lettura della cartella di lavoro", e)
        return 1

    # Per ogni file passa il suo nome alla main
    for name in onlyfiles:
        #Un file sbagliato non deve bloccare gli altri: scrive l'errore nel log e passa al successivo
        try:
            main(name, config)
        except Exception as e:
            if not getattr(e, "registrato", False):
                registro.registra_errore(name, "elaborazione", e)
            segnala_errore(name, "elaborazione", e)
    return 0


def trova_font(nome_font):
    """File del font per il PDF.

    Prima la cartella Fonts del programma, poi i font installati nel sistema (Windows, Linux, macOS).
    Se il font non è installato (es. Calibri o Segoe UI su Linux) usa il più simile della cartella
    Fonts: Helvetica al posto dei caratteri senza grazie, Times al posto di Georgia.
    """
    nomi = [nome_font]
    for famiglia, sostituto in FONT_SOSTITUTI.items():
        if nome_font.startswith(famiglia):
            nomi.append(sostituto + nome_font[len(famiglia):])
    for nome in nomi:
        file_font, file_sistema = FONT_PDF[nome]
        percorso_font = os.path.join(CARTELLA_FONT, file_font)
        if os.path.isfile(percorso_font):
            return percorso_font
        percorso_font = cerca_font(file_sistema)
        if percorso_font is not None:
            return percorso_font
    raise FileNotFoundError("Font non trovato: " + nome_font + " (manca la cartella Fonts accanto al programma?)")

# file .gui global
linea_inizio_format_pdf = 7
linea_inizio_format_excel = 7

# Font del PDF: cartella Fonts accanto al programma (o all'eseguibile creato con PyInstaller)
if getattr(sys, "frozen", False):
    CARTELLA_FONT = os.path.join(os.path.dirname(sys.executable), "Fonts")
else:
    CARTELLA_FONT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Fonts")
# Lettera del font nei codici del .gui -> (nome nel PDF, nome in Word)
FONT_CODICI = {
    "t": ("Times", "Times New Roman"),
    "h": ("Helvetica", "Arial"),
    "c": ("Courier", "Courier New"),
    "k": ("Calibri", "Calibri"),
    "v": ("Verdana", "Verdana"),
    "g": ("Georgia", "Georgia"),
    "u": ("SegoeUI", "Segoe UI"),
    #A larghezza fissa: Consolas è di Windows, gli altri sono nella cartella Fonts (licenze libere)
    "o": ("Consolas", "Consolas"),
    "d": ("DejaVuSansMono", "DejaVu Sans Mono"),
    "p": ("IBMPlexMono", "IBM Plex Mono"),
    "l": ("LiberationMono", "Liberation Mono"),
}
# Font che di solito ci sono solo su Windows -> font della cartella Fonts da usare se mancano
# (al posto di Consolas un altro a larghezza fissa, così le colonne restano allineate)
FONT_SOSTITUTI = {"Calibri": "Helvetica", "Verdana": "Helvetica", "SegoeUI": "Helvetica", "Georgia": "Times",
                  "Consolas": "DejaVuSansMono"}
# Nome usato nel PDF -> (file nella cartella Fonts, file del sistema se manca)
FONT_PDF = {
    "Times": ("times.ttf", "times.ttf"),
    "TimesBold": ("timesbd.ttf", "timesbd.ttf"),
    "TimesBoldIt": ("timesbi.ttf", "timesbi.ttf"),
    "TimesIt": ("timesi.ttf", "timesi.ttf"),
    "Helvetica": ("Helvetica.ttf", "arial.ttf"),
    "HelveticaBold": ("Helvetica-Bold.ttf", "arialbd.ttf"),
    "HelveticaBoldIt": ("Helvetica-BoldOblique.ttf", "arialbi.ttf"),
    "HelveticaIt": ("Helvetica-Oblique.ttf", "ariali.ttf"),
    "Courier": ("cour.ttf", "cour.ttf"),
    "CourierBold": ("courbd.ttf", "courbd.ttf"),
    "CourierBoldIt": ("courbi.ttf", "courbi.ttf"),
    "CourierIt": ("couri.ttf", "couri.ttf"),
    "Calibri": ("calibri.ttf", "calibri.ttf"),
    "CalibriBold": ("calibrib.ttf", "calibrib.ttf"),
    "CalibriBoldIt": ("calibriz.ttf", "calibriz.ttf"),
    "CalibriIt": ("calibrii.ttf", "calibrii.ttf"),
    "Verdana": ("verdana.ttf", "verdana.ttf"),
    "VerdanaBold": ("verdanab.ttf", "verdanab.ttf"),
    "VerdanaBoldIt": ("verdanaz.ttf", "verdanaz.ttf"),
    "VerdanaIt": ("verdanai.ttf", "verdanai.ttf"),
    "Georgia": ("georgia.ttf", "georgia.ttf"),
    "GeorgiaBold": ("georgiab.ttf", "georgiab.ttf"),
    "GeorgiaBoldIt": ("georgiaz.ttf", "georgiaz.ttf"),
    "GeorgiaIt": ("georgiai.ttf", "georgiai.ttf"),
    "SegoeUI": ("segoeui.ttf", "segoeui.ttf"),
    "SegoeUIBold": ("segoeuib.ttf", "segoeuib.ttf"),
    "SegoeUIBoldIt": ("segoeuiz.ttf", "segoeuiz.ttf"),
    "SegoeUIIt": ("segoeuii.ttf", "segoeuii.ttf"),
    "Consolas": ("consola.ttf", "consola.ttf"),
    "ConsolasBold": ("consolab.ttf", "consolab.ttf"),
    "ConsolasBoldIt": ("consolaz.ttf", "consolaz.ttf"),
    "ConsolasIt": ("consolai.ttf", "consolai.ttf"),
    "DejaVuSansMono": ("DejaVuSansMono.ttf", "DejaVuSansMono.ttf"),
    "DejaVuSansMonoBold": ("DejaVuSansMono-Bold.ttf", "DejaVuSansMono-Bold.ttf"),
    "DejaVuSansMonoBoldIt": ("DejaVuSansMono-BoldOblique.ttf", "DejaVuSansMono-BoldOblique.ttf"),
    "DejaVuSansMonoIt": ("DejaVuSansMono-Oblique.ttf", "DejaVuSansMono-Oblique.ttf"),
    "IBMPlexMono": ("IBMPlexMono-Regular.ttf", "IBMPlexMono-Regular.ttf"),
    "IBMPlexMonoBold": ("IBMPlexMono-Bold.ttf", "IBMPlexMono-Bold.ttf"),
    "IBMPlexMonoBoldIt": ("IBMPlexMono-BoldItalic.ttf", "IBMPlexMono-BoldItalic.ttf"),
    "IBMPlexMonoIt": ("IBMPlexMono-Italic.ttf", "IBMPlexMono-Italic.ttf"),
    "LiberationMono": ("LiberationMono-Regular.ttf", "LiberationMono-Regular.ttf"),
    "LiberationMonoBold": ("LiberationMono-Bold.ttf", "LiberationMono-Bold.ttf"),
    "LiberationMonoBoldIt": ("LiberationMono-BoldItalic.ttf", "LiberationMono-BoldItalic.ttf"),
    "LiberationMonoIt": ("LiberationMono-Italic.ttf", "LiberationMono-Italic.ttf"),
}
# Lettera del colore nei codici del .gui -> RGB (g = verde, come v, per i .gui vecchi)
COLORI_TESTO = {
    "n": (0, 0, 0),
    "r": (255, 0, 0),
    "v": (0, 128, 0),
    "g": (0, 128, 0),
    "b": (0, 0, 255),
    "w": (255, 255, 255),
}
# fpdf salva i dati dei font in file .pkl; li mette nella cartella temporanea con un nome
# legato al percorso completo del font, così non si usano .pkl vecchi con percorsi sbagliati
fpdf.fpdf.FPDF_CACHE_MODE = 2
fpdf.fpdf.FPDF_CACHE_DIR = tempfile.gettempdir()
# Solo quando è avviato come programma (EditorModelli lo importa per l'anteprima)
if __name__ == "__main__":
    sys.exit(start())
