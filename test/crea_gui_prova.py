"""Crea file .txt / .gui di prova nella cartella di lavoro, da far elaborare al programma vero.

Uso:  python test\\crea_gui_prova.py [cartella di lavoro] [--senza-finestre]

Senza cartella usa la cartella di lavoro di config.dat (es. C:\\prova).
I template vengono presi dalla cartella dei template di config.dat (fat.pdf, ord.pdf, ddtl.pdf).
Non crea casi che inviano email / Telegram / WhatsApp o che stampano su carta.
--senza-finestre salta il caso che apre una finestra (il PDF nel visualizzatore).
"""
import os
import shutil
import sys

from PIL import Image, ImageDraw

CARTELLA_PROGRAMMA = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, CARTELLA_PROGRAMMA)
from config_sicuro import carica_config


def colonne(*campi):
    """Allinea i campi alle colonne (posizione in caratteri) per scriverli in Courier sul modulo."""
    riga = ""
    for posizione, testo in campi:
        riga = riga.ljust(posizione) + testo
    return riga


def scrivi(cartella, nome, righe_txt, righe_gui):
    # Prima il .txt e poi il .gui: il programma parte solo quando trova il .gui
    if righe_txt is not None:
        with open(os.path.join(cartella, nome + ".txt"), "w") as f:
            f.write("\n".join(righe_txt) + "\n")
    with open(os.path.join(cartella, nome + ".gui"), "w") as f:
        f.write("\n".join(righe_gui) + "\n")
    print("creato", nome + ".gui")


def intestazione(nome, azione, uscita, template="", rotazione="0", invio="0", destinatario=""):
    return [nome + ".txt", azione, uscita, template, rotazione, invio, destinatario]


def crea(cartella, cartella_template, senza_finestre=False):
    # --- 1. PDF semplice con tutti i formati, aperto nel visualizzatore ---
    if not senza_finestre:
        testo = [
            "INFAS SERVICE - prova dei formati PDF",
            "Times grassetto 14 (riga sopra), Helvetica 10 (questa riga)",
            "",
            "Courier grassetto corsivo 12",
            "Cliente: Mario Rossi   Totale: 1.234,56 €",
            "Helvetica corsivo 10 con accenti: città, perché, più, così",
            "Riga senza formato: tiene il formato precedente",
        ]
        gui = intestazione("prova_pdf", "pdf", "prova_pdf.pdf") + [
            "001\ttsn14", "002\thnn10", "003", "004\tcss12", "005\ttnn11 00 08\ttsn11 10 23",
            "006\thns10", "007", "fine"]
        scrivi(cartella, "prova_pdf", testo, gui)

    # --- 2. Fattura sul modello fat.pdf (testo in Courier allineato alle colonne del modulo) ---
    testo = [""] * 5 + [
        colonne((50, "Spett.le")),
        colonne((50, "ROSSI MARIO S.R.L.")),
        colonne((50, "Via del Serrone, 12")),
        colonne((50, "47890 San Marino (RSM)")),
        colonne((50, "COE SM12345")),
    ] + [""] * 7 + [
        colonne((0, "VT-M6x20"), (19, "Viti testa esagonale M6x20"), (48, "PZ"), (52, "100"), (61, "0,15"),
                (85, "15,00")),
        colonne((0, "BL-M6"), (19, "Bulloni M6 zincati"), (48, "PZ"), (52, "50"), (61, "1,20"), (85, "60,00")),
        colonne((0, "RN-12"), (19, "Rondelle piane 12 mm"), (48, "PZ"), (52, "200"), (61, "0,05"),
                (72, "10%"), (85, "9,00")),
        colonne((0, "ASS-01"), (19, "Assistenza tecnica"), (48, "H"), (52, "2"), (61, "45,00"), (85, "90,00")),
    ]
    gui = intestazione("prova_fattura", "file", "prova_fattura.pdf", "fat.pdf") + [
        "001", "002", "003", "004", "005", "006\tcsn10", "007\tcnn10", "008\tcnn10", "009\tcnn10",
        "010\tcnn10", "011", "012", "013", "014", "015", "016", "017", "018\tcnn09", "019\tcnn09",
        "020\tcnn09", "021\tcnn09", "fine"]
    scrivi(cartella, "prova_fattura", testo, gui)

    # --- 3. Ordine sul modello orizzontale ord.pdf ---
    testo = [""] * 6 + [
        colonne((0, "VT-M6x20"), (13, "Viti testa esagonale M6x20"), (77, "100"), (84, "L2026-14")),
        colonne((0, "BL-M6"), (13, "Bulloni M6 zincati"), (77, "50"), (84, "L2026-09")),
        colonne((0, "RN-12"), (13, "Rondelle piane 12 mm"), (77, "200"), (84, "L2026-11")),
    ]
    gui = intestazione("prova_ordine", "file", "prova_ordine.pdf", "ord.pdf") + [
        "001", "002", "003", "004", "005", "006", "007\tcnn09", "008\tcnn09", "009\tcnn09", "fine"]
    scrivi(cartella, "prova_ordine", testo, gui)

    # --- 4. Excel: una linea di formato per colonna (linea 8 = colonna A) ---
    testo = [
        "Codice;Descrizione;Quantita;Prezzo;Data;",
        "1;Viti testa esagonale M6x20;100;0,15;20/10/2026;",
        "2;Bulloni M6 zincati;50;1,20;21/10/2026;",
        "3;Rondelle piane 12 mm;200;0,05;22/10/2026;",
        "",
        "4;Assistenza tecnica (dopo una riga vuota);2;45,00;23/10/2026;",
    ]
    gui = intestazione("prova_excel", "excel", "prova_excel.xlsx") + ["1", "0", "1", "8", "12", "fine"]
    scrivi(cartella, "prova_excel", testo, gui)

    # --- 5. Word, con la formattazione del .gui ---
    testo = [
        "INFAS SERVICE S.r.l.",
        "Via Roma 1 - 00100 Roma",
        "",
        "Comunicazione di prova",
        "Cliente: Mario Rossi   Totale: 1.234,56 €",
        "Testo in Arial corsivo: città, perché, più, così",
        "Riga senza formato",
    ]
    gui = intestazione("prova_word", "docx", "prova_word.docx") + [
        "001\ttsn14n", "002\thnn10b", "003", "004\tcss12r", "005\ttnn11n 00 08\ttsn11r 10 23",
        "006\thns10n", "007", "fine"]
    scrivi(cartella, "prova_word", testo, gui)

    # --- 6. Barcode su una copia del DDT (il PDF viene modificato nella cartella di lavoro) ---
    pdf_barcode = os.path.join(cartella, "prova_barcode_ddt.pdf")
    shutil.copy(os.path.join(cartella_template, "ddtl.pdf"), pdf_barcode)
    scrivi(cartella, "prova_barcode", None, [
        "prova_barcode_ddt", "barcode", "2",
        "8001234567895", "ean13", "430,95", "560,140", pdf_barcode, "0",
        "DDT-2026-123", "code128", "30,770", "230,800", pdf_barcode, "0"])

    # --- 7. Immagine (logo) su una copia della fattura ---
    pdf_immagine = os.path.join(cartella, "prova_immagine_fattura.pdf")
    shutil.copy(os.path.join(cartella_template, "fat.pdf"), pdf_immagine)
    logo = os.path.join(cartella, "prova_logo.png")
    immagine = Image.new("RGB", (300, 120), (255, 255, 255))
    disegno = ImageDraw.Draw(immagine)
    disegno.rectangle((4, 4, 295, 115), outline=(200, 30, 30), width=6)
    disegno.text((20, 45), "TIMBRO DI PROVA", fill=(200, 30, 30))
    immagine.save(logo)
    scrivi(cartella, "prova_immagine", None, [
        "prova_immagine_fattura", "immagini", "1",
        logo, "380,700", "530,760", pdf_immagine, "0"])


if __name__ == "__main__":
    argomenti = [a for a in sys.argv[1:] if not a.startswith("--")]
    config = carica_config()
    cartella = argomenti[0] if argomenti else config["cartella_lavoro"]
    crea(cartella, config["cartella_template"], "--senza-finestre" in sys.argv)
