"""Prova degli eseguibili già creati (SingleGUY.exe, back.exe, gli editor, ConfigManager.exe).

Uso:  python test\\test_eseguibili.py [cartella degli exe, es. C:\\prova]
Gli exe vengono copiati in una cartella temporanea con un config.dat suo (cartelle di lavoro, template e
output temporanee): la cartella di prova e la sua cartella di lavoro non vengono toccate.

1. SingleGUY.exe elabora PDF su template con i font a larghezza fissa, Word, Excel con le righe speciali,
   barcode e un errore voluto; nessun avviso sulla console, esiti nel log.
2. back.exe resta acceso come servizio: un documento messo nella cartella di lavoro viene elaborato da solo.
3. EditorModelli.exe, EditorExcel.exe e ConfigManager.exe si aprono e restano aperti (nessun errore all'avvio).
"""
import datetime
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import traceback

CARTELLA_PROGRAMMA = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, CARTELLA_PROGRAMMA)
os.chdir(CARTELLA_PROGRAMMA)

if os.name != "nt":
    print("Gli eseguibili .exe si provano su Windows.")
    sys.exit(0)

import config_sicuro
import modello_gui
import pymupdf

CARTELLA_EXE = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else r"C:\prova")
PROGRAMMI = ["SingleGUY.exe", "back.exe", "EditorModelli.exe", "EditorExcel.exe", "ConfigManager.exe"]
# I sorgenti da cui è creato ogni exe: se uno è più recente dell'exe, l'exe va ricompilato
COMUNI = ["config_sicuro.py", "piattaforma.py"]
SERVIZIO = ["SingleGUY.py", "modello_gui.py", "registro.py", "stampa_pdf.py", "whatsapp_meta.py"] + COMUNI
SORGENTI = {
    "SingleGUY.exe": SERVIZIO,
    "back.exe": ["back.py", "registro.py"] + COMUNI,
    "EditorModelli.exe": ["EditorModelli.py"] + SERVIZIO,
    "EditorExcel.exe": ["EditorExcel.py", "EditorModelli.py"] + SERVIZIO,
    "ConfigManager.exe": ["ConfigManager.py", "stampa_pdf.py", "whatsapp_meta.py"] + COMUNI,
}

risultati = []
note = []
base = tempfile.mkdtemp(prefix="fileguy_exe_")
lavoro, template, uscita = (os.path.join(base, n) for n in ("lavoro", "template", "output"))
for cartella in (lavoro, template, uscita):
    os.makedirs(cartella)
LOG = os.path.join(base, "FileGUY.log")


def prova(funzione):
    try:
        funzione()
        risultati.append((funzione.__name__, None))
    except Exception:
        risultati.append((funzione.__name__, traceback.format_exc(limit=-3).strip()))
    return funzione


def nota(testo):
    note.append(testo)


def chiudi(processo):
    # back.exe e gli exe di PyInstaller avviano altri processi: si chiude tutto l'albero
    subprocess.run(["taskkill", "/T", "/F", "/PID", str(processo.pid)], capture_output=True)


def righe_log():
    if not os.path.isfile(LOG):
        return []
    with open(LOG, encoding="utf-8") as f:
        return [r for r in f.read().splitlines() if not r.startswith("    ")]


def coppia(nome, righe, linee, **impostazioni):
    """Un .txt e il suo .gui nella cartella di lavoro: prima il testo, poi il .gui (come il gestionale)."""
    estensione = impostazioni.pop("estensione")
    with open(os.path.join(lavoro, nome + ".txt"), "w", encoding="cp1252") as f:
        f.write("\n".join(righe) + "\n")
    intestazione = {"testo": nome + ".txt", "azione": "file", "uscita": nome + estensione, "template": "",
                    "rotazione": "0", "invio": "0", "destinatario": "", "coda": []}
    intestazione.update(impostazioni)
    modello_gui.scrivi_gui(os.path.join(lavoro, nome + ".gui"), intestazione, linee)


@prova
def preparazione():
    mancanti = [p for p in PROGRAMMI if not os.path.isfile(os.path.join(CARTELLA_EXE, p))]
    assert not mancanti, "mancano in %s: %s" % (CARTELLA_EXE, ", ".join(mancanti))
    for programma in PROGRAMMI:
        shutil.copy2(os.path.join(CARTELLA_EXE, programma), base)
    fonts = os.path.join(CARTELLA_EXE, "Fonts")
    shutil.copytree(fonts if os.path.isdir(fonts) else os.path.join(CARTELLA_PROGRAMMA, "Fonts"),
                    os.path.join(base, "Fonts"))
    # Eseguibili più vecchi dei loro sorgenti: si provano lo stesso, ma va detto
    for programma in PROGRAMMI:
        data = os.path.getmtime(os.path.join(CARTELLA_EXE, programma))
        piu_recenti = [f for f in SORGENTI[programma]
                       if os.path.getmtime(os.path.join(CARTELLA_PROGRAMMA, f)) > data]
        nota("%-18s del %s%s" % (programma, datetime.datetime.fromtimestamp(data).strftime("%d/%m/%Y %H:%M"),
                                 "  <- %s cambiati dopo: ricompilalo" % ", ".join(piu_recenti) if piu_recenti else ""))
    # Template vero se c'è, altrimenti uno di prova
    vero = os.path.join(CARTELLA_EXE, "modelli", "fat.pdf")
    if os.path.isfile(vero):
        shutil.copy2(vero, os.path.join(template, "modulo.pdf"))
    else:
        documento = pymupdf.open()
        documento.new_page(width=595, height=842).draw_rect(pymupdf.Rect(20, 20, 575, 822), color=(0, 0, 1))
        documento.save(os.path.join(template, "modulo.pdf"))
    config_sicuro.salva_config(dict(config_sicuro.CAMPI, cartella_lavoro=lavoro, cartella_template=template,
                                    cartella_output=uscita, file_log=LOG), os.path.join(base, "config.dat"))
    nota("Cartella temporanea: %s" % base)


FATTURA = ["DITTA ESEMPIO S.r.l.                          FATTURA N. 123/2026",
           "Spett.le Mario Rossi S.p.A. - Città perché €",
           "A001      Viti autofilettanti 4x30         100      0,15       15,00",
           "B010      Tasselli nylon 8 mm               75      0,22       16,50",
           "                                   TOTALE EURO              123,83"]


@prova
def singleguy_elabora_i_documenti():
    fissi = modello_gui.FONT_FISSI
    formati = [[(fissi[n % len(fissi)], n == 0, n == 4, 10, "r" if n == 4 else "n")] * len(r)
               for n, r in enumerate(FATTURA)]
    coppia("pdf_su_template", FATTURA, modello_gui.linee_formato(formati, False), template="modulo.pdf",
           estensione=".pdf")
    coppia("word", FATTURA, modello_gui.linee_formato(formati, True), azione="docx", estensione=".docx")
    listino = ["Listino;;;", "Codice;Prezzo;Data;", "A001;1.234,50;20/10/2026;", "Totale;1.234,50;;"]
    stili = [dict(modello_gui.stile_riga_vuoto(), righe="1-2", carattere=("k", True, False, 12, "b"), testo=True),
             dict(modello_gui.stile_riga_vuoto(), righe="u1", carattere=("k", True, False, 11, "n"), sfondo="giallo")]
    colonne = [modello_gui.colonna_excel_vuota(), dict(modello_gui.colonna_excel_vuota(), formato=8),
               dict(modello_gui.colonna_excel_vuota(), formato=12)]
    coppia("excel", listino, modello_gui.linee_excel(stili, colonne), azione="excel", estensione=".xlsx")
    documento = pymupdf.open()
    documento.new_page(width=595, height=842).insert_text((72, 72), "Documento per il barcode")
    documento.save(os.path.join(lavoro, "barcode_doc.pdf"))
    with open(os.path.join(lavoro, "barcode.gui"), "w") as f:
        f.write("\n".join(["barcode", "barcode", "1", "123456789012", "ean13", "100,150", "300,250",
                           os.path.join(lavoro, "barcode_doc.pdf"), "0"]) + "\n")
    coppia("errore_voluto", FATTURA, modello_gui.linee_formato(formati, False), template="non_esiste.pdf",
           estensione=".pdf")

    risultato = subprocess.run([os.path.join(base, "SingleGUY.exe")], cwd=base, capture_output=True, text=True,
                               encoding="utf-8", errors="replace", timeout=300)
    console = risultato.stdout + risultato.stderr
    assert "FutureWarning" not in console and "fitz" not in console, "avvisi sulla console:\n" + console[-500:]
    # L'unico errore sulla console deve essere quello voluto
    errori_console = re.findall(r"Errore con il file (\S+)", console)
    assert errori_console == ["errore_voluto.txt:"], errori_console

    # PDF: font a larghezza fissa incorporati e testo con accenti ed euro
    pdf = pymupdf.open(os.path.join(uscita, "pdf_su_template.pdf"))
    incorporati = {re.sub(r"^[A-Z]{6}\+", "", f[3]).replace("-", "") for f in pdf[0].get_fonts()}
    for atteso in ("Consolas", "DejaVuSansMono", "IBMPlexMono", "LiberationMono", "CourierNew"):
        assert any(atteso in nome for nome in incorporati), (atteso, incorporati)
    assert "Città perché €" in pdf[0].get_text()
    nota("SingleGUY.exe: PDF su template con %d font, Word, Excel, barcode" % len(incorporati))
    # Word ed Excel
    from docx import Document
    assert "Consolas" in {r.font.name for r in Document(os.path.join(uscita, "word.docx")).paragraphs[0].runs}
    from openpyxl import load_workbook
    foglio = load_workbook(os.path.join(uscita, "excel.xlsx")).worksheets[0]
    assert foglio["B3"].value == 1234.5 and foglio["B4"].fill.start_color.rgb == "FFFFF2CC"
    assert foglio["A1"].font.color.rgb == "FF0000FF" and foglio["B2"].value == "Prezzo"
    # Barcode aggiunto al PDF della cartella di lavoro
    assert len(pymupdf.open(os.path.join(lavoro, "barcode_doc.pdf"))[0].get_images()) == 1
    # Log: tutto OK tranne l'errore voluto, i cui file restano nella cartella di lavoro
    log = righe_log()
    errori_log = [r for r in log if " | ERRORE" in r]
    assert len(errori_log) == 1 and "errore_voluto.txt" in errori_log[0], errori_log
    assert sum(" | OK" in r for r in log) == 4, log
    rimasti = sorted(f for f in os.listdir(lavoro) if f.endswith((".txt", ".gui")))
    assert rimasti == ["errore_voluto.gui", "errore_voluto.txt"], rimasti
    for nome in rimasti:
        os.remove(os.path.join(lavoro, nome))


@prova
def back_resta_acceso_e_elabora():
    processo = subprocess.Popen([os.path.join(base, "back.exe")], cwd=base, stdout=subprocess.DEVNULL,
                                stderr=subprocess.DEVNULL, creationflags=subprocess.CREATE_NEW_PROCESS_GROUP)
    try:
        # Il servizio parte e scrive nel log; poi arriva un documento
        time.sleep(4)
        assert processo.poll() is None, "back.exe si è chiuso subito (codice %s)" % processo.returncode
        coppia("dal_servizio", FATTURA, ["001\tcsn12", "fine"], estensione=".pdf")
        fine = time.time() + 120
        while time.time() < fine and not os.path.isfile(os.path.join(uscita, "dal_servizio.pdf")):
            time.sleep(1)
        assert os.path.isfile(os.path.join(uscita, "dal_servizio.pdf")), "back.exe non ha elaborato il documento"
        # I file elaborati vengono tolti dalla cartella di lavoro
        fine = time.time() + 20
        while time.time() < fine and os.path.isfile(os.path.join(lavoro, "dal_servizio.gui")):
            time.sleep(1)
        assert not os.path.isfile(os.path.join(lavoro, "dal_servizio.gui"))
        assert processo.poll() is None, "back.exe si è chiuso dopo il documento"
        log = righe_log()
        assert any("avvio del servizio (back)" in r for r in log), log
        assert any("dal_servizio.txt" in r and r.endswith("| OK") for r in log), log
        nota("back.exe: acceso come servizio, documento elaborato da solo e tolto dalla cartella di lavoro")
    finally:
        chiudi(processo)


@prova
def finestre_si_aprono():
    for programma in ("EditorModelli.exe", "EditorExcel.exe", "ConfigManager.exe"):
        processo = subprocess.Popen([os.path.join(base, programma)], cwd=base)
        try:
            # Un exe di PyInstaller si estrae e poi apre la finestra: se c'è un errore all'avvio si chiude
            time.sleep(10)
            assert processo.poll() is None, "%s si è chiuso all'avvio (codice %s)" % (programma, processo.returncode)
            nota("%-18s si apre e resta aperto" % programma)
        finally:
            chiudi(processo)


print("\n".join(note))
print()
for nome, errore in risultati:
    print("[%s] %s" % ("OK " if errore is None else "ERR", nome))
    if errore:
        print("       " + errore.replace("\n", "\n       "))
tutto_ok = all(e is None for _, e in risultati)
print("\n%d/%d prove OK (eseguibili di %s)" % (sum(e is None for _, e in risultati), len(risultati), CARTELLA_EXE))
if tutto_ok:
    shutil.rmtree(base, ignore_errors=True)
else:
    print("La cartella temporanea resta per controllare: %s" % base)
sys.exit(0 if tutto_ok else 1)
