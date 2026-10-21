"""Prova dei moduli di supporto: stampa_pdf, config_sicuro, whatsapp_meta, piattaforma, registro.

Uso:  python test\\test_moduli.py
Non stampa su carta: usa la stampante virtuale "Microsoft Print to PDF" con un file di uscita.
Non invia messaggi: le richieste a WhatsApp sono intercettate e solo controllate.
"""
import json
import os
import shutil
import sys
import tempfile
import traceback

import pymupdf

CARTELLA_PROGRAMMA = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, CARTELLA_PROGRAMMA)

risultati = []


def prova(funzione):
    try:
        funzione()
        risultati.append((funzione.__name__, None))
    except Exception:
        risultati.append((funzione.__name__, traceback.format_exc(limit=-3).strip()))
    return funzione


cartella = tempfile.mkdtemp(prefix="infas_test_")


@prova
def stampa_su_microsoft_print_to_pdf():
    # Stampante virtuale di Windows; su Linux la stampa è provata da stampa_linux_con_cups
    if os.name != "nt":
        return
    from stampa_pdf import stampa_pdf
    sorgente = os.path.join(cartella, "da_stampare.pdf")
    doc = pymupdf.open()
    for n in range(2):
        pagina = doc.new_page(width=595, height=842)
        pagina.insert_text((72, 100), "Pagina di prova %d" % (n + 1), fontsize=20)
    doc.save(sorgente)
    doc.close()
    uscita = os.path.join(cartella, "stampato.pdf")
    stampa_pdf(sorgente, "Microsoft Print to PDF", uscita)
    assert os.path.isfile(uscita), "la stampante virtuale non ha creato il file"
    stampato = pymupdf.open(uscita)
    assert stampato.page_count == 2, "pagine stampate: %d" % stampato.page_count
    stampato.close()


@prova
def config_cifrata_salva_e_rilegge():
    # Cifratura del sistema in uso: DPAPI su Windows, AES-GCM su Linux
    import config_sicuro
    percorso = os.path.join(cartella, "config.dat")
    dati = dict(config_sicuro.CAMPI, cartella_lavoro="C:\\lavoro", email_password="segreta è")
    config_sicuro.salva_config(dati, percorso)
    with open(percorso, "rb") as f:
        assert b"segreta" not in f.read(), "la password è in chiaro nel file"
    assert config_sicuro.carica_config(percorso) == dati
    config_sicuro.salva_config(dati, percorso, tutto_il_pc=True)
    assert config_sicuro.carica_config(percorso) == dati


@prova
def config_vera_leggibile():
    import config_sicuro
    if config_sicuro.config_esiste():
        with open(config_sicuro.FILE_CONFIG, "rb") as f:
            intestazione = f.read(9)
        # Un config.dat di Windows non si legge su Linux (e viceversa): va ricreato
        if intestazione != (config_sicuro._INTESTAZIONE if os.name == "nt" else config_sicuro._INTESTAZIONE_POSIX):
            return
        config = config_sicuro.carica_config()
        for campo in ("cartella_lavoro", "cartella_template", "cartella_output"):
            if config[campo] and not os.path.isdir(config[campo]):
                raise AssertionError("config.dat: la cartella %s non esiste: %s" % (campo, config[campo]))


@prova
def whatsapp_numeri_e_richieste():
    import whatsapp_meta
    assert whatsapp_meta.normalizza_numero("+39 333 123.45-67") == "393331234567"
    assert whatsapp_meta.normalizza_numero("0039 3331234567") == "393331234567"
    try:
        whatsapp_meta.normalizza_numero("333abc")
        raise AssertionError("numero non valido accettato")
    except whatsapp_meta.WhatsAppError:
        pass

    inviati = []

    class Risposta:
        ok, status_code, text = True, 200, ""

        def __init__(self, dati):
            self.dati = dati

        def json(self):
            return self.dati

    wa = whatsapp_meta.WhatsApp("token", "12345")

    def post(url, json=None, data=None, files=None, timeout=None):
        inviati.append((url, json, data))
        if url.endswith("/media"):
            return Risposta({"id": "media-1"})
        return Risposta({"messages": [{"id": "wamid.1"}]})
    wa.session.post = post

    allegato = os.path.join(cartella, "fattura.pdf")
    with open(allegato, "wb") as f:
        f.write(b"%PDF-1.4\n")
    assert wa.invia_testo("+39 333 1234567", "ciao") == "wamid.1"
    wa.invia_documento("393331234567", allegato)
    wa.invia_template("393331234567", "fattura", "it", ["Mario", "123"], allegato)
    testo, media, documento, media2, template = inviati
    assert testo[0].endswith("/v23.0/12345/messages") and testo[1]["to"] == "393331234567"
    assert documento[1]["document"] == {"id": "media-1", "filename": "fattura.pdf"}
    componenti = template[1]["template"]["components"]
    assert componenti[0]["type"] == "header" and componenti[1]["parameters"][1]["text"] == "123"
    assert wa.session.headers["Authorization"] == "Bearer token"

    class Errore(Risposta):
        ok, status_code = False, 401
    wa.session.post = lambda *a, **k: Errore({"error": {"message": "Invalid OAuth", "code": 190}})
    try:
        wa.invia_testo("393331234567", "x")
        raise AssertionError("errore API non segnalato")
    except whatsapp_meta.WhatsAppError as e:
        assert "190" in str(e)


# --- Linux: su Windows le parti di Linux si provano simulando il sistema ---------

@prova
def config_linux_cifrata_salva_e_rilegge():
    import config_sicuro
    vecchio_nome, vecchio_xdg = os.name, os.environ.get("XDG_CONFIG_HOME")
    machine_id = os.path.join(cartella, "machine-id")
    with open(machine_id, "w") as f:
        f.write("0123456789abcdef0123456789abcdef\n")
    vecchi_id = config_sicuro._FILE_MACHINE_ID
    os.environ["XDG_CONFIG_HOME"] = os.path.join(cartella, "xdg")
    config_sicuro._FILE_MACHINE_ID = [os.path.join(cartella, "manca"), machine_id]
    percorso = os.path.join(cartella, "config_linux.dat")
    dati = dict(config_sicuro.CAMPI, cartella_lavoro="/srv/lavoro", email_password="segreta è")
    try:
        os.name = "posix"
        # Solo questo utente: chiave casuale in ~/.config/fileguy/config.key
        config_sicuro.salva_config(dati, percorso)
        with open(percorso, "rb") as f:
            contenuto = f.read()
        assert contenuto.startswith(b"INFASCFG\x02U") and b"segreta" not in contenuto
        assert config_sicuro.carica_config(percorso) == dati
        chiave = config_sicuro.file_chiave()
        assert os.path.isfile(chiave) and chiave.endswith(os.path.join("fileguy", "config.key"))
        # Un altro utente (chiave diversa) non la legge
        with open(chiave, "wb") as f:
            f.write(os.urandom(32))
        try:
            config_sicuro.carica_config(percorso)
            raise AssertionError("letta con la chiave sbagliata")
        except config_sicuro.ConfigError:
            pass
        # Tutto il PC: chiave dall'identificativo della macchina, nessun file di chiave
        os.remove(chiave)
        config_sicuro.salva_config(dati, percorso, tutto_il_pc=True)
        assert not os.path.isfile(chiave)
        assert config_sicuro.carica_config(percorso) == dati
        # Su un altro PC no
        with open(machine_id, "w") as f:
            f.write("ffffffffffffffffffffffffffffffff\n")
        try:
            config_sicuro.carica_config(percorso)
            raise AssertionError("letta su un altro PC")
        except config_sicuro.ConfigError:
            pass
        # Un config.dat di Windows su Linux dà un messaggio chiaro
        with open(percorso, "wb") as f:
            f.write(b"INFASCFG\x01" + b"x" * 40)
        try:
            config_sicuro.carica_config(percorso)
            raise AssertionError("config di Windows accettata")
        except config_sicuro.ConfigError as e:
            assert "Windows" in str(e)
    finally:
        os.name = vecchio_nome
        config_sicuro._FILE_MACHINE_ID = vecchi_id
        if vecchio_xdg is None:
            os.environ.pop("XDG_CONFIG_HOME", None)
        else:
            os.environ["XDG_CONFIG_HOME"] = vecchio_xdg


@prova
def stampa_linux_con_cups():
    import stampa_pdf
    import subprocess
    comandi = []

    class Risultato:
        returncode, stdout, stderr = 0, "", ""

    vecchi = stampa_pdf.WINDOWS, stampa_pdf.shutil.which, stampa_pdf.subprocess.run
    try:
        stampa_pdf.WINDOWS = False
        stampa_pdf.shutil.which = lambda nome: "/usr/bin/" + nome
        stampa_pdf.subprocess.run = lambda comando, **k: comandi.append(comando) or Risultato()
        stampa_pdf.stampa_pdf("/tmp/fattura 1.pdf")
        stampa_pdf.stampa_pdf("/tmp/fattura 1.pdf", "Ufficio")
        assert comandi[0] == ["lp", "-o", "print-scaling=auto", "--", "/tmp/fattura 1.pdf"], comandi[0]
        assert comandi[1][:5] == ["lp", "-o", "print-scaling=auto", "-d", "Ufficio"], comandi[1]
        # Senza lp usa lpr
        stampa_pdf.shutil.which = lambda nome: "/usr/bin/lpr" if nome == "lpr" else None
        stampa_pdf.stampa_pdf("/tmp/a.pdf", "Ufficio")
        assert comandi[2][0] == "lpr" and "-P" in comandi[2]
        # Errore di CUPS: eccezione con il messaggio
        stampa_pdf.shutil.which = lambda nome: "/usr/bin/" + nome
        Risultato.returncode, Risultato.stderr = 1, "lp: No default destination."
        try:
            stampa_pdf.stampa_pdf("/tmp/a.pdf")
            raise AssertionError("errore di stampa non segnalato")
        except RuntimeError as e:
            assert "No default destination" in str(e)
        # Elenco delle stampanti da lpstat -e
        Risultato.returncode, Risultato.stdout = 0, "Ufficio\nMagazzino_HP\n"
        assert stampa_pdf.elenca_stampanti() == ["Magazzino_HP", "Ufficio"]
        # Nessun comando di stampa
        stampa_pdf.shutil.which = lambda nome: None
        try:
            stampa_pdf.stampa_pdf("/tmp/a.pdf")
            raise AssertionError("manca CUPS ma non è segnalato")
        except RuntimeError as e:
            assert "CUPS" in str(e)
    finally:
        stampa_pdf.WINDOWS, stampa_pdf.shutil.which, stampa_pdf.subprocess.run = vecchi
    assert subprocess.run is vecchi[2]


@prova
def stampanti_elenco_windows():
    if os.name != "nt":
        return
    import stampa_pdf
    assert "Microsoft Print to PDF" in stampa_pdf.elenca_stampanti()


@prova
def apri_pdf_a_schermo_linux():
    import piattaforma
    avviati = []
    vecchi = piattaforma.WINDOWS, piattaforma.MACOS, piattaforma.subprocess.Popen, dict(os.environ)
    try:
        piattaforma.WINDOWS = piattaforma.MACOS = False
        piattaforma.subprocess.Popen = lambda comando, **k: avviati.append((comando, k))
        os.environ.pop("DISPLAY", None)
        os.environ.pop("WAYLAND_DISPLAY", None)
        try:
            piattaforma.apri_file("/tmp/a.pdf")
            raise AssertionError("senza schermo non segnala l'errore")
        except RuntimeError as e:
            assert "DISPLAY" in str(e)
        os.environ["DISPLAY"] = ":0"
        piattaforma.apri_file("/tmp/a.pdf")
        assert avviati[0][0] == ["xdg-open", "/tmp/a.pdf"] and avviati[0][1]["start_new_session"]
        # Eseguibile di PyInstaller: LD_LIBRARY_PATH torna quello di prima
        sys.frozen = True
        os.environ["LD_LIBRARY_PATH"] = "/tmp/_MEI123"
        os.environ["LD_LIBRARY_PATH_ORIG"] = "/usr/lib/extra"
        assert piattaforma.ambiente_esterno()["LD_LIBRARY_PATH"] == "/usr/lib/extra"
        del os.environ["LD_LIBRARY_PATH_ORIG"]
        assert "LD_LIBRARY_PATH" not in piattaforma.ambiente_esterno()
    finally:
        piattaforma.WINDOWS, piattaforma.MACOS, piattaforma.subprocess.Popen, ambiente = vecchi
        os.environ.clear()
        os.environ.update(ambiente)
        if hasattr(sys, "frozen"):
            del sys.frozen


@prova
def testi_di_windows_e_di_linux():
    import piattaforma
    casi = {
        "windows.txt": "Città, perché\r\nTotale: 1.234,56 €\r\n".encode("cp1252"),
        "linux.txt": "Città, perché\nTotale: 1.234,56 €\n".encode("utf-8"),
        "bom.txt": "﻿Città, perché\nTotale: 1.234,56 €\n".encode("utf-8"),
    }
    for nome, dati in casi.items():
        percorso = os.path.join(cartella, nome)
        with open(percorso, "wb") as f:
            f.write(dati)
        righe = piattaforma.leggi_righe(percorso)
        assert righe == ["Città, perché\n", "Totale: 1.234,56 €\n"], (nome, righe)
    # Il salto pagina resta dentro la riga (come readlines), non la spezza
    percorso = os.path.join(cartella, "salto.txt")
    with open(percorso, "wb") as f:
        f.write(b"pagina 1\x0cpagina 2\nultima")
    assert piattaforma.leggi_righe(percorso) == ["pagina 1\x0cpagina 2\n", "ultima"]


@prova
def font_mancanti_sostituiti():
    # Su Linux Calibri, Verdana, Segoe UI e Georgia di solito non ci sono
    sys.path.insert(0, os.path.join(CARTELLA_PROGRAMMA, "test"))
    import esegui_test
    servizio = esegui_test.carica_programma("SingleGUY")
    servizio.cerca_font = lambda nome: None
    fonts = os.path.join(CARTELLA_PROGRAMMA, "Fonts")
    assert servizio.trova_font("CalibriBold") == os.path.join(fonts, "Helvetica-Bold.ttf")
    assert servizio.trova_font("GeorgiaIt") == os.path.join(fonts, "timesi.ttf")
    assert servizio.trova_font("Courier") == os.path.join(fonts, "cour.ttf")
    servizio.CARTELLA_FONT = os.path.join(cartella, "nessuna")
    try:
        servizio.trova_font("Times")
        raise AssertionError("font mancante non segnalato")
    except FileNotFoundError:
        pass


@prova
def registro_scrive_e_ruota():
    import registro
    percorso = os.path.join(cartella, "log", "FileGUY.log")
    vecchio = registro.DIMENSIONE_MASSIMA
    try:
        registro.configura(percorso)
        registro.registra("a.txt", "pdf a.pdf")
        try:
            raise ValueError("prova di errore")
        except ValueError as e:
            registro.registra_errore("b.txt", "email a x@y.it", e)
        with open(percorso, encoding="utf-8") as f:
            righe = f.read().splitlines()
        assert righe[0].endswith(" | a.txt | pdf a.pdf | OK"), righe[0]
        assert righe[1].endswith(" | b.txt | email a x@y.it | ERRORE | ValueError: prova di errore"), righe[1]
        assert righe[2] == "    Traceback (most recent call last):", righe[2]
        registro.DIMENSIONE_MASSIMA = 10
        for n in range(5):
            registro.registra("c%d.txt" % n, "pdf")
        assert os.path.isfile(percorso + ".1") and os.path.isfile(percorso + ".3")
        assert not os.path.isfile(percorso + ".4")
        # Un log non scrivibile non blocca
        registro.configura(os.path.join(cartella, "manca", "\0non valido.log"))
        registro.registra("d.txt", "pdf")
    finally:
        registro.DIMENSIONE_MASSIMA = vecchio
        registro._percorso = None


@prova
def programmi_senza_console():
    # Negli exe senza console (editor, ConfigManager) stdout e stderr sono None: un avviso di PyPDF2
    # (es. "Xref table not zero-indexed" di alcuni template) faceva fallire l'anteprima
    import subprocess
    for modulo in ("EditorModelli", "EditorExcel", "ConfigManager"):
        # PyPDF2 cambia il gestore degli avvisi quando apre un PDF (PdfFileReader), poi lo usa
        codice = ("import io, sys, warnings\nsys.stdout = sys.stderr = None\nimport %s\n"
                  "import pymupdf, PyPDF2, PyPDF2.utils\n"
                  "doc = pymupdf.open()\ndoc.new_page()\n"
                  "PyPDF2.PdfFileReader(io.BytesIO(doc.tobytes()))\n"
                  "warnings.warn('Xref table not zero-indexed', PyPDF2.utils.PdfReadWarning)\n"
                  "print('scritto su stdout')\n" % modulo)
        # In un file (non con -c): PyPDF2 1.26 vuole un percorso nel nome del file che dà l'avviso
        script = os.path.join(cartella, "senza_console_%s.py" % modulo)
        with open(script, "w", encoding="utf-8") as f:
            f.write("import sys\nsys.path.insert(0, %r)\n" % CARTELLA_PROGRAMMA + codice)
        risultato = subprocess.run([sys.executable, script], cwd=CARTELLA_PROGRAMMA, timeout=120,
                                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        assert risultato.returncode == 0, "%s senza console: il programma si ferma (codice %d)" % (
            modulo, risultato.returncode)


@prova
def config_manager_si_avvia():
    import ConfigManager
    finestra = ConfigManager.ConfigManager()
    finestra.update()
    finestra.destroy()


shutil.rmtree(cartella, ignore_errors=True)
for nome, errore in risultati:
    print("[%s] %s" % ("OK " if errore is None else "ERR", nome))
    if errore:
        print("       " + errore.replace("\n", "\n       "))
print("\n%d/%d prove OK" % (sum(e is None for _, e in risultati), len(risultati)))
sys.exit(0 if all(e is None for _, e in risultati) else 1)
