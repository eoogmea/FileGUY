"""Modelli .gui: conversione tra la formattazione dei caratteri e i codici letti da SingleGUY.

Un modello è un .gui completo in cui il nome del testo e del documento contengono {nome}:
applicandolo a un file di testo (es. fattura_123.txt) {nome} diventa "fattura_123".

Struttura del .gui (una voce per riga):
    0 file di testo        1 azione (pdf, file, stampa, docx)   2 documento da creare
    3 template PDF         4 rotazione                          5 invio (0, email, telegram, whatsapp)
    6 destinatario         7... formati, una linea per ogni riga del testo, fino a "fine"
    dopo "fine": dati dell'invio (destinatario email, oggetto, testo del messaggio...)

Linea di formato: "<numero di riga 3 cifre>\\t<codice>" oppure, con più formati sulla riga,
"<numero>\\t<codice> <da> <a>\\t<codice> <da> <a>..." con da/a di 2 cifre (colonne 0-99).
Codice: <font><grassetto s/n><corsivo s/n><dimensione 2 cifre>[<colore>]       es. tsn12 o tsn12r
    font:   t Times, h Helvetica (Arial), c Courier, k Calibri, v Verdana, g Georgia, u Segoe UI
    colore: n nero, r rosso, v verde, b blu, w bianco. In Word c'è sempre; nel PDF si può omettere
            (= nero), ma sulla stessa riga i codici devono essere tutti con o tutti senza colore.
Una riga senza codice mantiene il formato della riga prima.
"""
import os

from piattaforma import leggi_testo

# Font: lettera del codice -> (nome mostrato, nome in Word e nell'editor)
FONT = {
    "t": ("Times", "Times New Roman"),
    "h": ("Helvetica", "Arial"),
    "c": ("Courier", "Courier New"),
    "k": ("Calibri", "Calibri"),
    "v": ("Verdana", "Verdana"),
    "g": ("Georgia", "Georgia"),
    "u": ("Segoe UI", "Segoe UI"),
}
# Colori: lettera del codice -> (nome, RRGGBB come in SingleGUY)
COLORI = {
    "n": ("nero", "000000"),
    "r": ("rosso", "FF0000"),
    "v": ("verde", "008000"),
    "b": ("blu", "0000FF"),
    "w": ("bianco", "FFFFFF"),
}
# Nei .gui vecchi il verde poteva essere "g"
COLORI_LETTI = dict(COLORI, g=COLORI["v"])
AZIONI_PDF = ["file", "pdf", "stampa"]
AZIONI_WORD = ["docx"]
INVII = ["0", "email", "telegram", "whatsapp"]
# Le colonne nei codici hanno 2 cifre: un formato può cambiare solo fino alla colonna 99
COLONNA_MASSIMA = 99
SEGNAPOSTO = "{nome}"

# Formato: (font, grassetto, corsivo, dimensione, colore)
PREDEFINITO = ("t", False, False, 10, "n")


class ErroreModello(Exception):
    pass


def codice(formato, con_colore):
    font, grassetto, corsivo, dimensione, colore = formato
    testo = font + ("s" if grassetto else "n") + ("s" if corsivo else "n") + "%02d" % dimensione
    return testo + colore if con_colore else testo


def leggi_codice(testo, precedente=PREDEFINITO):
    """Il contrario di codice(); un font o una dimensione non validi restano quelli precedenti.

    Senza la lettera del colore il testo è nero, come nel programma.
    """
    font = testo[0:1].lower() if testo[0:1].lower() in FONT else precedente[0]
    try:
        dimensione = int(testo[3:5])
    except ValueError:
        dimensione = precedente[3]
    colore = testo[5:6] if testo[5:6] in COLORI_LETTI else "n"
    if colore == "g":
        colore = "v"
    return (font, testo[1:2].lower() == "s", testo[2:3].lower() == "s", dimensione, colore)


def pezzi(formati_riga):
    """Raggruppa i caratteri consecutivi con lo stesso formato: [(da, a, formato), ...]."""
    risultato = []
    for n, formato in enumerate(formati_riga):
        if risultato and risultato[-1][2] == formato:
            risultato[-1] = (risultato[-1][0], n + 1, formato)
        else:
            risultato.append((n, n + 1, formato))
    return risultato


def linee_formato(formati, word):
    """Crea le linee di formato del .gui.

    formati: per ogni riga del testo la lista dei formati dei suoi caratteri.
    Le righe vuote non hanno codice (mantengono il formato della riga prima).
    Solleva ErroreModello se un formato cambia oltre la colonna 99.
    """
    linee = []
    errori = []
    for n, formati_riga in enumerate(formati):
        numero = "%03d" % (n + 1)
        gruppi = pezzi(formati_riga)
        # Nel PDF il colore si scrive solo se sulla riga ce n'è uno diverso dal nero:
        # così i modelli senza colori restano leggibili anche dai programmi vecchi
        con_colore = word or any(formato[4] != "n" for _, _, formato in gruppi)
        if not gruppi:
            linee.append(numero)
        elif len(gruppi) == 1:
            linee.append(numero + "\t" + codice(gruppi[0][2], con_colore))
        else:
            if gruppi[-1][0] > COLONNA_MASSIMA:
                errori.append(n + 1)
                continue
            parti = []
            for da, a, formato in gruppi:
                # L'ultimo pezzo arriva al massimo alla colonna 99: il resto della riga prende il suo formato
                parti.append(codice(formato, con_colore) + " %02d %02d" % (da, min(a, COLONNA_MASSIMA)))
            linee.append(numero + "\t" + "\t".join(parti))
    if errori:
        raise ErroreModello("Nelle righe " + ", ".join(str(e) for e in errori) + " il formato cambia dopo la "
                            "colonna %d: nel .gui le colonne hanno solo 2 cifre." % COLONNA_MASSIMA)
    return linee + ["fine"]


def leggi_formati(linee, righe_testo, word):
    """Applica le linee di formato di un .gui alle righe del testo, come fa SingleGUY.

    Restituisce per ogni riga la lista dei formati dei caratteri.
    Se il testo ha più righe dei formati si riparte dal primo formato (come nel programma).
    """
    formati_linee = []
    for linea in linee:
        if linea[0:4] == "fine":
            break
        formati_linee.append(linea.rstrip("\n"))
    risultato = []
    corrente = PREDEFINITO
    for n, riga in enumerate(righe_testo):
        riga = riga.rstrip("\n")
        linea = formati_linee[n % len(formati_linee)] if formati_linee else ""
        parte_codici = linea[4:]
        if parte_codici.strip() == "":
            risultato.append([corrente] * len(riga))
            continue
        # Codici di 6 caratteri in Word o con il colore, altrimenti di 5
        lunghezza_codice = 6 if word or parte_codici[5:6] in COLORI_LETTI else 5
        if len(parte_codici.rstrip()) <= lunghezza_codice:
            corrente = leggi_codice(parte_codici, corrente)
            risultato.append([corrente] * len(riga))
            continue
        # Più formati: "<codice> <da> <a>" separati da un carattere
        lunghezza_pezzo = lunghezza_codice + 6
        primo = leggi_codice(parte_codici[0:lunghezza_codice], corrente)
        formati_riga = [primo] * len(riga)
        ultimo, fine_ultimo = primo, 0
        for inizio in range(0, len(parte_codici), lunghezza_pezzo + 1):
            pezzo = parte_codici[inizio:inizio + lunghezza_pezzo]
            if len(pezzo) < lunghezza_codice:
                break
            formato = leggi_codice(pezzo[0:lunghezza_codice], ultimo)
            try:
                da = int(pezzo[lunghezza_codice + 1:lunghezza_codice + 4])
                a = int(pezzo[lunghezza_codice + 4:])
            except ValueError:
                da, a = fine_ultimo, len(riga)
            for c in range(da, min(a, len(riga))):
                formati_riga[c] = formato
            ultimo, fine_ultimo = formato, a
        # Dopo l'ultimo pezzo il testo prende il suo formato
        for c in range(fine_ultimo, len(riga)):
            formati_riga[c] = ultimo
        corrente = ultimo
        risultato.append(formati_riga)
    return risultato


def nuovo_modello(word):
    """Intestazione predefinita di un modello."""
    return {
        "testo": SEGNAPOSTO + ".txt",
        "azione": "docx" if word else "file",
        "uscita": SEGNAPOSTO + (".docx" if word else ".pdf"),
        "template": "",
        "rotazione": "0",
        "invio": "0",
        "destinatario": "",
        "coda": [],
    }


def leggi_gui(percorso):
    """Legge un .gui: restituisce (intestazione, linee di formato)."""
    # Il .gui può essere stato creato su Windows o su Linux
    righe = leggi_testo(percorso).split("\n")
    while len(righe) < 7:
        righe.append("")
    intestazione = {
        "testo": righe[0], "azione": righe[1], "uscita": righe[2], "template": righe[3],
        "rotazione": righe[4] or "0", "invio": righe[5] or "0", "destinatario": righe[6],
    }
    formati = []
    n = 7
    while n < len(righe):
        formati.append(righe[n])
        if righe[n][0:4] == "fine":
            break
        n += 1
    coda = righe[n + 1:]
    while coda and coda[-1] == "":
        coda.pop()
    intestazione["coda"] = coda
    return intestazione, formati


def testo_gui(intestazione, linee):
    righe = [intestazione["testo"], intestazione["azione"], intestazione["uscita"], intestazione["template"],
             intestazione["rotazione"], intestazione["invio"], intestazione["destinatario"]]
    righe += linee + list(intestazione.get("coda", []))
    return "\n".join(righe) + "\n"


def scrivi_gui(percorso, intestazione, linee):
    # Codifica predefinita del sistema (cp1252 su Windows, UTF-8 su Linux): SingleGUY le legge tutte e due
    with open(percorso, "w") as f:
        f.write(testo_gui(intestazione, linee))


def applica(percorso_modello, file_testo, cartella_lavoro):
    """Crea nella cartella di lavoro il .gui per un file di testo, partendo dal modello.

    Il file di testo viene copiato nella cartella di lavoro se non c'è già.
    Restituisce il percorso del .gui creato.
    """
    intestazione, linee = leggi_gui(percorso_modello)
    nome = os.path.splitext(os.path.basename(file_testo))[0]
    intestazione = {k: (v.replace(SEGNAPOSTO, nome) if isinstance(v, str) else [r.replace(SEGNAPOSTO, nome)
                                                                                   for r in v])
                    for k, v in intestazione.items()}
    intestazione["testo"] = nome + ".txt"
    destinazione_testo = os.path.join(cartella_lavoro, nome + ".txt")
    if os.path.normcase(os.path.abspath(file_testo)) != os.path.normcase(os.path.abspath(destinazione_testo)):
        with open(file_testo, "rb") as f:
            contenuto = f.read()
        with open(destinazione_testo, "wb") as f:
            f.write(contenuto)
    # Prima il testo e poi il .gui: il programma parte solo quando trova il .gui
    percorso_gui = os.path.join(cartella_lavoro, nome + ".gui")
    scrivi_gui(percorso_gui, intestazione, linee)
    return percorso_gui
