"""Modelli .gui: conversione tra la formattazione dei caratteri e i codici letti da SingleGUY.

Un modello è un .gui completo in cui il nome del testo e del documento contengono {nome}:
applicandolo a un file di testo (es. fattura_123.txt) {nome} diventa "fattura_123".

Struttura del .gui (una voce per riga):
    0 file di testo        1 azione (pdf, file, stampa, docx)   2 documento da creare
    3 template PDF         4 rotazione                          5 invio (0, email, telegram, whatsapp)
    6 destinatario (indirizzo email, "telegram <chat>", "whatsapp <numero>")
    7... formati, una linea per ogni riga del testo, fino a "fine"
    dopo "fine": dati dell'invio (email: oggetto, righe "allegato: <percorso>", riga vuota e testo;
                 WhatsApp: documento, eventuale "template: <nome> [<lingua>]" e testo o parametri)

Linea di formato: "<numero di riga 3 cifre>\\t<codice>" oppure, con più formati sulla riga,
"<numero>\\t<codice> <da> <a>\\t<codice> <da> <a>..." con da/a di 2 cifre (colonne 0-99).
Codice: <font><grassetto s/n><corsivo s/n><dimensione 2 cifre>[<colore>]       es. tsn12 o tsn12r
    font a larghezza fissa: c Courier, o Consolas, d DejaVu Sans Mono, p IBM Plex Mono, l Liberation Mono
    font proporzionali:     t Times, h Helvetica (Arial), k Calibri, v Verdana, g Georgia, u Segoe UI
    colore: n nero, r rosso, v verde, b blu, w bianco. In Word c'è sempre; nel PDF si può omettere
            (= nero), ma sulla stessa riga i codici devono essere tutti con o tutti senza colore.
Una riga senza codice mantiene il formato della riga prima.

Excel (azione "excel"): il testo ha i campi separati da ";", una riga del foglio per ogni riga.
Dalla linea 7: gli stili delle righe speciali (titoli, totali...), la linea "colonna", una linea
per ogni colonna (A, B, C...) e infine "fine". Le voci sono separate da tabulatori.
    Riga speciale: "<righe>[\\t<chiave>=<valore>...]"       es. 1-2, u1\\tnome=Totale\\tsfondo=giallo
        righe: numeri e intervalli separati da virgole; u1 = ultima riga, u2 = penultima...
        nome=<testo>            solo per l'editor
        carattere=<codice>      font della riga (codice come sopra, con il colore)
        sfondo=grigio|giallo|azzurro|verde|arancio
        allineamento=sinistra|centro|destra
        testo=si                i valori restano testo (non diventano numeri o date)
        Se una riga è in più stili vale l'ultimo.
    Colonna: "<formato numerico>[\\t<chiave>=<valore>...]"  es. 8\\tcarattere=kns11r\\tlarghezza=14
        formato numerico: numero di FORMATI_EXCEL (0 = generale)
        carattere=<codice>      font delle celle della colonna
        larghezza=<1-255>       larghezza della colonna in caratteri (senza: automatica)
        allineamento=sinistra|centro|destra
    Nelle righe speciali il carattere, lo sfondo e l'allineamento della riga valgono più di quelli
    della colonna. Le chiavi sono facoltative e in qualsiasi ordine.
    Senza la linea "colonna" (i .gui di prima) sono tutte colonne e la prima riga resta testo.
"""
import datetime
import os

from piattaforma import leggi_testo

# Font: lettera del codice -> (nome mostrato, nome in Word/Excel e nell'editor).
# Prima quelli a larghezza fissa, che mantengono le colonne del testo (i più utili nei modelli)
FONT = {
    "c": ("Courier", "Courier New"),
    "o": ("Consolas", "Consolas"),
    "d": ("DejaVu Sans Mono", "DejaVu Sans Mono"),
    "p": ("IBM Plex Mono", "IBM Plex Mono"),
    "l": ("Liberation Mono", "Liberation Mono"),
    "t": ("Times", "Times New Roman"),
    "h": ("Helvetica", "Arial"),
    "k": ("Calibri", "Calibri"),
    "v": ("Verdana", "Verdana"),
    "g": ("Georgia", "Georgia"),
    "u": ("Segoe UI", "Segoe UI"),
}
# Font a larghezza fissa (tutte le lettere e gli spazi larghi uguali)
FONT_FISSI = ["c", "o", "d", "p", "l"]
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
# Come gli editor mostrano azioni e invii (nel .gui restano i codici qui sopra)
DESCRIZIONI_AZIONI = {
    "file": "Crea solo il file",
    "pdf": "Crea e apri a schermo",
    "stampa": "Crea e stampa",
    "docx": "Crea il documento Word",
}
DESCRIZIONI_INVII = {"0": "Nessun invio", "email": "Email", "telegram": "Telegram", "whatsapp": "WhatsApp"}
# Le colonne nei codici hanno 2 cifre: un formato può cambiare solo fino alla colonna 99
COLONNA_MASSIMA = 99
SEGNAPOSTO = "{nome}"

# Formato: (font, grassetto, corsivo, dimensione, colore)
PREDEFINITO = ("t", False, False, 10, "n")

# --- Excel ---------------------------------------------------------------------
# Formati numerici: codice -> (descrizione, formato di Excel). Nel file xlsx i formati vanno
# scritti in notazione inglese: Excel li mostra poi con le impostazioni italiane del PC.
FORMATI_EXCEL = {
    0: ("Generale (il testo così com'è)", ""),
    1: ("Numero intero  1234", "0"),
    2: ("Numero con 2 decimali  1234,56", "0.00"),
    3: ("Intero con migliaia  1.234", "#,##0"),
    4: ("Decimali con migliaia  1.234,56", "#,##0.00"),
    5: ("Valuta intera  1.234 €", '#,##0 "€";-#,##0 "€"'),
    6: ("Valuta intera, negativi in rosso", '#,##0 "€";[Red]-#,##0 "€"'),
    7: ("Valuta intera, negativi in rosso (come 6)", '#,##0 "€";[Red]-#,##0 "€"'),
    8: ("Valuta con decimali, negativi in rosso  1.234,56 €", '#,##0.00 "€";[Red]-#,##0.00 "€"'),
    9: ("Percentuale  12%", "0%"),
    10: ("Percentuale con decimali  12,50%", "0.00%"),
    11: ("Scientifico  1,23E+03", "0.00E+00"),
    12: ("Data  20/10/2026", "dd/mm/yyyy"),
    13: ("Data  20-ott-2026", "dd-mmm-yyyy"),
    14: ("Giorno e mese  20-ott", "dd-mmm"),
    15: ("Mese e anno  ott 2026", "mmm yyyy"),
    16: ("Ora  1:30 PM", "h:mm AM/PM"),
    17: ("Ora con secondi  1:30:55 PM", "h:mm:ss AM/PM"),
    18: ("Ora  13:30", "hh:mm"),
    19: ("Ora con secondi  13:30:55", "hh:mm:ss"),
    20: ("Data e ora  20/10/2026 13:30", "dd/mm/yyyy hh:mm"),
    21: ("Ingegneristico  1,2E+3", "##0.0E+0"),
    22: ("Minuti e secondi  30:55", "mm:ss"),
    23: ("Testo", "@"),
    24: ("Valuta con decimali  1.234,56 €", '#,##0.00 "€";-#,##0.00 "€"'),
}
# Quali formati trasformano il testo in numero, data o ora
CODICI_NUMERO = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 21, 24]
CODICI_DATA = [12, 13, 14, 15, 20]
CODICI_ORA = [16, 17, 18, 19, 22]
ALLINEAMENTI = {"sinistra": "left", "centro": "center", "destra": "right"}
# Font predefinito di Excel (Calibri 11 nero): una colonna senza "carattere" resta così
PREDEFINITO_EXCEL = ("k", False, False, 11, "n")
# Sfondi delle righe speciali (colori chiari, su cui il testo nero si legge): nome -> (descrizione, RRGGBB)
SFONDI = {
    "grigio": ("grigio chiaro", "D9D9D9"),
    "giallo": ("giallo chiaro", "FFF2CC"),
    "azzurro": ("azzurro", "DDEBF7"),
    "verde": ("verde chiaro", "E2EFDA"),
    "arancio": ("arancio chiaro", "FCE4D6"),
}


def colonna_excel_vuota():
    return {"formato": 0, "carattere": None, "larghezza": None, "allineamento": None}


def stile_riga_vuoto():
    return {"righe": "", "nome": "", "carattere": None, "sfondo": None, "allineamento": None, "testo": False}


def _parti(linea):
    """Prima parte e coppie chiave=valore di una linea Excel (separate da tabulatori)."""
    parti = linea.rstrip("\r\n").split("\t")
    voci = {}
    for parte in parti[1:]:
        chiave, _, valore = parte.partition("=")
        voci[chiave.strip().lower()] = valore.strip()
    return parti[0].strip(), voci


def _leggi_comuni(voci, risultato):
    """Carattere e allineamento, comuni a colonne e righe; i valori sbagliati si ignorano."""
    if voci.get("carattere"):
        risultato["carattere"] = leggi_codice(voci["carattere"], PREDEFINITO_EXCEL)
    if voci.get("allineamento", "").lower() in ALLINEAMENTI:
        risultato["allineamento"] = voci["allineamento"].lower()


def leggi_colonna_excel(linea):
    """Legge una linea di colonna Excel; i valori sbagliati o sconosciuti vengono ignorati."""
    colonna = colonna_excel_vuota()
    primo, voci = _parti(linea)
    try:
        formato = int(primo)
        colonna["formato"] = formato if formato in FORMATI_EXCEL else 0
    except ValueError:
        pass
    _leggi_comuni(voci, colonna)
    larghezza = voci.get("larghezza", "")
    if larghezza.isdigit() and 1 <= int(larghezza) <= 255:
        colonna["larghezza"] = int(larghezza)
    return colonna


def linea_colonna_excel(colonna):
    parti = [str(colonna["formato"])]
    if colonna["carattere"] is not None:
        # Con il colore sempre (6 caratteri), come nel Word
        parti.append("carattere=" + codice(colonna["carattere"], True))
    if colonna["larghezza"]:
        parti.append("larghezza=%d" % colonna["larghezza"])
    if colonna["allineamento"]:
        parti.append("allineamento=" + colonna["allineamento"])
    return "\t".join(parti)


def leggi_stile_riga(linea):
    stile = stile_riga_vuoto()
    stile["righe"], voci = _parti(linea)
    stile["nome"] = voci.get("nome", "")
    _leggi_comuni(voci, stile)
    if voci.get("sfondo", "").lower() in SFONDI:
        stile["sfondo"] = voci["sfondo"].lower()
    stile["testo"] = voci.get("testo", "").lower() in ("si", "sì")
    return stile


def linea_stile_riga(stile):
    parti = [stile["righe"]]
    if stile["nome"]:
        # Il nome serve solo all'editor; un tabulatore lo spezzerebbe
        parti.append("nome=" + " ".join(stile["nome"].split()))
    if stile["carattere"] is not None:
        parti.append("carattere=" + codice(stile["carattere"], True))
    if stile["sfondo"]:
        parti.append("sfondo=" + stile["sfondo"])
    if stile["allineamento"]:
        parti.append("allineamento=" + stile["allineamento"])
    if stile["testo"]:
        parti.append("testo=si")
    return "\t".join(parti)


def _numero_riga(testo, numero_righe):
    if testo.startswith("u"):
        # u1 = ultima riga, u2 = penultima...
        distanza = int(testo[1:] or "1")
        if distanza < 1:
            raise ValueError
        return numero_righe - distanza + 1
    riga = int(testo)
    if riga < 1:
        raise ValueError
    return riga


def righe_excel(testo, numero_righe):
    """Le righe del foglio (da 1) indicate da un testo come "1-3, 10, u1".

    Numeri e intervalli separati da virgole; u1 è l'ultima riga, u2 la penultima e così via.
    Le righe che il foglio non ha si ignorano. ValueError se il testo non è valido.
    """
    risultato = set()
    for parte in testo.replace(";", ",").split(","):
        parte = parte.replace(" ", "").lower()
        if parte == "":
            continue
        try:
            inizio, _, fine = parte.partition("-")
            da = _numero_riga(inizio, numero_righe)
            a = _numero_riga(fine, numero_righe) if fine else da
        except ValueError:
            raise ValueError("Righe non valide: \"%s\" (esempi: 1-3, 10, u1 = ultima riga)" % parte)
        risultato.update(range(max(1, min(da, a)), min(numero_righe, max(da, a)) + 1))
    return risultato


def leggi_excel(linee):
    """Stili delle righe e colonne di un .gui Excel (le linee dalla 7 in poi).

    Prima gli stili delle righe, poi la linea "colonna", poi le colonne fino a "fine".
    Senza la linea "colonna" sono tutte colonne (i .gui di prima) e la prima riga resta testo,
    come allora.
    """
    stili, colonne, dopo_colonna = [], [], False
    for linea in linee:
        parola = linea.strip().lower()
        if parola[0:4] == "fine":
            break
        if parola == "colonna":
            dopo_colonna = True
            continue
        (colonne if dopo_colonna else stili).append(linea)
    if not dopo_colonna:
        return [dict(stile_riga_vuoto(), righe="1", testo=True)], [leggi_colonna_excel(l) for l in stili]
    return [leggi_stile_riga(l) for l in stili], [leggi_colonna_excel(l) for l in colonne]


def linee_excel(stili, colonne):
    """Linee di formato di un .gui Excel. Le ultime colonne senza formato si omettono: dopo
    "fine" il servizio usa il formato generale."""
    colonne = list(colonne)
    while colonne and colonne[-1] == colonna_excel_vuota():
        colonne.pop()
    return [linea_stile_riga(s) for s in stili] + ["colonna"] + [linea_colonna_excel(c) for c in colonne] + ["fine"]


def valore_excel(testo, codice):
    #Trasforma il testo nel valore giusto per il formato; se non ci riesce lo lascia come testo
    #(la usano il servizio e l'editor dei modelli Excel per mostrare le celle)
    testo = testo.strip()
    try:
        if codice in CODICI_NUMERO:
            #Numeri all'italiana: 1.234,56 -> 1234.56 (i punti delle migliaia si tolgono)
            numero = testo.replace("€", "").replace("%", "").strip()
            if "," in numero:
                numero = numero.replace(".", "").replace(",", ".")
            valore = float(numero)
            if codice in [9, 10] and "%" in testo:
                valore = valore / 100
            return valore
        if codice in CODICI_DATA:
            for formato_data in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y"):
                try:
                    return datetime.datetime.strptime(testo, formato_data)
                except ValueError:
                    pass
        if codice in CODICI_ORA:
            for formato_ora in ("%H:%M:%S", "%H:%M"):
                try:
                    return datetime.datetime.strptime(testo, formato_ora).time()
                except ValueError:
                    pass
    except ValueError:
        pass
    return testo


def campi_excel(riga):
    """I campi di una riga del testo Excel, come li legge il servizio: ogni campo finisce con ";"
    (il testo dopo l'ultimo ";" non è un campo)."""
    return riga.rstrip("\n").split(";")[:-1]


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


def testo_segnaposto(righe, in_piu=3, simboli=True, carattere="x"):
    """Il testo con "x" al posto dei caratteri, per preparare un modello valido per tutte le fatture.

    Un campo è un pezzo di riga con le parole separate da un solo spazio (due o più spazi separano
    i campi): diventa tutto x, spazi interni compresi. Con simboli=True punteggiatura e simboli
    (/ . , : - €...) restano, così si riconoscono date, importi e righe di separazione.
    in_piu: x aggiunte dopo ogni campo per fare spazio a testi più lunghi; occupano solo gli spazi
    che seguono, lasciandone almeno uno, quindi i campi dopo restano nella stessa colonna.
    """
    import re
    risultato = []
    for riga in righe:
        caratteri = list(riga)
        campi = list(re.finditer(r"\S+(?: \S+)*", riga))
        for n, campo in enumerate(campi):
            cambiato = False
            for posizione in range(campo.start(), campo.end()):
                c = riga[posizione]
                if c == " " or not simboli or c.isalnum():
                    caratteri[posizione] = carattere
                    cambiato = True
            if not cambiato:
                # Solo simboli (es. una riga di trattini): resta com'è
                continue
            if n + 1 < len(campi):
                spazio = campi[n + 1].start() - campo.end() - 1
            else:
                spazio = in_piu
            aggiunte = max(0, min(in_piu, spazio))
            fine = campo.end()
            caratteri[fine:fine + aggiunte] = [carattere] * aggiunte
        risultato.append("".join(caratteri))
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


def invio_da_mostrare(intestazione):
    """Destinatario e dati dell'invio come li mostra l'editor, più la scelta WhatsApp.

    Nell'editor il destinatario di Telegram e WhatsApp è senza la parola "telegram"/"whatsapp" e
    il template WhatsApp si sceglie a parte (nel .gui è la riga "template: <nome> [<lingua>]" dopo
    quella del documento). Nei .gui di prima nome e lingua del template erano dopo il numero.
    Restituisce (destinatario, coda, whatsapp) con whatsapp = {"modo", "template", "lingua"}.
    """
    from whatsapp_meta import leggi_destinatario, leggi_riga_template
    invio, destinatario, coda = intestazione["invio"], intestazione["destinatario"].strip(), list(intestazione["coda"])
    whatsapp = {"modo": "messaggio", "template": "", "lingua": ""}
    parole = destinatario.split(None, 1)
    if invio in ("telegram", "whatsapp") and parole and parole[0].lower() == invio:
        destinatario = parole[1] if len(parole) > 1 else ""
    if invio == "whatsapp":
        numero, template, lingua = leggi_destinatario(destinatario)
        if template:
            destinatario = numero
            whatsapp = {"modo": "template", "template": template, "lingua": lingua}
        riga = leggi_riga_template(coda[1]) if len(coda) > 1 else None
        if riga is not None:
            whatsapp = {"modo": "template", "template": riga[0], "lingua": riga[1]}
            del coda[1]
    return destinatario, coda, whatsapp


def invio_da_salvare(invio, destinatario, coda, whatsapp):
    """Il contrario di invio_da_mostrare(): (destinatario, coda) da scrivere nel .gui."""
    destinatario = destinatario.strip()
    coda = list(coda)
    if invio in ("telegram", "whatsapp") and destinatario and destinatario.split()[0].lower() != invio:
        destinatario = invio + " " + destinatario
    if invio == "whatsapp" and whatsapp["modo"] == "template":
        riga = "template: " + (whatsapp["template"].strip() + " " + whatsapp["lingua"].strip()).strip()
        coda = [coda[0] if coda else ""] + [riga] + coda[1:]
    return destinatario, coda


def aggiorna_email(intestazione):
    """Porta un modello email di prima al formato nuovo; restituisce True se l'ha cambiato.

    Prima l'indirizzo era la prima linea dopo "fine", seguita da oggetto, secondo allegato e
    testo; ora l'indirizzo è nel destinatario e dopo "fine" ci sono l'oggetto, le righe
    "allegato: <percorso>", una riga vuota e il testo.
    """
    coda = intestazione["coda"]
    if intestazione["invio"] != "email" or intestazione["destinatario"].strip() or not coda:
        return False
    intestazione["destinatario"] = coda[0].strip()
    allegato = coda[2].strip() if len(coda) > 2 else ""
    testo = coda[3:]
    intestazione["coda"] = ([coda[1] if len(coda) > 1 else ""] + (["allegato: " + allegato] if allegato else []) +
                            ([""] + testo if testo else []))
    return True


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
