"""Editor dei modelli .gui per fatture, bolle e altri documenti ripetitivi.

Avvio: python EditorModelli.py

1. File > Nuovo da testo di esempio: si apre un file di testo come quelli che arrivano al servizio.
2. Si seleziona una riga intera, una parte o anche solo alcune lettere e si sceglie carattere,
   dimensione, grassetto, corsivo e colore. Il testo non si modifica: cambia solo la formattazione.
3. Si scelgono tipo di documento, template PDF, rotazione ed eventuale invio.
4. Anteprima (F5): crea il documento con lo stesso codice del servizio.
5. File > Salva modello: crea un .gui con {nome} al posto del nome del file, più una copia
   del testo di esempio accanto (serve solo per riaprire il modello).
6. File > Applica a file di testo: per ogni .txt scelto crea il .gui nella cartella di lavoro,
   dove il servizio lo elabora.
La guida completa è Guida_EditorModelli.html (menu Aiuto, F1).
"""
import base64
import importlib
import os
import shutil
import sys
import tempfile
import tkinter as tk
import traceback
from tkinter import filedialog, messagebox, ttk

import modello_gui
from config_sicuro import carica_config
from modello_gui import COLONNA_MASSIMA, COLORI, FONT, PREDEFINITO

if getattr(sys, "frozen", False):
    CARTELLA_PROGRAMMA = os.path.dirname(sys.executable)
else:
    CARTELLA_PROGRAMMA = os.path.dirname(os.path.abspath(__file__))
CARTELLA_MODELLI = os.path.join(CARTELLA_PROGRAMMA, "ModelliGUI")
# La guida sta accanto al programma; nell'exe creato con PyInstaller anche dentro l'exe
GUIDA = "Guida_EditorModelli.html"

NOMI_FONT = {nome: lettera for lettera, (nome, _) in FONT.items()}
DIMENSIONI = ["6", "7", "8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "36"]
NOMI_COLORI = {nome: lettera for lettera, (nome, _) in COLORI.items()}
TASTI_MOVIMENTO = {"Left", "Right", "Up", "Down", "Home", "End", "Prior", "Next", "Shift_L", "Shift_R",
                   "Control_L", "Control_R", "Alt_L", "Alt_R", "Escape", "Tab"}


class EditorModelli(tk.Tk):
    def __init__(self):
        super().__init__()
        self.geometry("1150x760")
        self.minsize(900, 500)
        try:
            self.config_servizio = carica_config()
        except Exception:
            self.config_servizio = {"cartella_lavoro": "", "cartella_template": "", "cartella_output": ""}
        self.percorso_modello = None
        self.percorso_testo = None
        self.righe = []
        self.modificato = False
        self.immagini_anteprima = []

        self.word = tk.BooleanVar(value=False)
        self.azione = tk.StringVar(value="file")
        self.template = tk.StringVar()
        self.rotazione = tk.StringVar(value="0")
        self.uscita = tk.StringVar(value=modello_gui.SEGNAPOSTO + ".pdf")
        self.invio = tk.StringVar(value="0")
        self.destinatario = tk.StringVar()

        self._crea_menu()
        self._crea_impostazioni()
        self._crea_barra()
        self._crea_testo()
        self.stato = ttk.Label(self, text="File > Nuovo da testo di esempio per iniziare", foreground="gray",
                               padding=(8, 2))
        self.stato.pack(side="bottom", fill="x")
        for variabile in (self.azione, self.template, self.rotazione, self.uscita, self.invio, self.destinatario):
            variabile.trace_add("write", lambda *a: self._segna_modificato())
        self._cambia_tipo(iniziale=True)
        self._aggiorna_titolo()
        self.protocol("WM_DELETE_WINDOW", self._esci)

    # --- Interfaccia -------------------------------------------------------

    def _crea_menu(self):
        menu = tk.Menu(self)
        file = tk.Menu(menu, tearoff=False)
        file.add_command(label="Nuovo da testo di esempio...", command=self.nuovo, accelerator="Ctrl+N")
        file.add_command(label="Apri modello...", command=self.apri, accelerator="Ctrl+O")
        file.add_separator()
        file.add_command(label="Salva modello", command=self.salva, accelerator="Ctrl+S")
        file.add_command(label="Salva modello con nome...", command=lambda: self.salva(con_nome=True))
        file.add_separator()
        file.add_command(label="Applica a file di testo...", command=self.applica)
        file.add_separator()
        file.add_command(label="Esci", command=self._esci)
        menu.add_cascade(label="File", menu=file)
        menu.add_command(label="Anteprima (F5)", command=self.anteprima)
        aiuto = tk.Menu(menu, tearoff=False)
        aiuto.add_command(label="Guida all'uso", command=self.apri_guida, accelerator="F1")
        aiuto.add_command(label="Informazioni", command=self._informazioni)
        menu.add_cascade(label="Aiuto", menu=aiuto)
        self.configure(menu=menu)
        self.bind("<Control-n>", lambda e: self.nuovo())
        self.bind("<Control-o>", lambda e: self.apri())
        self.bind("<Control-s>", lambda e: self.salva())
        self.bind("<F5>", lambda e: self.anteprima())
        self.bind("<F1>", lambda e: self.apri_guida())

    def _crea_impostazioni(self):
        riquadro = ttk.LabelFrame(self, text="Documento", padding=(8, 4))
        riquadro.pack(side="top", fill="x", padx=6, pady=(6, 0))

        ttk.Label(riquadro, text="Tipo").grid(row=0, column=0, sticky="w")
        tipo = ttk.Frame(riquadro)
        tipo.grid(row=0, column=1, sticky="w")
        ttk.Radiobutton(tipo, text="PDF", variable=self.word, value=False,
                        command=self._cambia_tipo).pack(side="left")
        ttk.Radiobutton(tipo, text="Word", variable=self.word, value=True,
                        command=self._cambia_tipo).pack(side="left", padx=(6, 0))

        ttk.Label(riquadro, text="Azione").grid(row=0, column=2, sticky="w", padx=(16, 4))
        self.scelta_azione = ttk.Combobox(riquadro, textvariable=self.azione, width=8, state="readonly")
        self.scelta_azione.grid(row=0, column=3, sticky="w")

        ttk.Label(riquadro, text="Template PDF").grid(row=0, column=4, sticky="w", padx=(16, 4))
        self.scelta_template = ttk.Combobox(riquadro, textvariable=self.template, width=22,
                                            postcommand=self._elenca_template)
        self.scelta_template.grid(row=0, column=5, sticky="w")

        ttk.Label(riquadro, text="Rotazione").grid(row=0, column=6, sticky="w", padx=(16, 4))
        self.scelta_rotazione = ttk.Combobox(riquadro, textvariable=self.rotazione, width=5,
                                             values=["0", "90", "180", "270"], state="readonly")
        self.scelta_rotazione.grid(row=0, column=7, sticky="w")

        ttk.Label(riquadro, text="Documento creato").grid(row=1, column=0, sticky="w", pady=(4, 0))
        ttk.Entry(riquadro, textvariable=self.uscita, width=24).grid(row=1, column=1, columnspan=3, sticky="we",
                                                                     pady=(4, 0))
        ttk.Label(riquadro, text="Invio").grid(row=1, column=4, sticky="w", padx=(16, 4), pady=(4, 0))
        ttk.Combobox(riquadro, textvariable=self.invio, width=10, values=modello_gui.INVII,
                     state="readonly").grid(row=1, column=5, sticky="w", pady=(4, 0))
        ttk.Label(riquadro, text="Destinatario").grid(row=1, column=6, sticky="w", padx=(16, 4), pady=(4, 0))
        ttk.Entry(riquadro, textvariable=self.destinatario, width=28).grid(row=1, column=7, columnspan=2,
                                                                           sticky="we", pady=(4, 0))

        ttk.Label(riquadro, text="Dati dell'invio\n(righe dopo \"fine\")").grid(row=2, column=0, sticky="nw",
                                                                             pady=(4, 0))
        self.coda = tk.Text(riquadro, height=3, width=60, font=("Consolas", 9), undo=True)
        self.coda.grid(row=2, column=1, columnspan=8, sticky="we", pady=(4, 0))
        self.coda.bind("<<Modified>>", self._coda_modificata)
        ttk.Label(riquadro, text="{nome} = nome del file di testo, senza .txt", foreground="gray").grid(
            row=3, column=1, columnspan=8, sticky="w")
        riquadro.columnconfigure(8, weight=1)

    def _crea_barra(self):
        barra = ttk.Frame(self, padding=(6, 6))
        barra.pack(side="top", fill="x")

        ttk.Label(barra, text="Carattere").pack(side="left")
        # exportselection=False: usando le tendine il testo selezionato resta selezionato
        self.scelta_font = ttk.Combobox(barra, values=list(NOMI_FONT), width=10, state="readonly",
                                        exportselection=False)
        self.scelta_font.pack(side="left", padx=(4, 10))
        self.scelta_font.bind("<<ComboboxSelected>>",
                              lambda e: self._imposta("f", NOMI_FONT[self.scelta_font.get()]))

        ttk.Label(barra, text="Dimensione").pack(side="left")
        self.scelta_dimensione = ttk.Combobox(barra, values=DIMENSIONI, width=4, exportselection=False)
        self.scelta_dimensione.pack(side="left", padx=(4, 10))
        self.scelta_dimensione.bind("<<ComboboxSelected>>", lambda e: self._imposta_dimensione())
        self.scelta_dimensione.bind("<Return>", lambda e: self._imposta_dimensione())

        self.pulsante_b = tk.Button(barra, text="G", width=3, font=("Arial", 10, "bold"),
                                    command=lambda: self._alterna("b"))
        self.pulsante_b.pack(side="left")
        self.pulsante_i = tk.Button(barra, text="C", width=3, font=("Arial", 10, "italic"),
                                    command=lambda: self._alterna("i"))
        self.pulsante_i.pack(side="left", padx=(2, 10))

        self.etichetta_colore = ttk.Label(barra, text="Colore")
        self.etichetta_colore.pack(side="left")
        self.scelta_colore = ttk.Combobox(barra, values=list(NOMI_COLORI), width=7, state="readonly",
                                          exportselection=False)
        self.scelta_colore.pack(side="left", padx=(4, 10))
        self.scelta_colore.bind("<<ComboboxSelected>>",
                                lambda e: self._imposta("c", NOMI_COLORI[self.scelta_colore.get()]))

        ttk.Button(barra, text="Formato predefinito", command=self._predefinito).pack(side="left", padx=(6, 0))
        ttk.Button(barra, text="Anteprima", command=self.anteprima).pack(side="right")

    def _crea_testo(self):
        cornice = ttk.Frame(self)
        cornice.pack(side="top", fill="both", expand=True, padx=6)
        self.numeri = tk.Canvas(cornice, width=44, background="#f0f0f0", highlightthickness=0)
        self.numeri.pack(side="left", fill="y")
        verticale = ttk.Scrollbar(cornice, orient="vertical")
        verticale.pack(side="right", fill="y")
        orizzontale = ttk.Scrollbar(self, orient="horizontal")
        orizzontale.pack(side="bottom", fill="x", padx=(50, 22))
        self.scorrimento_verticale = verticale
        # Il testo senza formato usa il carattere predefinito, lo stesso del programma (Times 10)
        # exportselection=False: la selezione non si perde quando si clicca o si scrive in un altro campo,
        # e resta visibile (inactiveselectbackground) mentre si usano le tendine
        self.testo = tk.Text(cornice, wrap="none", undo=False, padx=8, pady=8, spacing1=2, spacing3=2,
                             exportselection=False, inactiveselectbackground="#9cc3ee",
                             font=(FONT[PREDEFINITO[0]][1], PREDEFINITO[3]), xscrollcommand=orizzontale.set,
                             yscrollcommand=self._scorrimento, insertwidth=2)
        self.testo.pack(side="left", fill="both", expand=True)
        verticale.configure(command=self._scorri)
        orizzontale.configure(command=self.testo.xview)
        # I caratteri oltre la colonna 99 non possono avere un formato diverso dal pezzo prima
        self.testo.tag_configure("oltre", background="#f4f4f4")

        self.testo.bind("<Key>", self._tasto)
        for evento in ("<<Paste>>", "<<Cut>>", "<<Clear>>"):
            self.testo.bind(evento, lambda e: "break")
        self.testo.bind("<Control-a>", lambda e: self._seleziona_tutto() or "break")
        for tasto in ("<Control-g>", "<Control-G>", "<Control-b>", "<Control-B>"):
            self.testo.bind(tasto, lambda e: self._alterna("b") or "break")
        for tasto in ("<Control-i>", "<Control-I>"):
            self.testo.bind(tasto, lambda e: self._alterna("i") or "break")
        for evento in ("<KeyRelease>", "<ButtonRelease-1>", "<B1-Motion>"):
            self.testo.bind(evento, lambda e: self._aggiorna_barra(), add="+")
        self.testo.bind("<Configure>", lambda e: self._disegna_numeri())

    # --- Numeri di riga ----------------------------------------------------

    def _scorrimento(self, inizio, fine):
        self.scorrimento_verticale.set(inizio, fine)
        self._disegna_numeri()

    def _scorri(self, *argomenti):
        self.testo.yview(*argomenti)
        self._disegna_numeri()

    def _disegna_numeri(self):
        self.numeri.delete("all")
        indice = self.testo.index("@0,0")
        while True:
            posizione = self.testo.dlineinfo(indice)
            if posizione is None:
                break
            riga = int(indice.split(".")[0])
            if riga > len(self.righe):
                break
            self.numeri.create_text(38, posizione[1] + posizione[3] // 2, anchor="e", text=str(riga),
                                    fill="gray", font=("Consolas", 9))
            indice = self.testo.index(indice + "+1line")
            if int(indice.split(".")[0]) == riga:
                break

    # --- Formattazione -----------------------------------------------------

    def _formato(self, tags):
        f, b, i, s, c = PREDEFINITO
        for tag in tags:
            if tag == "b":
                b = True
            elif tag == "i":
                i = True
            elif tag.startswith("f:"):
                f = tag[2:]
            elif tag.startswith("s:"):
                s = int(tag[2:])
            elif tag.startswith("c:"):
                c = tag[2:]
        return (f, b, i, s, c)

    def _tag_da_formato(self, formato):
        f, b, i, s, c = formato
        tags = ["f:" + f, "s:" + str(s)]
        if b:
            tags.append("b")
        if i:
            tags.append("i")
        if c != "n":
            tags.append("c:" + c)
        return tags

    def _aggiorna_vista(self):
        # Ricalcola i tag di visualizzazione combinando i tag di formato di ogni pezzo di testo
        for tag in self.testo.tag_names():
            if tag.startswith("v|"):
                self.testo.tag_remove(tag, "1.0", "end")
        attivi = set()
        for chiave, valore, indice in self.testo.dump("1.0", "end-1c", tag=True, text=True):
            if chiave == "tagon":
                attivi.add(valore)
            elif chiave == "tagoff":
                attivi.discard(valore)
            elif chiave == "text":
                logici = [t for t in attivi if t[:2] in ("f:", "s:", "c:") or t in ("b", "i")]
                if not logici:
                    continue
                f, b, i, s, c = self._formato(logici)
                stile = " ".join(n for n, attivo in (("bold", b), ("italic", i)) if attivo) or "normal"
                colore = COLORI[c][1]
                nome = "v|%s|%d|%s|%s" % (f, s, stile, colore)
                # Il bianco si vede solo su uno sfondo scuro
                self.testo.tag_configure(nome, font=(FONT[f][1], s, stile), foreground="#" + colore,
                                         background="#8a8a8a" if c == "w" else "")
                self.testo.tag_add(nome, indice, indice + "+%dc" % len(valore))
        self.testo.tag_raise("sel")
        self._disegna_numeri()

    def _selezione(self):
        try:
            return self.testo.index("sel.first"), self.testo.index("sel.last")
        except tk.TclError:
            self._messaggio("Seleziona prima il testo da formattare (doppio clic = parola, triplo clic = riga)")
            return None

    def _rimuovi_prefisso(self, prefisso, inizio, fine):
        for tag in self.testo.tag_names():
            if tag.startswith(prefisso):
                self.testo.tag_remove(tag, inizio, fine)

    def _imposta(self, chiave, valore):
        selezione = self._selezione()
        if selezione is None:
            return
        self._rimuovi_prefisso(chiave + ":", *selezione)
        self.testo.tag_add(chiave + ":" + str(valore), *selezione)
        self._dopo_formattazione()

    def _imposta_dimensione(self):
        valore = self.scelta_dimensione.get().strip()
        if not valore.isdigit() or not 1 <= int(valore) <= 99:
            self._messaggio("Dimensione non valida: deve essere tra 1 e 99")
            return
        self._imposta("s", int(valore))

    def _alterna(self, tag):
        selezione = self._selezione()
        if selezione is None:
            return
        inizio, fine = selezione
        indice, tutto = inizio, True
        while self.testo.compare(indice, "<", fine):
            if self.testo.get(indice) != "\n" and tag not in self.testo.tag_names(indice):
                tutto = False
                break
            indice = self.testo.index(indice + "+1c")
        if tutto:
            self.testo.tag_remove(tag, inizio, fine)
        else:
            self.testo.tag_add(tag, inizio, fine)
        self._dopo_formattazione()

    def _predefinito(self):
        selezione = self._selezione()
        if selezione is None:
            return
        for prefisso in ("f:", "s:", "c:", "b", "i"):
            self._rimuovi_prefisso(prefisso, *selezione)
        self._dopo_formattazione()

    def _dopo_formattazione(self):
        self._segna_modificato()
        self._aggiorna_vista()
        self._aggiorna_barra()
        self.testo.focus_set()

    def _tasto(self, evento):
        # Il testo non si modifica: passano solo spostamenti, selezione e copia
        if evento.keysym in TASTI_MOVIMENTO:
            return None
        if evento.state & 0x4 and evento.keysym.lower() in ("c", "a", "g", "b", "i", "home", "end"):
            return None
        return "break"

    def _seleziona_tutto(self):
        self.testo.tag_add("sel", "1.0", "end-1c")
        self._aggiorna_barra()

    def _aggiorna_barra(self):
        selezione = self.testo.tag_ranges("sel")
        indice = selezione[0] if selezione else self.testo.index("insert")
        if not selezione and self.testo.compare("insert", ">", "insert linestart"):
            indice = self.testo.index("insert-1c")
        f, b, i, s, c = self._formato(self.testo.tag_names(indice))
        self.scelta_font.set(FONT[f][0])
        self.scelta_dimensione.set(str(s))
        self.scelta_colore.set(COLORI[c][0])
        self.pulsante_b.configure(relief="sunken" if b else "raised")
        self.pulsante_i.configure(relief="sunken" if i else "raised")
        riga, colonna = self.testo.index("insert").split(".")
        testo = "Riga %s, colonna %s" % (riga, colonna)
        if selezione:
            r1, c1 = str(selezione[0]).split(".")
            r2, c2 = str(selezione[1]).split(".")
            if r1 == r2:
                testo += "   -   selezione: colonne %s-%s" % (c1, c2)
            else:
                testo += "   -   selezione: righe %s-%s" % (r1, r2)
        self._messaggio(testo)

    # --- Lettura e scrittura del modello ------------------------------------

    def _formati_testo(self):
        """Per ogni riga la lista dei formati dei suoi caratteri."""
        formati = []
        for n, riga in enumerate(self.righe):
            formati.append([self._formato(self.testo.tag_names("%d.%d" % (n + 1, c))) for c in range(len(riga))])
        return formati

    def _mostra_testo(self, righe, formati=None):
        self.righe = righe
        self.testo.delete("1.0", "end")
        self.testo.insert("1.0", "\n".join(righe))
        for tag in self.testo.tag_names():
            if tag != "sel":
                self.testo.tag_remove(tag, "1.0", "end")
        if formati:
            for n, formati_riga in enumerate(formati):
                for da, a, formato in modello_gui.pezzi(formati_riga):
                    if formato != PREDEFINITO:
                        for tag in self._tag_da_formato(formato):
                            self.testo.tag_add(tag, "%d.%d" % (n + 1, da), "%d.%d" % (n + 1, a))
        for n, riga in enumerate(righe):
            if len(riga) > COLONNA_MASSIMA:
                self.testo.tag_add("oltre", "%d.%d" % (n + 1, COLONNA_MASSIMA), "%d.end" % (n + 1))
        self.testo.tag_lower("oltre")
        self._aggiorna_vista()
        self.testo.mark_set("insert", "1.0")
        self.testo.focus_set()

    def _leggi_testo(self, percorso):
        # Stessa codifica con cui lo legge il programma (quella predefinita di Windows)
        with open(percorso, "r", errors="replace") as f:
            return f.read().split("\n")

    def _intestazione(self):
        return {
            "testo": modello_gui.SEGNAPOSTO + ".txt",
            "azione": self.azione.get(),
            "uscita": self.uscita.get().strip(),
            "template": self.template.get().strip() if not self.word.get() else "",
            "rotazione": self.rotazione.get() if not self.word.get() else "0",
            "invio": self.invio.get(),
            "destinatario": self.destinatario.get().strip(),
            "coda": self.coda.get("1.0", "end-1c").split("\n") if self.coda.get("1.0", "end-1c") else [],
        }

    def _imposta_intestazione(self, intestazione):
        self.word.set(intestazione["azione"] in modello_gui.AZIONI_WORD)
        self._cambia_tipo(iniziale=True)
        self.azione.set(intestazione["azione"])
        self.template.set(intestazione["template"])
        self.rotazione.set(intestazione["rotazione"])
        self.uscita.set(intestazione["uscita"])
        self.invio.set(intestazione["invio"] if intestazione["invio"] in modello_gui.INVII else "0")
        self.destinatario.set(intestazione["destinatario"])
        self.coda.delete("1.0", "end")
        self.coda.insert("1.0", "\n".join(intestazione["coda"]))
        self.coda.edit_modified(False)

    def _cambia_tipo(self, iniziale=False):
        word = self.word.get()
        self.scelta_azione.configure(values=modello_gui.AZIONI_WORD if word else modello_gui.AZIONI_PDF)
        if self.azione.get() not in self.scelta_azione.cget("values"):
            self.azione.set("docx" if word else "file")
        stato = "disabled" if word else "normal"
        self.scelta_template.configure(state=stato)
        self.scelta_rotazione.configure(state="disabled" if word else "readonly")
        uscita = self.uscita.get()
        if word and uscita.lower().endswith(".pdf"):
            self.uscita.set(uscita[:-4] + ".docx")
        elif not word and uscita.lower().endswith(".docx"):
            self.uscita.set(uscita[:-5] + ".pdf")
        if not iniziale:
            self._segna_modificato()
            self._aggiorna_vista()

    def _elenca_template(self):
        cartella = self.config_servizio.get("cartella_template", "")
        nomi = sorted(f for f in os.listdir(cartella) if f.lower().endswith(".pdf")) if os.path.isdir(
            cartella) else []
        self.scelta_template.configure(values=[""] + nomi)

    def nuovo(self):
        if not self._conferma_perdita():
            return
        percorso = filedialog.askopenfilename(parent=self, title="Testo di esempio (es. una fattura)",
                                              filetypes=[("File di testo", "*.txt"), ("Tutti i file", "*.*")])
        if not percorso:
            return
        self.percorso_testo = percorso
        self.percorso_modello = None
        self._imposta_intestazione(modello_gui.nuovo_modello(self.word.get()))
        self._mostra_testo(self._leggi_testo(percorso))
        self.modificato = False
        self._aggiorna_titolo()

    def apri(self):
        if not self._conferma_perdita():
            return
        os.makedirs(CARTELLA_MODELLI, exist_ok=True)
        percorso = filedialog.askopenfilename(parent=self, title="Apri modello", initialdir=CARTELLA_MODELLI,
                                              filetypes=[("Modelli", "*.gui"), ("Tutti i file", "*.*")])
        if not percorso:
            return
        self.apri_modello(percorso)

    def apri_modello(self, percorso, percorso_testo=None):
        intestazione, linee = modello_gui.leggi_gui(percorso)
        if percorso_testo is None:
            # Il testo di esempio è salvato accanto al modello con lo stesso nome
            percorso_testo = os.path.splitext(percorso)[0] + ".txt"
            if not os.path.isfile(percorso_testo):
                messagebox.showinfo("Testo di esempio", "Scegli un file di testo di esempio per questo modello.",
                                    parent=self)
                percorso_testo = filedialog.askopenfilename(parent=self, title="Testo di esempio",
                                                            filetypes=[("File di testo", "*.txt")])
                if not percorso_testo:
                    return
        righe = self._leggi_testo(percorso_testo)
        self._imposta_intestazione(intestazione)
        self._mostra_testo(righe, modello_gui.leggi_formati(linee, righe, self.word.get()))
        self.percorso_modello = percorso
        self.percorso_testo = percorso_testo
        self.modificato = False
        self._aggiorna_titolo()

    def _linee_formato(self):
        try:
            return modello_gui.linee_formato(self._formati_testo(), self.word.get())
        except modello_gui.ErroreModello as e:
            messagebox.showerror("Modello", str(e), parent=self)
            return None

    def salva(self, con_nome=False):
        if not self.righe:
            self._messaggio("Prima apri un testo di esempio")
            return False
        linee = self._linee_formato()
        if linee is None:
            return False
        percorso = self.percorso_modello
        if con_nome or percorso is None:
            os.makedirs(CARTELLA_MODELLI, exist_ok=True)
            percorso = filedialog.asksaveasfilename(parent=self, title="Salva modello", initialdir=CARTELLA_MODELLI,
                                                    defaultextension=".gui", filetypes=[("Modelli", "*.gui")])
            if not percorso:
                return False
        lavoro = self.config_servizio.get("cartella_lavoro", "")
        if lavoro and os.path.normcase(os.path.abspath(os.path.dirname(percorso))) == \
                os.path.normcase(os.path.abspath(lavoro)):
            messagebox.showerror("Modello", "Non salvare il modello nella cartella di lavoro: il servizio "
                                 "proverebbe a elaborarlo. Usa ad esempio " + CARTELLA_MODELLI, parent=self)
            return False
        modello_gui.scrivi_gui(percorso, self._intestazione(), linee)
        # Copia del testo di esempio accanto al modello, per poterlo riaprire
        copia_testo = os.path.splitext(percorso)[0] + ".txt"
        if self.percorso_testo and os.path.normcase(os.path.abspath(self.percorso_testo)) != \
                os.path.normcase(os.path.abspath(copia_testo)):
            shutil.copyfile(self.percorso_testo, copia_testo)
            self.percorso_testo = copia_testo
        self.percorso_modello = percorso
        self.modificato = False
        self._aggiorna_titolo()
        self._messaggio("Modello salvato: " + percorso)
        return True

    def applica(self):
        if self.modificato or self.percorso_modello is None:
            if not messagebox.askokcancel("Applica", "Prima va salvato il modello. Salvare ora?", parent=self):
                return
            if not self.salva():
                return
        lavoro = self.config_servizio.get("cartella_lavoro", "")
        if not os.path.isdir(lavoro):
            messagebox.showerror("Applica", "Cartella di lavoro non trovata: imposta la configurazione con "
                                 "ConfigManager.py", parent=self)
            return
        file_testo = filedialog.askopenfilenames(parent=self, title="File di testo a cui applicare il modello",
                                                 filetypes=[("File di testo", "*.txt")])
        if not file_testo:
            return
        if not messagebox.askokcancel("Applica", "Verranno creati %d file .gui nella cartella di lavoro\n%s\n"
                                      "e il servizio li elaborerà. Continuare?" % (len(file_testo), lavoro),
                                      parent=self):
            return
        creati = [os.path.basename(modello_gui.applica(self.percorso_modello, f, lavoro)) for f in file_testo]
        messagebox.showinfo("Applica", "Creati:\n" + "\n".join(creati), parent=self)

    # --- Anteprima ---------------------------------------------------------

    def crea_documento(self, cartella):
        """Crea il documento con il codice del servizio; restituisce il percorso."""
        linee = self._linee_formato()
        if linee is None:
            return None
        intestazione = self._intestazione()
        estensione = ".docx" if self.word.get() else ".pdf"
        intestazione.update(testo="anteprima.txt", uscita="anteprima" + estensione,
                            azione="docx" if self.word.get() else "file", invio="0", destinatario="", coda=[])
        if intestazione["template"] and not os.path.isfile(
                os.path.join(self.config_servizio.get("cartella_template", ""), intestazione["template"])):
            messagebox.showerror("Anteprima", "Template non trovato nella cartella dei template: "
                                 + intestazione["template"], parent=self)
            return None
        with open(os.path.join(cartella, "anteprima.txt"), "w", errors="replace") as f:
            f.write("\n".join(self.righe))
        percorso_gui = os.path.join(cartella, "anteprima.gui")
        modello_gui.scrivi_gui(percorso_gui, intestazione, linee)
        with open(percorso_gui, "r") as f:
            guida = f.readlines()
        config = dict(self.config_servizio, cartella_lavoro=cartella, cartella_output=cartella)
        servizio = importlib.import_module("SingleGUY_editor")
        if self.word.get():
            servizio.docx(guida, config)
        else:
            servizio.pdf(guida, config)
        return os.path.join(cartella, "anteprima" + estensione)

    def anteprima(self):
        if not self.righe:
            self._messaggio("Prima apri un testo di esempio")
            return
        cartella = tempfile.mkdtemp(prefix="anteprima_modello_")
        self.configure(cursor="watch")
        self.update_idletasks()
        try:
            percorso = self.crea_documento(cartella)
        except Exception:
            messagebox.showerror("Anteprima", "Errore nel creare il documento:\n\n" + traceback.format_exc(limit=-3),
                                 parent=self)
            return
        finally:
            self.configure(cursor="")
        if percorso is None:
            return
        if self.word.get():
            os.startfile(percorso)
        else:
            FinestraAnteprima(self, percorso)

    # --- Aiuto ---------------------------------------------------------------

    def apri_guida(self):
        cartelle = [CARTELLA_PROGRAMMA, getattr(sys, "_MEIPASS", CARTELLA_PROGRAMMA)]
        for cartella in cartelle:
            percorso = os.path.join(cartella, GUIDA)
            if os.path.isfile(percorso):
                os.startfile(percorso)
                return
        messagebox.showerror("Guida", "Guida non trovata: " + os.path.join(CARTELLA_PROGRAMMA, GUIDA), parent=self)

    def _informazioni(self):
        messagebox.showinfo("Informazioni", "Editor modelli di InfasService\n\nPrepara i modelli .gui con la "
                            "formattazione di fatture, bolle e altri documenti, da applicare ai file di testo "
                            "elaborati dal servizio.\n\nGuida completa: Aiuto > Guida all'uso (F1)", parent=self)

    # --- Stato, titolo, chiusura -------------------------------------------

    def _segna_modificato(self):
        if self.righe:
            self.modificato = True
            self._aggiorna_titolo()

    def _coda_modificata(self, evento):
        if self.coda.edit_modified():
            self._segna_modificato()
            self.coda.edit_modified(False)

    def _aggiorna_titolo(self):
        nome = os.path.basename(self.percorso_modello) if self.percorso_modello else "nuovo modello"
        self.title("Editor modelli - " + nome + (" *" if self.modificato else ""))

    def _messaggio(self, testo):
        self.stato.configure(text=testo)

    def _conferma_perdita(self):
        if not self.modificato:
            return True
        risposta = messagebox.askyesnocancel("Modello", "Salvare le modifiche al modello?", parent=self)
        if risposta is None:
            return False
        return self.salva() if risposta else True

    def _esci(self):
        if self._conferma_perdita():
            self.destroy()


class FinestraAnteprima(tk.Toplevel):
    def __init__(self, genitore, percorso):
        super().__init__(genitore)
        import fitz
        self.title("Anteprima - " + os.path.basename(genitore.percorso_modello or "nuovo modello"))
        self.geometry("820x900")
        self.documento = fitz.open(percorso)
        self.percorso = percorso
        self.pagina = 0
        self.zoom = tk.StringVar(value="100%")

        barra = ttk.Frame(self, padding=6)
        barra.pack(side="top", fill="x")
        ttk.Button(barra, text="<", width=3, command=lambda: self._vai(-1)).pack(side="left")
        self.etichetta = ttk.Label(barra, width=14, anchor="center")
        self.etichetta.pack(side="left")
        ttk.Button(barra, text=">", width=3, command=lambda: self._vai(1)).pack(side="left")
        zoom = ttk.Combobox(barra, textvariable=self.zoom, values=["50%", "75%", "100%", "125%", "150%"], width=6,
                            state="readonly")
        zoom.pack(side="left", padx=10)
        zoom.bind("<<ComboboxSelected>>", lambda e: self._mostra())
        ttk.Button(barra, text="Apri nel visualizzatore", command=lambda: os.startfile(self.percorso)).pack(
            side="right")

        cornice = ttk.Frame(self)
        cornice.pack(fill="both", expand=True)
        self.tela = tk.Canvas(cornice, background="#808080", highlightthickness=0)
        verticale = ttk.Scrollbar(cornice, orient="vertical", command=self.tela.yview)
        orizzontale = ttk.Scrollbar(self, orient="horizontal", command=self.tela.xview)
        verticale.pack(side="right", fill="y")
        orizzontale.pack(side="bottom", fill="x")
        self.tela.pack(side="left", fill="both", expand=True)
        self.tela.configure(yscrollcommand=verticale.set, xscrollcommand=orizzontale.set)
        self.tela.bind("<MouseWheel>", lambda e: self.tela.yview_scroll(-e.delta // 120, "units"))
        self.protocol("WM_DELETE_WINDOW", self._chiudi)
        self._mostra()

    def _vai(self, passo):
        self.pagina = max(0, min(self.documento.page_count - 1, self.pagina + passo))
        self._mostra()

    def _mostra(self):
        dpi = int(96 * int(self.zoom.get().rstrip("%")) / 100)
        pix = self.documento[self.pagina].get_pixmap(dpi=dpi, alpha=False)
        self.immagine = tk.PhotoImage(data=base64.b64encode(pix.tobytes("png")).decode("ascii"), format="png")
        self.tela.delete("all")
        self.tela.create_image(10, 10, anchor="nw", image=self.immagine)
        self.tela.configure(scrollregion=(0, 0, pix.width + 20, pix.height + 20))
        self.etichetta.configure(text="Pagina %d di %d" % (self.pagina + 1, self.documento.page_count))

    def _chiudi(self):
        self.documento.close()
        self.destroy()


if __name__ == "__main__":
    app = EditorModelli()
    if len(sys.argv) > 1:
        app.apri_modello(sys.argv[1])
    app.mainloop()
