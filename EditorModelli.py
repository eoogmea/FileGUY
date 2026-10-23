"""Editor dei modelli .gui per fatture, bolle e altri documenti ripetitivi.

Avvio: python EditorModelli.py

1. File > Nuovo da testo di esempio: si apre un file di testo come quelli che arrivano al servizio.
2. Si seleziona una riga intera, una parte o anche solo alcune lettere e si sceglie carattere,
   dimensione, grassetto, corsivo e colore. Il testo si modifica solo con la spunta "Modifica testo":
   le modifiche vanno nella copia salvata accanto al modello, il file scelto resta com'era.
3. Si scelgono tipo di documento, template PDF, rotazione ed eventuale invio.
4. Anteprima (F5): crea il documento con lo stesso codice del servizio.
   Filigrana (F6): il documento sul template accanto all'editor, aggiornato a ogni modifica;
   un clic sul documento porta il cursore alla riga e colonna corrispondenti.
5. File > Salva modello: crea un .gui con {nome} al posto del nome del file, più una copia
   del testo di esempio accanto (serve solo per riaprire il modello).
6. File > Applica a file di testo: per ogni .txt scelto crea il .gui nella cartella di lavoro,
   dove il servizio lo elabora.
Aiuti alla formattazione (producono gli stessi codici del .gui): annulla/ripeti, copia e incolla
formato, selezione a blocco (Alt+trascina), copia del formato di una riga su altre righe,
righello delle colonne, modelli recenti e zoom.
File > Barcode su un PDF: crea i .gui "barcode" (barcode da mettere su un PDF che c'è già),
disegnandone la posizione sulla pagina.
Il campo Documento creato vuoto = il modello non crea documenti e fa solo l'invio.
File > Nuovo modello solo invio: un modello senza testo di esempio, che non crea documenti e manda solo
email, Telegram o WhatsApp (con file già presenti sul computer); il servizio lo elabora anche senza .txt.
La guida completa è Guida_EditorModelli.html (menu Aiuto, F1).
"""
import base64
import importlib
import json
import os
import shutil
import sys
import tempfile
import tkinter as tk
import traceback
from tkinter import filedialog, messagebox, ttk

import modello_gui
from config_sicuro import carica_config
from piattaforma import (SCALA_FINESTRE, apri_file, avvia_finestra, cartella_utente, console_sicura, leggi_testo,
                         registra_font, scala_finestra)

# L'exe dell'editor non ha la console: i messaggi scritti su stdout/stderr (es. gli avvisi di
# PyPDF2 nell'anteprima con un template) farebbero fallire l'operazione
console_sicura()
from modello_gui import COLONNA_MASSIMA, COLORI, FONT, PREDEFINITO

if getattr(sys, "frozen", False):
    CARTELLA_PROGRAMMA = os.path.dirname(sys.executable)
else:
    CARTELLA_PROGRAMMA = os.path.dirname(os.path.abspath(__file__))
CARTELLA_MODELLI = os.path.join(CARTELLA_PROGRAMMA, "ModelliGUI")
# I font della cartella Fonts (Helvetica, DejaVu Sans Mono, IBM Plex Mono...) anche nell'editor,
# prima di creare le finestre (vale anche per EditorExcel, che importa questo modulo)
registra_font(os.path.join(CARTELLA_PROGRAMMA, "Fonts"))
# La guida sta accanto al programma; nell'exe creato con PyInstaller anche dentro l'exe
GUIDA = "Guida_EditorModelli.html"

NOMI_FONT = {nome: lettera for lettera, (nome, _) in FONT.items()}
DIMENSIONI = ["6", "7", "8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "28", "36"]
NOMI_COLORI = {nome: lettera for lettera, (nome, _) in COLORI.items()}
TASTI_MOVIMENTO = {"Left", "Right", "Up", "Down", "Home", "End", "Prior", "Next", "Shift_L", "Shift_R",
                   "Control_L", "Control_R", "Alt_L", "Alt_R", "Escape", "Tab"}
# Con Ctrl passano anche questi tasti (copia, seleziona tutto, grassetto, corsivo, zoom)
TASTI_CONTROL = {"c", "a", "g", "b", "i", "home", "end", "plus", "equal", "minus", "0", "kp_add",
                 "kp_subtract", "kp_0"}

# Modelli recenti e zoom, per utente (su Linux accanto alla chiave di config.dat)
FILE_PREFERENZE = os.path.join(cartella_utente(), "editor_modelli.json")
MAX_RECENTI = 8
LIVELLI_ZOOM = [0.75, 0.9, 1.0, 1.1, 1.25, 1.5, 1.75, 2.0, 2.5]
# Passi di formattazione che si possono annullare
MAX_ANNULLA = 100


def apri_con_avviso(percorso, genitore):
    """Apre il file con il programma predefinito; se non si può (es. Linux senza xdg-open) lo dice."""
    try:
        apri_file(percorso)
    except (OSError, RuntimeError) as e:
        messagebox.showerror("Apri", "Impossibile aprire " + percorso + ":\n" + str(e), parent=genitore)


def proposta_salvataggio(percorso_modello, percorso_testo, cartella_modelli):
    """Cartella e nome proposti da "Salva con nome": quelli del modello aperto, altrimenti il nome
    del testo di esempio con .gui nella cartella dei modelli."""
    if percorso_modello:
        return os.path.dirname(percorso_modello), os.path.basename(percorso_modello)
    nome = os.path.splitext(os.path.basename(percorso_testo))[0] + ".gui" if percorso_testo else ""
    return cartella_modelli, nome


def dimensiona(finestra, larghezza, altezza, minima=None):
    """Dimensione iniziale della finestra: quella indicata, ma al massimo l'85% dello schermo
    (sui portatili e con lo zoom di Windows lo schermo è piccolo). La finestra va al centro,
    un po' in alto per lasciare posto alla barra delle applicazioni. Su Linux le misure si riducono
    come i caratteri (SCALA_FINESTRE)."""
    schermo_x, schermo_y = finestra.winfo_screenwidth(), finestra.winfo_screenheight()
    larghezza, altezza = int(larghezza * SCALA_FINESTRE), int(altezza * SCALA_FINESTRE)
    if minima:
        minima = (int(minima[0] * SCALA_FINESTRE), int(minima[1] * SCALA_FINESTRE))
    larghezza = min(larghezza, int(schermo_x * 0.85))
    altezza = min(altezza, int(schermo_y * 0.85))
    finestra.geometry("%dx%d+%d+%d" % (larghezza, altezza, max(0, (schermo_x - larghezza) // 2),
                                       max(0, (schermo_y - altezza) // 3)))
    if minima:
        finestra.minsize(min(minima[0], larghezza), min(minima[1], altezza))


# Al posto del testo di esempio in un modello di solo invio
TESTO_SOLO_INVIO = ("Modello di solo invio: non crea documenti e non ha un testo di esempio.\n\n"
                    "Scegli in alto il canale (email, Telegram o WhatsApp), il destinatario e cosa mandare: "
                    "testo, allegati o file che ci sono già sul computer (con il percorso completo).\n\n"
                    "Salvato, il .gui funziona anche da solo nella cartella di lavoro, senza il .txt: "
                    "il servizio fa l'invio e lo cancella.")


# Se un font non è installato (es. Consolas o Calibri su Linux) l'editor mostra quello che il servizio usa al
# suo posto nel PDF (SingleGUY.FONT_SOSTITUTI), così le colonne si vedono come nel documento
FAMIGLIE_SOSTITUTE = {"o": "DejaVu Sans Mono", "h": "Helvetica", "k": "Helvetica", "v": "Helvetica",
                      "u": "Helvetica", "g": "Times New Roman"}
_famiglie_disponibili = None


def famiglia_tk(lettera):
    """Il carattere con cui l'editor mostra il font della lettera del .gui (serve una finestra già creata)."""
    global _famiglie_disponibili
    if _famiglie_disponibili is None:
        import tkinter.font as tkfont
        _famiglie_disponibili = set(tkfont.families())
    famiglia = FONT[lettera][1]
    if famiglia in _famiglie_disponibili:
        return famiglia
    return FAMIGLIE_SOSTITUTE.get(lettera, famiglia)


def leggi_preferenze():
    """Preferenze dell'utente (modelli recenti, zoom), comuni a EditorModelli ed EditorExcel."""
    try:
        with open(FILE_PREFERENZE, encoding="utf-8") as f:
            preferenze = json.load(f)
        return preferenze if isinstance(preferenze, dict) else {}
    except (OSError, ValueError):
        return {}


def scrivi_preferenze(preferenze):
    # Non sono essenziali: se non si possono salvare si va avanti
    try:
        os.makedirs(os.path.dirname(FILE_PREFERENZE), exist_ok=True)
        with open(FILE_PREFERENZE, "w", encoding="utf-8") as f:
            json.dump(preferenze, f, indent=1)
    except OSError:
        pass


def apri_guida(genitore):
    # La guida sta accanto al programma o dentro l'exe creato con PyInstaller
    for cartella in (CARTELLA_PROGRAMMA, getattr(sys, "_MEIPASS", CARTELLA_PROGRAMMA)):
        percorso = os.path.join(cartella, GUIDA)
        if os.path.isfile(percorso):
            apri_con_avviso(percorso, genitore)
            return
    messagebox.showerror("Guida", "Guida non trovata: " + os.path.join(CARTELLA_PROGRAMMA, GUIDA), parent=genitore)


class SceltaDescritta(ttk.Combobox):
    """Tendina che mostra le descrizioni ma tiene nella variabile il codice del .gui
    (es. mostra "Crea e stampa", la variabile vale "stampa")."""

    def __init__(self, genitore, variabile, descrizioni, **opzioni):
        self._vista = tk.StringVar()
        super().__init__(genitore, textvariable=self._vista, state="readonly", **opzioni)
        self.variabile = variabile
        self.imposta_scelte(descrizioni)
        self.bind("<<ComboboxSelected>>", lambda e: self.variabile.set(self._codici[self._vista.get()]))
        # Se il codice cambia da programma (modello aperto, tipo cambiato) si aggiorna la descrizione
        variabile.trace_add("write", lambda *a: self._mostra())

    def imposta_scelte(self, descrizioni):
        """descrizioni: {codice: descrizione}, nell'ordine in cui vanno mostrate."""
        self._descrizioni = dict(descrizioni)
        self._codici = {descrizione: codice for codice, descrizione in descrizioni.items()}
        self.configure(values=list(self._descrizioni.values()))
        self._mostra()

    def codici(self):
        return list(self._descrizioni)

    def _mostra(self):
        codice = self.variabile.get()
        self._vista.set(self._descrizioni.get(codice, codice))


def tag_logico(tag):
    """I tag che descrivono il formato (gli altri servono solo a mostrarlo: v|..., sel, oltre)."""
    return tag[:2] in ("f:", "s:", "c:") or tag in ("b", "i")


class DialogoNumeri(tk.Toplevel):
    """Finestra con alcuni campi numerici: risultato = lista dei valori, None se annullata."""

    def __init__(self, genitore, titolo, spiegazione, campi):
        super().__init__(genitore)
        self.title(titolo)
        self.transient(genitore)
        self.resizable(False, False)
        self.risultato = None
        corpo = ttk.Frame(self, padding=12)
        corpo.pack(fill="both", expand=True)
        ttk.Label(corpo, text=spiegazione, wraplength=340, justify="left").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        self.campi = []
        for n, (etichetta, valore, minimo, massimo) in enumerate(campi):
            ttk.Label(corpo, text=etichetta).grid(row=n + 1, column=0, sticky="w", padx=(0, 8), pady=2)
            variabile = tk.StringVar(value=str(valore))
            voce = ttk.Spinbox(corpo, from_=minimo, to=massimo, textvariable=variabile, width=7)
            voce.grid(row=n + 1, column=1, sticky="w", pady=2)
            if n == 0:
                voce.focus_set()
                voce.selection_range(0, "end")
            self.campi.append((etichetta, variabile, minimo, massimo))
        pulsanti = ttk.Frame(corpo)
        pulsanti.grid(row=len(campi) + 1, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(pulsanti, text="Annulla", command=self.destroy).pack(side="right")
        ttk.Button(pulsanti, text="OK", command=self._ok).pack(side="right", padx=(0, 6))
        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self.destroy())
        self.grab_set()

    def _ok(self):
        valori = []
        for etichetta, variabile, minimo, massimo in self.campi:
            testo = variabile.get().strip()
            if not testo.isdigit() or not minimo <= int(testo) <= massimo:
                messagebox.showerror(self.title(), "%s: scrivi un numero da %d a %d" % (etichetta, minimo, massimo),
                                     parent=self)
                return
            valori.append(int(testo))
        self.risultato = valori
        self.destroy()


class DialogoSegnaposto(tk.Toplevel):
    """Opzioni del testo con le x, con l'anteprima delle prime righe: risultato = (x in più, simboli) o None."""

    def __init__(self, genitore, righe):
        super().__init__(genitore)
        self.title("Testo con segnaposto x")
        self.transient(genitore)
        self.resizable(False, False)
        self.righe = righe
        self.risultato = None
        self.in_piu = tk.StringVar(value="3")
        self.simboli = tk.BooleanVar(value=True)
        corpo = ttk.Frame(self, padding=12)
        corpo.pack(fill="both", expand=True)
        ttk.Label(corpo, text="Ogni carattere del testo diventa una x, così il modello vale per tutti i documenti "
                              "e si vede subito dove andrà ogni dato. Il file di partenza non cambia: il testo con "
                              "le x va nella copia salvata accanto al modello.", wraplength=520,
                  justify="left").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))
        ttk.Label(corpo, text="x in più dopo ogni campo").grid(row=1, column=0, sticky="w")
        voce = ttk.Spinbox(corpo, from_=0, to=30, textvariable=self.in_piu, width=5, command=self._anteprima)
        voce.grid(row=1, column=1, sticky="w", padx=(6, 0))
        voce.bind("<KeyRelease>", lambda e: self._anteprima())
        ttk.Label(corpo, text="per fare spazio a nomi e importi più lunghi (le colonne dopo non si spostano)",
                  foreground="gray").grid(row=1, column=2, sticky="w", padx=(6, 0))
        ttk.Checkbutton(corpo, text="Lascia punteggiatura e simboli (/ . , : - €): date, importi e righe di "
                                    "separazione restano riconoscibili", variable=self.simboli,
                        command=self._anteprima).grid(row=2, column=0, columnspan=3, sticky="w", pady=(6, 6))
        self.esempio = tk.Text(corpo, height=12, width=86, wrap="none", font=(famiglia_tk("o"), 9), background="#f6f6f6")
        self.esempio.grid(row=3, column=0, columnspan=3, sticky="we")
        pulsanti = ttk.Frame(corpo)
        pulsanti.grid(row=4, column=0, columnspan=3, sticky="e", pady=(10, 0))
        ttk.Button(pulsanti, text="Annulla", command=self.destroy).pack(side="right")
        ttk.Button(pulsanti, text="Sostituisci", command=self._ok).pack(side="right", padx=(0, 6))
        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self.destroy())
        self._anteprima()
        self.grab_set()
        voce.focus_set()

    def _valori(self):
        testo = self.in_piu.get().strip()
        if not testo.isdigit() or int(testo) > 30:
            return None
        return int(testo), self.simboli.get()

    def _anteprima(self):
        valori = self._valori()
        self.esempio.configure(state="normal")
        self.esempio.delete("1.0", "end")
        if valori is None:
            self.esempio.insert("1.0", "Scrivi un numero da 0 a 30")
        else:
            # Le prime righe con del testo, prima e dopo
            prime = [r for r in self.righe if r.strip()][:6]
            dopo = modello_gui.testo_segnaposto(prime, *valori)
            self.esempio.insert("1.0", "\n".join(prime) + "\n\ndiventa:\n\n" + "\n".join(dopo))
        self.esempio.configure(state="disabled")

    def _ok(self):
        valori = self._valori()
        if valori is None:
            messagebox.showerror(self.title(), "x in più: scrivi un numero da 0 a 30", parent=self)
            return
        self.risultato = valori
        self.destroy()


def chiedi_numeri(genitore, titolo, spiegazione, campi):
    finestra = DialogoNumeri(genitore, titolo, spiegazione, campi)
    genitore.wait_window(finestra)
    return finestra.risultato


class EditorModelli(tk.Tk):
    def __init__(self):
        super().__init__()
        scala_finestra(self)
        dimensiona(self, 1060, 680, minima=(820, 460))
        try:
            self.config_servizio = carica_config()
        except Exception:
            self.config_servizio = {"cartella_lavoro": "", "cartella_template": "", "cartella_output": ""}
        self.percorso_modello = None
        self.percorso_testo = None
        self.righe = []
        # Modello di solo invio: niente testo di esempio e niente documento
        self.solo_invio = False
        self.modificato = False
        # Il testo di esempio è stato cambiato nell'editor (va scritto al salvataggio)
        self.testo_cambiato = False
        self.immagini_anteprima = []
        self.modifica_testo = tk.BooleanVar(value=False)
        # Annulla/ripeti della formattazione: istantanee dei tag di formato
        self.da_annullare = []
        self.da_ripetere = []
        # Formato preso con "Copia formato" (font, grassetto, corsivo, dimensione, colore)
        self.formato_copiato = None
        # Inizio della selezione a blocco con Alt+trascina
        self.ancora_blocco = None
        # Finestra della filigrana (il template con il testo, aggiornato a ogni modifica)
        self.finestra_filigrana = None
        self.attesa_filigrana = None
        self.preferenze = self._leggi_preferenze()
        self.zoom = self.preferenze.get("zoom", 1.0) if self.preferenze.get("zoom") in LIVELLI_ZOOM else 1.0

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
        file.add_command(label="Nuovo da testo con segnaposto x...", command=self.nuovo_segnaposto)
        file.add_command(label="Nuovo modello solo invio (senza testo)...", command=self.nuovo_solo_invio)
        file.add_command(label="Apri modello...", command=self.apri, accelerator="Ctrl+O")
        self.menu_recenti = tk.Menu(file, tearoff=False, postcommand=self._elenca_recenti)
        file.add_cascade(label="Modelli recenti", menu=self.menu_recenti)
        file.add_separator()
        file.add_command(label="Salva modello", command=self.salva, accelerator="Ctrl+S")
        file.add_command(label="Salva modello con nome...", command=lambda: self.salva(con_nome=True))
        file.add_separator()
        file.add_command(label="Applica a file di testo...", command=self.applica)
        file.add_separator()
        file.add_command(label="Barcode su un PDF...", command=lambda: FinestraBarcode(self))
        file.add_separator()
        file.add_command(label="Esci", command=self._esci)
        menu.add_cascade(label="File", menu=file)
        modifica = tk.Menu(menu, tearoff=False)
        modifica.add_command(label="Annulla", command=lambda: self._annulla_o_ripeti(False), accelerator="Ctrl+Z")
        modifica.add_command(label="Ripeti", command=lambda: self._annulla_o_ripeti(True), accelerator="Ctrl+Y")
        modifica.add_separator()
        modifica.add_command(label="Copia formato", command=self.copia_formato, accelerator="Ctrl+Maiusc+C")
        modifica.add_command(label="Incolla formato", command=self.incolla_formato, accelerator="Ctrl+Maiusc+V")
        modifica.add_separator()
        modifica.add_command(label="Selezione a blocco...", command=self.dialogo_blocco, accelerator="Alt+trascina")
        modifica.add_command(label="Copia il formato di una riga su altre righe...", command=self.dialogo_copia_riga)
        modifica.add_separator()
        modifica.add_command(label="Sostituisci il testo con segnaposto x...", command=self.segnaposto)
        menu.add_cascade(label="Modifica", menu=modifica)
        vista = tk.Menu(menu, tearoff=False)
        vista.add_command(label="Ingrandisci", command=lambda: self.cambia_zoom(1), accelerator="Ctrl++")
        vista.add_command(label="Riduci", command=lambda: self.cambia_zoom(-1), accelerator="Ctrl+-")
        vista.add_command(label="Dimensione normale", command=lambda: self.cambia_zoom(0), accelerator="Ctrl+0")
        menu.add_cascade(label="Visualizza", menu=vista)
        menu.add_command(label="Anteprima (F5)", command=self.anteprima)
        menu.add_command(label="Filigrana (F6)", command=self.filigrana)
        aiuto = tk.Menu(menu, tearoff=False)
        aiuto.add_command(label="Guida all'uso", command=self.apri_guida, accelerator="F1")
        aiuto.add_command(label="Informazioni", command=self._informazioni)
        menu.add_cascade(label="Aiuto", menu=aiuto)
        self.configure(menu=menu)
        self.bind("<Control-n>", lambda e: self.nuovo())
        self.bind("<Control-o>", lambda e: self.apri())
        self.bind("<Control-s>", lambda e: self.salva())
        self.bind("<F5>", lambda e: self.anteprima())
        self.bind("<F6>", lambda e: self.filigrana())
        self.bind("<F1>", lambda e: self.apri_guida())
        for tasto in ("<Control-plus>", "<Control-equal>", "<Control-KP_Add>"):
            self.bind(tasto, lambda e: self.cambia_zoom(1))
        for tasto in ("<Control-minus>", "<Control-KP_Subtract>"):
            self.bind(tasto, lambda e: self.cambia_zoom(-1))
        for tasto in ("<Control-Key-0>", "<Control-KP_0>"):
            self.bind(tasto, lambda e: self.cambia_zoom(0))

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
        self.scelta_azione = SceltaDescritta(riquadro, self.azione, {}, width=22)
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
        self.campo_uscita = ttk.Entry(riquadro, textvariable=self.uscita, width=24)
        self.campo_uscita.grid(row=1, column=1, columnspan=3, sticky="we", pady=(4, 0))
        ttk.Label(riquadro, text="Invio").grid(row=1, column=4, sticky="w", padx=(16, 4), pady=(4, 0))
        SceltaDescritta(riquadro, self.invio, modello_gui.DESCRIZIONI_INVII, width=12).grid(row=1, column=5, sticky="w", pady=(4, 0))
        self.etichetta_destinatario = ttk.Label(riquadro, text="Destinatario")
        self.etichetta_destinatario.grid(row=1, column=6, sticky="w", padx=(16, 4), pady=(4, 0))
        self.campo_destinatario = ttk.Entry(riquadro, textvariable=self.destinatario, width=28)
        self.campo_destinatario.grid(row=1, column=7, columnspan=2, sticky="we", pady=(4, 0))
        # Sotto: i dati dell'invio, diversi per email, Telegram e WhatsApp
        self.pannello = PannelloInvio(riquadro, self)
        self.pannello.grid(row=2, column=0, columnspan=9, sticky="we")
        self.invio.trace_add("write", lambda *a: self._invio_cambiato())
        self._invio_cambiato()
        riquadro.columnconfigure(8, weight=1)

    def _invio_cambiato(self):
        invio = self.invio.get()
        self.etichetta_destinatario.configure(text=ETICHETTE_DESTINATARIO.get(invio, "Destinatario"))
        self.campo_destinatario.configure(state="disabled" if invio == "0" else "normal")
        self.pannello.ricostruisci()

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
        ttk.Button(barra, text="Filigrana", command=self.filigrana).pack(side="right", padx=(0, 4))
        ttk.Checkbutton(barra, text="Modifica testo", variable=self.modifica_testo,
                        command=self._cambia_modifica).pack(side="right", padx=(0, 10))

    def _crea_testo(self):
        # Righello delle colonne della riga del cursore, allineato al testo (a sinistra lo spazio dei numeri)
        cornice_righello = ttk.Frame(self)
        cornice_righello.pack(side="top", fill="x", padx=6)
        tk.Frame(cornice_righello, width=44, height=1).pack(side="left")
        self.righello = tk.Canvas(cornice_righello, height=18, background="#f0f0f0", highlightthickness=0)
        self.righello.pack(side="left", fill="x", expand=True)
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
        # undo=True serve solo quando il testo si modifica (i formati non entrano nell'annulla)
        self.testo = tk.Text(cornice, wrap="none", undo=True, padx=8, pady=8, spacing1=2, spacing3=2, height=10,
                             exportselection=False, inactiveselectbackground="#9cc3ee",
                             font=(famiglia_tk(PREDEFINITO[0]), self._dimensione_vista(PREDEFINITO[3])),
                             xscrollcommand=self._scorrimento_orizzontale, yscrollcommand=self._scorrimento,
                             insertwidth=2)
        self.testo.pack(side="left", fill="both", expand=True)
        self.sfondo_testo = self.testo.cget("background")
        verticale.configure(command=self._scorri)
        orizzontale.configure(command=self.testo.xview)
        self.scorrimento_orizzontale = orizzontale
        # I caratteri oltre la colonna 99 non possono avere un formato diverso dal pezzo prima
        self.testo.tag_configure("oltre", background="#f4f4f4")

        self.testo.bind("<Key>", self._tasto)
        for evento in ("<<Paste>>", "<<Cut>>", "<<Clear>>", "<<Undo>>", "<<Redo>>"):
            self.testo.bind(evento, lambda e: None if self.modifica_testo.get() else "break")
        self.testo.bind("<<Modified>>", self._testo_modificato)
        self.testo.bind("<Control-a>", lambda e: self._seleziona_tutto() or "break")
        for tasto in ("<Control-g>", "<Control-G>", "<Control-b>", "<Control-B>"):
            self.testo.bind(tasto, lambda e: self._alterna("b") or "break")
        for tasto in ("<Control-i>", "<Control-I>"):
            self.testo.bind(tasto, lambda e: self._alterna("i") or "break")
        for evento in ("<KeyRelease>", "<ButtonRelease-1>", "<B1-Motion>"):
            self.testo.bind(evento, lambda e: self._aggiorna_barra(), add="+")
        self.testo.bind("<Configure>", lambda e: (self._disegna_numeri(), self._disegna_righello()))

        # Annulla/ripeti: il testo se "Modifica testo" è attiva, altrimenti la formattazione
        for tasto in ("<Control-z>", "<Control-Z>"):
            self.testo.bind(tasto, lambda e: self._annulla_o_ripeti(False) or "break")
        for tasto in ("<Control-y>", "<Control-Y>", "<Control-Shift-Z>"):
            self.testo.bind(tasto, lambda e: self._annulla_o_ripeti(True) or "break")
        self.testo.bind("<Control-Shift-C>", lambda e: self.copia_formato() or "break")
        self.testo.bind("<Control-Shift-V>", lambda e: self.incolla_formato() or "break")
        # Selezione a blocco: le stesse colonne su più righe
        self.testo.bind("<Alt-Button-1>", self._blocco_inizio)
        self.testo.bind("<Alt-B1-Motion>", self._blocco_trascina)
        self.testo.bind("<Alt-ButtonRelease-1>", lambda e: self._aggiorna_barra() or "break")
        # Dopo Alt+trascina, rilasciando Alt Windows attiverebbe il menu
        for tasto in ("<KeyRelease-Alt_L>", "<KeyRelease-Alt_R>"):
            self.testo.bind(tasto, self._rilascio_alt)
        # Ctrl + rotellina: zoom (su Linux la rotellina è Button-4/5)
        self.testo.bind("<Control-MouseWheel>", lambda e: self.cambia_zoom(1 if e.delta > 0 else -1) or "break")
        self.testo.bind("<Control-Button-4>", lambda e: self.cambia_zoom(1) or "break")
        self.testo.bind("<Control-Button-5>", lambda e: self.cambia_zoom(-1) or "break")

    # --- Numeri di riga ----------------------------------------------------

    def _scorrimento(self, inizio, fine):
        self.scorrimento_verticale.set(inizio, fine)
        self._disegna_numeri()
        self._disegna_righello()

    def _scorrimento_orizzontale(self, inizio, fine):
        self.scorrimento_orizzontale.set(inizio, fine)
        self._disegna_righello()

    def _scorri(self, *argomenti):
        self.testo.yview(*argomenti)
        self._disegna_numeri()

    def _disegna_righello(self):
        """Colonne della riga del cursore (se non si vede, della prima riga visibile).

        Con i caratteri proporzionali le colonne non sono allineate tra le righe: il righello
        segue i caratteri della riga, così si vede subito in che colonna cade un pezzo di testo.
        """
        if not hasattr(self, "righello"):
            return
        self.righello.delete("all")
        if not self.righe:
            return
        riga = int(self.testo.index("insert").split(".")[0])
        if self.testo.dlineinfo("%d.0" % riga) is None:
            riga = int(self.testo.index("@0,0").split(".")[0])
        if riga > len(self.righe):
            return
        lunghezza = len(self.righe[riga - 1])
        selezionate = set()
        intervalli = self.testo.tag_ranges("sel")
        for n in range(0, len(intervalli), 2):
            r1, c1 = (int(x) for x in str(intervalli[n]).split("."))
            r2, c2 = (int(x) for x in str(intervalli[n + 1]).split("."))
            if r1 <= riga <= r2:
                selezionate.update(range(c1 if r1 == riga else 0, c2 if r2 == riga else lunghezza))
        altezza = int(self.righello.cget("height"))
        for colonna in range(lunghezza + 1):
            box = self.testo.bbox("%d.%d" % (riga, colonna))
            if box is None:
                continue
            x = box[0]
            if colonna in selezionate:
                self.righello.create_rectangle(x, 0, x + box[2], altezza, fill="#cfe3f7", width=0)
            if colonna == COLONNA_MASSIMA:
                # Oltre la colonna 99 il formato non può cambiare
                self.righello.create_line(x, 0, x, altezza, fill="red")
                self.righello.create_text(x + 2, 1, anchor="nw", text=str(colonna), fill="red", font=(famiglia_tk("o"), 7))
            elif colonna % 10 == 0:
                self.righello.create_line(x, altezza - 7, x, altezza, fill="gray")
                self.righello.create_text(x + 1, 1, anchor="nw", text=str(colonna), fill="gray", font=(famiglia_tk("o"), 7))
            elif colonna % 5 == 0:
                self.righello.create_line(x, altezza - 4, x, altezza, fill="gray")
            else:
                self.righello.create_line(x, altezza - 2, x, altezza, fill="#b0b0b0")

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
                                    fill="gray", font=(famiglia_tk("o"), 9))
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
                # Lo zoom cambia solo la dimensione mostrata, non quella salvata nel .gui
                dimensione = self._dimensione_vista(s)
                nome = "v|%s|%d|%s|%s" % (f, dimensione, stile, colore)
                # Il bianco si vede solo su uno sfondo scuro
                self.testo.tag_configure(nome, font=(famiglia_tk(f), dimensione, stile), foreground="#" + colore,
                                         background="#8a8a8a" if c == "w" else "")
                self.testo.tag_add(nome, indice, indice + "+%dc" % len(valore))
        self.testo.tag_raise("sel")
        self._disegna_numeri()
        self._disegna_righello()

    def _selezione(self):
        """Pezzi selezionati [(inizio, fine), ...]: uno solo, o uno per riga con la selezione a blocco."""
        intervalli = self.testo.tag_ranges("sel")
        if not intervalli:
            self._messaggio("Seleziona prima il testo da formattare (doppio clic = parola, triplo clic = riga, "
                            "Alt+trascina = blocco di colonne)")
            return None
        return [(str(intervalli[n]), str(intervalli[n + 1])) for n in range(0, len(intervalli), 2)]

    def _rimuovi_prefisso(self, prefisso, inizio, fine):
        for tag in self.testo.tag_names():
            if tag.startswith(prefisso):
                self.testo.tag_remove(tag, inizio, fine)

    def _imposta(self, chiave, valore):
        selezione = self._selezione()
        if selezione is None:
            return
        self._prima_di_formattare()
        for inizio, fine in selezione:
            self._rimuovi_prefisso(chiave + ":", inizio, fine)
            self.testo.tag_add(chiave + ":" + str(valore), inizio, fine)
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
        # Se tutto il selezionato ha già il grassetto (o corsivo) si toglie, altrimenti si mette
        tutto = True
        for inizio, fine in selezione:
            indice = inizio
            while tutto and self.testo.compare(indice, "<", fine):
                if self.testo.get(indice) != "\n" and tag not in self.testo.tag_names(indice):
                    tutto = False
                indice = self.testo.index(indice + "+1c")
        self._prima_di_formattare()
        for inizio, fine in selezione:
            if tutto:
                self.testo.tag_remove(tag, inizio, fine)
            else:
                self.testo.tag_add(tag, inizio, fine)
        self._dopo_formattazione()

    def _predefinito(self):
        selezione = self._selezione()
        if selezione is None:
            return
        self._prima_di_formattare()
        for inizio, fine in selezione:
            self._rimuovi_formato(inizio, fine)
        self._dopo_formattazione()

    def _rimuovi_formato(self, inizio, fine):
        for tag in self.testo.tag_names():
            if tag_logico(tag):
                self.testo.tag_remove(tag, inizio, fine)

    def _applica_formato(self, formato, inizio, fine):
        """Sostituisce il formato del pezzo di testo con quello dato."""
        self._rimuovi_formato(inizio, fine)
        if formato != PREDEFINITO:
            for tag in self._tag_da_formato(formato):
                self.testo.tag_add(tag, inizio, fine)

    # --- Annulla / ripeti della formattazione --------------------------------

    def _istantanea(self):
        return [(tag, [str(i) for i in self.testo.tag_ranges(tag)]) for tag in self.testo.tag_names()
                if tag_logico(tag)]

    def _ripristina(self, istantanea):
        for tag in self.testo.tag_names():
            if tag_logico(tag):
                self.testo.tag_remove(tag, "1.0", "end")
        for tag, intervalli in istantanea:
            if intervalli:
                self.testo.tag_add(tag, *intervalli)

    def _prima_di_formattare(self):
        self.da_annullare.append(self._istantanea())
        del self.da_annullare[:-MAX_ANNULLA]
        self.da_ripetere = []

    def _azzera_annulla(self):
        # Le istantanee valgono per il testo com'era: se il testo cambia non si possono più usare
        self.da_annullare = []
        self.da_ripetere = []

    def annulla_formato(self):
        if not self.da_annullare:
            self._messaggio("Niente da annullare")
            return
        self.da_ripetere.append(self._istantanea())
        self._ripristina(self.da_annullare.pop())
        self._dopo_formattazione()

    def ripeti_formato(self):
        if not self.da_ripetere:
            self._messaggio("Niente da ripetere")
            return
        self.da_annullare.append(self._istantanea())
        self._ripristina(self.da_ripetere.pop())
        self._dopo_formattazione()

    def _annulla_o_ripeti(self, ripeti):
        if self.modifica_testo.get():
            try:
                if ripeti:
                    self.testo.edit_redo()
                else:
                    self.testo.edit_undo()
            except tk.TclError:
                self._messaggio("Niente da " + ("ripetere" if ripeti else "annullare"))
        elif ripeti:
            self.ripeti_formato()
        else:
            self.annulla_formato()

    # --- Copia formato, selezione a blocco, formato di una riga su altre ------

    def _formato_cursore(self):
        """Formato del primo carattere selezionato o, senza selezione, di quello prima del cursore."""
        selezione = self.testo.tag_ranges("sel")
        indice = selezione[0] if selezione else self.testo.index("insert")
        if not selezione and self.testo.compare("insert", ">", "insert linestart"):
            indice = self.testo.index("insert-1c")
        return self._formato(self.testo.tag_names(indice))

    def copia_formato(self):
        if not self.righe:
            return
        self.formato_copiato = self._formato_cursore()
        f, b, i, s, c = self.formato_copiato
        descrizione = "%s %d%s%s, %s" % (FONT[f][0], s, " grassetto" if b else "", " corsivo" if i else "",
                                        COLORI[c][0])
        self._messaggio("Formato copiato: " + descrizione + ". Seleziona il testo e usa Incolla formato "
                        "(Ctrl+Maiusc+V)")

    def incolla_formato(self):
        if self.formato_copiato is None:
            self._messaggio("Prima copia un formato (Ctrl+Maiusc+C)")
            return
        selezione = self._selezione()
        if selezione is None:
            return
        self._prima_di_formattare()
        for inizio, fine in selezione:
            self._applica_formato(self.formato_copiato, inizio, fine)
        self._dopo_formattazione()

    def seleziona_blocco(self, riga_da, riga_a, colonna_da, colonna_a):
        """Seleziona le stesse colonne (colonna_a esclusa) su più righe."""
        riga_da, riga_a = sorted((riga_da, riga_a))
        colonna_da, colonna_a = sorted((colonna_da, colonna_a))
        self.testo.tag_remove("sel", "1.0", "end")
        for riga in range(max(1, riga_da), min(riga_a, len(self.righe)) + 1):
            fine = min(colonna_a, len(self.righe[riga - 1]))
            if colonna_da < fine:
                self.testo.tag_add("sel", "%d.%d" % (riga, colonna_da), "%d.%d" % (riga, fine))
        self._aggiorna_barra()

    def _blocco_inizio(self, evento):
        self.ancora_blocco = self.testo.index("@%d,%d" % (evento.x, evento.y))
        self.testo.mark_set("insert", self.ancora_blocco)
        self.testo.tag_remove("sel", "1.0", "end")
        self.testo.focus_set()
        return "break"

    def _blocco_trascina(self, evento):
        if self.ancora_blocco is None:
            return "break"
        fine = self.testo.index("@%d,%d" % (evento.x, evento.y))
        r1, c1 = (int(x) for x in self.ancora_blocco.split("."))
        r2, c2 = (int(x) for x in fine.split("."))
        self.testo.mark_set("insert", fine)
        self.seleziona_blocco(r1, r2, c1, c2)
        return "break"

    def _rilascio_alt(self, evento):
        if self.ancora_blocco is not None:
            self.ancora_blocco = None
            return "break"
        return None

    def dialogo_blocco(self):
        if not self.righe:
            self._messaggio("Prima apri un testo di esempio")
            return
        selezione = self.testo.tag_ranges("sel")
        if selezione:
            r1, c1 = (int(x) for x in str(selezione[0]).split("."))
            r2, c2 = (int(x) for x in str(selezione[-1]).split("."))
        else:
            r1 = r2 = int(self.testo.index("insert").split(".")[0])
            c1, c2 = 0, len(self.righe[r1 - 1])
        lunghezza = max(len(r) for r in self.righe)
        valori = chiedi_numeri(self, "Selezione a blocco",
                               "Seleziona le stesse colonne su più righe, per esempio la colonna degli importi. "
                               "Le colonne si contano da 0 come nella barra di stato; l'ultima è esclusa.",
                               [("Dalla riga", r1, 1, len(self.righe)), ("Alla riga", r2, 1, len(self.righe)),
                                ("Dalla colonna", min(c1, c2), 0, lunghezza),
                                ("Alla colonna", max(c1, c2), 0, lunghezza)])
        if valori:
            self.seleziona_blocco(*valori)
            self.testo.focus_set()

    def copia_formato_riga(self, sorgente, riga_da, riga_a):
        """Dà alle righe da riga_da a riga_a la stessa formattazione della riga sorgente.

        L'ultimo formato della riga sorgente continua fino a fine riga, come fa il servizio.
        """
        lunghezza = len(self.righe[sorgente - 1])
        formati = [self._formato(self.testo.tag_names("%d.%d" % (sorgente, c))) for c in range(lunghezza)]
        pezzi = modello_gui.pezzi(formati) or [(0, 0, PREDEFINITO)]
        self._prima_di_formattare()
        riga_da, riga_a = sorted((riga_da, riga_a))
        for riga in range(riga_da, riga_a + 1):
            if riga == sorgente or riga > len(self.righe):
                continue
            fine_riga = len(self.righe[riga - 1])
            self._rimuovi_formato("%d.0" % riga, "%d.end" % riga)
            for n, (da, a, formato) in enumerate(pezzi):
                if n == len(pezzi) - 1:
                    a = fine_riga
                a = min(a, fine_riga)
                if da < a:
                    self._applica_formato(formato, "%d.%d" % (riga, da), "%d.%d" % (riga, a))
        self._dopo_formattazione()

    def dialogo_copia_riga(self):
        if not self.righe:
            self._messaggio("Prima apri un testo di esempio")
            return
        riga = int(self.testo.index("insert").split(".")[0])
        selezione = self.testo.tag_ranges("sel")
        da, a = riga + 1, riga + 1
        if selezione:
            da = int(str(selezione[0]).split(".")[0])
            a = int(str(selezione[-1]).split(".")[0])
        n = len(self.righe)
        valori = chiedi_numeri(self, "Copia il formato di una riga",
                               "La formattazione della riga scelta viene data alle altre righe, colonna per "
                               "colonna (es. tutte le righe degli articoli uguali alla prima).",
                               [("Copia il formato della riga", riga, 1, n), ("Sulle righe da", min(da, n), 1, n),
                                ("alla riga", min(a, n), 1, n)])
        if valori:
            self.copia_formato_riga(*valori)

    # --- Zoom, modelli recenti ------------------------------------------------

    def _dimensione_vista(self, dimensione):
        return max(1, int(round(dimensione * self.zoom)))

    def cambia_zoom(self, passo):
        """passo: 1 ingrandisce, -1 riduce, 0 torna alla dimensione normale."""
        if passo == 0:
            nuovo = 1.0
        else:
            posizione = LIVELLI_ZOOM.index(self.zoom) + passo
            nuovo = LIVELLI_ZOOM[max(0, min(len(LIVELLI_ZOOM) - 1, posizione))]
        if nuovo == self.zoom:
            return
        self.zoom = nuovo
        self.testo.configure(font=(famiglia_tk(PREDEFINITO[0]), self._dimensione_vista(PREDEFINITO[3])))
        self._aggiorna_vista()
        self._messaggio("Zoom %d%% (solo nell'editor: il documento non cambia)" % round(nuovo * 100))
        self.preferenze["zoom"] = nuovo
        self._scrivi_preferenze()

    def _leggi_preferenze(self):
        return leggi_preferenze()

    def _scrivi_preferenze(self):
        scrivi_preferenze(self.preferenze)

    def _aggiungi_recente(self, percorso):
        percorso = os.path.abspath(percorso)
        recenti = [p for p in self.preferenze.get("recenti", [])
                   if os.path.normcase(p) != os.path.normcase(percorso)]
        self.preferenze["recenti"] = ([percorso] + recenti)[:MAX_RECENTI]
        self._scrivi_preferenze()

    def _elenca_recenti(self):
        self.menu_recenti.delete(0, "end")
        recenti = [p for p in self.preferenze.get("recenti", []) if os.path.isfile(p)]
        if not recenti:
            self.menu_recenti.add_command(label="(nessuno)", state="disabled")
        for n, percorso in enumerate(recenti):
            self.menu_recenti.add_command(label="%d  %s" % (n + 1, os.path.basename(percorso)),
                                          command=lambda p=percorso: self.apri_recente(p))

    def apri_recente(self, percorso):
        if not self._conferma_perdita():
            return
        if not os.path.isfile(percorso):
            messagebox.showerror("Modelli recenti", "Il modello non esiste più:\n" + percorso, parent=self)
            return
        self.apri_modello(percorso)

    def _dopo_formattazione(self):
        self._segna_modificato()
        self._aggiorna_vista()
        self._aggiorna_barra()
        self.testo.focus_set()

    def _tasto(self, evento):
        # Senza la spunta "Modifica testo" passano solo spostamenti, selezione e copia
        if self.solo_invio and evento.keysym not in TASTI_MOVIMENTO:
            return "break"
        if self.modifica_testo.get() or evento.keysym in TASTI_MOVIMENTO:
            return None
        if evento.state & 0x4 and evento.keysym.lower() in TASTI_CONTROL:
            return None
        return "break"

    def _cambia_modifica(self):
        # Sfondo diverso per ricordare che si sta cambiando il testo e non solo il formato
        self.testo.configure(background="#fffbe6" if self.modifica_testo.get() else self.sfondo_testo)
        if self.modifica_testo.get():
            self._messaggio("Modifica testo attiva: le modifiche vanno nella copia del testo salvata con il "
                            "modello, il file di partenza non cambia")
        else:
            self._messaggio("Modifica testo disattivata: si cambia solo la formattazione")
        self.testo.focus_set()

    def _testo_modificato(self, evento):
        if not self.testo.edit_modified():
            return
        self.testo.edit_modified(False)
        righe = self.testo.get("1.0", "end-1c").split("\n")
        if righe == self.righe:
            return
        self.righe = righe
        self.testo_cambiato = True
        self._azzera_annulla()
        self._segna_oltre()
        self._segna_modificato()
        self._aggiorna_vista()

    def _segna_oltre(self):
        self.testo.tag_remove("oltre", "1.0", "end")
        for n, riga in enumerate(self.righe):
            if len(riga) > COLONNA_MASSIMA:
                self.testo.tag_add("oltre", "%d.%d" % (n + 1, COLONNA_MASSIMA), "%d.end" % (n + 1))
        self.testo.tag_lower("oltre")

    def _seleziona_tutto(self):
        self.testo.tag_add("sel", "1.0", "end-1c")
        self._aggiorna_barra()

    def _aggiorna_barra(self):
        selezione = self.testo.tag_ranges("sel")
        f, b, i, s, c = self._formato_cursore()
        self.scelta_font.set(FONT[f][0])
        self.scelta_dimensione.set(str(s))
        self.scelta_colore.set(COLORI[c][0])
        self.pulsante_b.configure(relief="sunken" if b else "raised")
        self.pulsante_i.configure(relief="sunken" if i else "raised")
        riga, colonna = self.testo.index("insert").split(".")
        testo = "Riga %s, colonna %s" % (riga, colonna)
        if len(selezione) > 2:
            r1, c1 = str(selezione[0]).split(".")
            r2, c2 = str(selezione[-1]).split(".")
            testo += "   -   blocco: righe %s-%s, colonne %s-%s" % (r1, r2, c1, c2)
        elif selezione:
            r1, c1 = str(selezione[0]).split(".")
            r2, c2 = str(selezione[1]).split(".")
            if r1 == r2:
                testo += "   -   selezione: colonne %s-%s" % (c1, c2)
            else:
                testo += "   -   selezione: righe %s-%s" % (r1, r2)
        self._messaggio(testo)
        self._disegna_righello()
        if self.finestra_filigrana is not None and self.finestra_filigrana.winfo_exists():
            self.finestra_filigrana.mostra_cursore(int(riga) - 1, int(colonna))

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
        self._segna_oltre()
        # Il testo appena caricato non è una modifica e non si può annullare
        self.testo.edit_reset()
        self.testo.edit_modified(False)
        self.testo_cambiato = False
        self._azzera_annulla()
        self._aggiorna_vista()
        self.testo.mark_set("insert", "1.0")
        self.testo.focus_set()
        self._programma_filigrana()

    def _leggi_testo(self, percorso):
        # Come lo legge il servizio: UTF-8 (Linux) o la codifica di Windows
        return leggi_testo(percorso).split("\n")

    def _intestazione(self):
        destinatario, coda = modello_gui.invio_da_salvare(self.invio.get(), self.destinatario.get(),
                                                          self.pannello.righe(), self.pannello.scelta_whatsapp())
        return {
            "testo": modello_gui.SEGNAPOSTO + ".txt",
            "azione": self.azione.get(),
            "uscita": self.uscita.get().strip(),
            "template": self.template.get().strip() if not self.word.get() else "",
            "rotazione": self.rotazione.get() if not self.word.get() else "0",
            "invio": self.invio.get(),
            "destinatario": destinatario,
            "coda": coda,
        }

    def _imposta_intestazione(self, intestazione):
        self.word.set(intestazione["azione"] in modello_gui.AZIONI_WORD)
        self._cambia_tipo(iniziale=True)
        self.azione.set(intestazione["azione"])
        self.template.set(intestazione["template"])
        self.rotazione.set(intestazione["rotazione"])
        self.uscita.set(intestazione["uscita"])
        self.invio.set(intestazione["invio"] if intestazione["invio"] in modello_gui.INVII else "0")
        destinatario, coda, whatsapp = modello_gui.invio_da_mostrare(intestazione)
        self.destinatario.set(destinatario)
        self.pannello.imposta(coda, whatsapp)

    def _cambia_tipo(self, iniziale=False):
        word = self.word.get()
        azioni = modello_gui.AZIONI_WORD if word else modello_gui.AZIONI_PDF
        self.scelta_azione.imposta_scelte({a: modello_gui.DESCRIZIONI_AZIONI[a] for a in azioni})
        if self.azione.get() not in self.scelta_azione.codici():
            self.azione.set("docx" if word else "file")
        # Senza documento (solo invio) template e rotazione non servono
        stato = "disabled" if word or self.solo_invio else "normal"
        self.scelta_template.configure(state=stato)
        self.scelta_rotazione.configure(state="disabled" if word or self.solo_invio else "readonly")
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

    def nuovo_segnaposto(self):
        """Nuovo modello da un testo di esempio, con le x al posto dei caratteri."""
        if self.nuovo():
            self.segnaposto()

    def segnaposto(self, opzioni=None):
        """Sostituisce il testo di esempio con le x, tenendo la formattazione già fatta.

        opzioni: (x in più, simboli); senza, le chiede con una finestra.
        """
        if not self.righe:
            self._messaggio("Prima apri un testo di esempio")
            return False
        if opzioni is None:
            finestra = DialogoSegnaposto(self, self.righe)
            self.wait_window(finestra)
            opzioni = finestra.risultato
            if opzioni is None:
                return False
        nuove = modello_gui.testo_segnaposto(self.righe, *opzioni)
        # Le colonne non cambiano, quindi i formati restano; le x aggiunte in fondo alla riga
        # prendono il formato dell'ultimo carattere
        formati = self._formati_testo()
        for n, riga in enumerate(nuove):
            ultimo = formati[n][-1] if formati[n] else PREDEFINITO
            formati[n] = formati[n] + [ultimo] * (len(riga) - len(formati[n]))
        self._mostra_testo(nuove, formati)
        # Il testo con le x va nella copia accanto al modello, al salvataggio
        self.testo_cambiato = True
        self._segna_modificato()
        self._messaggio("Testo sostituito con le x: al salvataggio va nella copia accanto al modello "
                        "(il file di partenza non cambia)")
        return True

    def nuovo(self):
        """Nuovo modello da un testo di esempio; False se non è stato scelto nessun testo."""
        if not self._conferma_perdita():
            return False
        percorso = filedialog.askopenfilename(parent=self, title="Testo di esempio (es. una fattura)",
                                              filetypes=[("File di testo", "*.txt"), ("Tutti i file", "*.*")])
        if not percorso:
            return False
        self.percorso_testo = percorso
        self.percorso_modello = None
        self._imposta_solo_invio(False)
        self._imposta_intestazione(modello_gui.nuovo_modello(self.word.get()))
        self._mostra_testo(self._leggi_testo(percorso))
        self.modificato = False
        self._aggiorna_titolo()
        return True

    def nuovo_solo_invio(self):
        """Nuovo modello che non crea documenti e fa solo l'invio (email, Telegram, WhatsApp, anche con
        file che ci sono già sul computer): non serve un testo di esempio e il servizio lo elabora
        anche senza .txt. False se si è rinunciato."""
        if not self._conferma_perdita():
            return False
        self.percorso_testo = None
        self.percorso_modello = None
        intestazione = modello_gui.nuovo_modello(False)
        intestazione.update(testo="", uscita="", invio="email")
        self._imposta_solo_invio(True)
        self._imposta_intestazione(intestazione)
        self.modificato = False
        self._aggiorna_titolo()
        self._messaggio("Modello solo invio: scegli il canale, il destinatario e cosa mandare, poi salva")
        return True

    def _imposta_solo_invio(self, solo_invio):
        """Passa al modello di solo invio (al posto del testo una spiegazione) o a quello normale."""
        self.solo_invio = solo_invio
        self.campo_uscita.configure(state="disabled" if solo_invio else "normal")
        if solo_invio:
            self._mostra_testo([])
            self.testo.tag_configure("spiegazione", foreground="gray")
            self.testo.insert("1.0", TESTO_SOLO_INVIO, "spiegazione")
            self.testo.edit_modified(False)

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
        if intestazione["azione"] == "excel":
            messagebox.showinfo("Apri modello", "È un modello Excel: aprilo con EditorExcel.", parent=self)
            return
        if intestazione["azione"] == "barcode":
            return FinestraBarcode(self, percorso)
        solo_invio = False
        if percorso_testo is None:
            # Il testo di esempio è salvato accanto al modello con lo stesso nome
            percorso_testo = os.path.splitext(percorso)[0] + ".txt"
            if os.path.isfile(percorso_testo):
                pass
            elif not intestazione["uscita"].strip():
                # Non crea documenti: è un modello di solo invio, il testo non serve
                solo_invio = True
                percorso_testo = None
            else:
                messagebox.showinfo("Testo di esempio", "Scegli un file di testo di esempio per questo modello.",
                                    parent=self)
                percorso_testo = filedialog.askopenfilename(parent=self, title="Testo di esempio",
                                                            filetypes=[("File di testo", "*.txt")])
                if not percorso_testo:
                    return
        email_aggiornata = modello_gui.aggiorna_email(intestazione)
        self._imposta_solo_invio(solo_invio)
        self._imposta_intestazione(intestazione)
        if not solo_invio:
            righe = self._leggi_testo(percorso_testo)
            self._mostra_testo(righe, modello_gui.leggi_formati(linee, righe, self.word.get()))
        self.percorso_modello = percorso
        self.percorso_testo = percorso_testo
        # Un modello email di prima, portato al formato nuovo, va salvato
        self.modificato = email_aggiornata
        self._aggiorna_titolo()
        self._aggiungi_recente(percorso)
        if email_aggiornata:
            self._messaggio("L'indirizzo email è passato nel campo Destinatario: salva per aggiornare il modello")

    def _linee_formato(self):
        try:
            return modello_gui.linee_formato(self._formati_testo(), self.word.get())
        except modello_gui.ErroreModello as e:
            messagebox.showerror("Modello", str(e), parent=self)
            return None

    def salva(self, con_nome=False):
        if not self.righe and not self.solo_invio:
            self._messaggio("Prima apri un testo di esempio")
            return False
        if self.solo_invio and self.invio.get() == "0":
            messagebox.showerror("Salva", "Il modello non crea documenti: scegli un invio (email, Telegram o "
                                 "WhatsApp).", parent=self)
            return False
        # Senza testo non ci sono formati: subito "fine" e poi i dati dell'invio
        linee = ["fine"] if self.solo_invio else self._linee_formato()
        if linee is None:
            return False
        percorso = self.percorso_modello
        if con_nome or percorso is None:
            os.makedirs(CARTELLA_MODELLI, exist_ok=True)
            cartella, nome = proposta_salvataggio(self.percorso_modello, self.percorso_testo, CARTELLA_MODELLI)
            percorso = filedialog.asksaveasfilename(parent=self, title="Salva modello", initialdir=cartella,
                                                    initialfile=nome, defaultextension=".gui",
                                                    filetypes=[("Modelli", "*.gui")])
            if not percorso:
                return False
        intestazione = self._intestazione()
        if not intestazione["uscita"] and intestazione["invio"] == "0" and not messagebox.askokcancel(
                "Salva", "Il campo Documento creato è vuoto e non c'è nessun invio: il modello non farà niente.\n"
                         "Salvare comunque?", parent=self):
            return False
        # Nella prima riga il testo con lo stesso nome del modello (applicandolo diventa quello del file);
        # vuota nel solo invio, che non ha testo
        intestazione["testo"] = "" if self.solo_invio else os.path.splitext(os.path.basename(percorso))[0] + ".txt"
        modello_gui.scrivi_gui(percorso, intestazione, linee)
        # Copia del testo di esempio accanto al modello, per poterlo riaprire
        copia_testo = os.path.splitext(percorso)[0] + ".txt"
        if self.solo_invio:
            pass
        elif self.testo_cambiato:
            # Testo modificato nell'editor: si salva quello, con la stessa codifica con cui si legge
            with open(copia_testo, "w", errors="replace") as f:
                f.write("\n".join(self.righe))
            self.percorso_testo = copia_testo
            self.testo_cambiato = False
        elif self.percorso_testo and os.path.normcase(os.path.abspath(self.percorso_testo)) != \
                os.path.normcase(os.path.abspath(copia_testo)):
            shutil.copyfile(self.percorso_testo, copia_testo)
            self.percorso_testo = copia_testo
        self.percorso_modello = percorso
        self.modificato = False
        self._aggiorna_titolo()
        self._messaggio("Modello salvato: " + percorso)
        self._aggiungi_recente(percorso)
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
        if self.solo_invio:
            # Niente testo: il .gui va da solo nella cartella di lavoro e il servizio fa l'invio
            if not messagebox.askokcancel("Applica", "Il modello verrà copiato nella cartella di lavoro\n%s\n"
                                          "e il servizio farà l'invio. Continuare?" % lavoro, parent=self):
                return
            creato = modello_gui.applica_solo_invio(self.percorso_modello, lavoro)
            messagebox.showinfo("Applica", "Creato:\n" + os.path.basename(creato), parent=self)
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

    def crea_documento(self, cartella, senza_template=False, avvisa=True):
        """Crea il documento con il codice del servizio; restituisce il percorso.

        senza_template=True: al posto del template una pagina bianca della stessa misura.
        avvisa=False: niente finestre di errore, gli errori diventano ErroreModello.
        """
        if avvisa:
            linee = self._linee_formato()
            if linee is None:
                return None
        else:
            linee = modello_gui.linee_formato(self._formati_testo(), self.word.get())
        intestazione = self._intestazione()
        estensione = ".docx" if self.word.get() else ".pdf"
        intestazione.update(testo="anteprima.txt", uscita="anteprima" + estensione,
                            azione="docx" if self.word.get() else "file", invio="0", destinatario="", coda=[])
        config = dict(self.config_servizio, cartella_lavoro=cartella, cartella_output=cartella)
        percorso_template = os.path.join(config.get("cartella_template", ""), intestazione["template"])
        if intestazione["template"] and not os.path.isfile(percorso_template):
            errore = "Template non trovato nella cartella dei template: " + intestazione["template"]
            if not avvisa:
                raise modello_gui.ErroreModello(errore)
            messagebox.showerror("Anteprima", errore, parent=self)
            return None
        if senza_template and intestazione["template"]:
            # Una pagina bianca grande come il template: il testo va negli stessi punti
            # (orizzontale o verticale), ma senza il testo del template
            import pymupdf
            with pymupdf.open(percorso_template) as template:
                pagina = template[0].rect
            vuoto = pymupdf.open()
            vuoto.new_page(width=pagina.width, height=pagina.height)
            vuoto.save(os.path.join(cartella, "template_vuoto.pdf"))
            vuoto.close()
            intestazione["template"] = "template_vuoto.pdf"
            config["cartella_template"] = cartella
        with open(os.path.join(cartella, "anteprima.txt"), "w", errors="replace") as f:
            f.write("\n".join(self.righe))
        percorso_gui = os.path.join(cartella, "anteprima.gui")
        modello_gui.scrivi_gui(percorso_gui, intestazione, linee)
        with open(percorso_gui, "r") as f:
            guida = f.readlines()
        servizio = importlib.import_module("SingleGUY")
        if self.word.get():
            servizio.docx(guida, config)
        else:
            servizio.pdf(guida, config)
        return os.path.join(cartella, "anteprima" + estensione)

    def anteprima(self):
        if self.solo_invio:
            self._messaggio("Il modello fa solo l'invio: non crea documenti da vedere")
            return
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
            apri_con_avviso(percorso, self)
        else:
            FinestraAnteprima(self, percorso)

    def filigrana(self):
        if self.solo_invio:
            self._messaggio("Il modello fa solo l'invio: non crea documenti da vedere")
            return
        if not self.righe:
            self._messaggio("Prima apri un testo di esempio")
            return
        if self.word.get():
            messagebox.showinfo("Filigrana", "La filigrana c'è solo per i modelli PDF: per Word usa l'Anteprima.",
                                parent=self)
            return
        if self.finestra_filigrana is not None and self.finestra_filigrana.winfo_exists():
            self.finestra_filigrana.lift()
            self.finestra_filigrana.aggiorna()
        else:
            self.finestra_filigrana = FinestraFiligrana(self)

    def _programma_filigrana(self):
        """Dopo una modifica la filigrana si ricrea, ma solo quando si smette per un attimo di modificare."""
        if self.finestra_filigrana is None or not self.finestra_filigrana.winfo_exists():
            return
        if self.attesa_filigrana is not None:
            self.after_cancel(self.attesa_filigrana)
        self.attesa_filigrana = self.after(600, self._aggiorna_filigrana)

    def _aggiorna_filigrana(self):
        self.attesa_filigrana = None
        if self.finestra_filigrana is not None and self.finestra_filigrana.winfo_exists():
            self.finestra_filigrana.aggiorna()

    def vai_a(self, riga, colonna):
        """Porta il cursore alla riga (da 0) e colonna indicate, ad esempio da un clic sulla filigrana."""
        indice = "%d.%d" % (riga + 1, colonna)
        self.testo.tag_remove("sel", "1.0", "end")
        self.testo.mark_set("insert", indice)
        self.testo.see(indice)
        self.testo.focus_set()
        self._aggiorna_barra()

    # --- Aiuto ---------------------------------------------------------------

    def apri_guida(self):
        apri_guida(self)

    def _informazioni(self):
        messagebox.showinfo("Informazioni", "Editor modelli di FileGUY\n\nPrepara i modelli .gui con la "
                            "formattazione di fatture, bolle e altri documenti, da applicare ai file di testo "
                            "elaborati dal servizio.\n\nGuida completa: Aiuto > Guida all'uso (F1)", parent=self)

    # --- Stato, titolo, chiusura -------------------------------------------

    def _segna_modificato(self):
        if self.righe or self.solo_invio:
            self.modificato = True
            self._aggiorna_titolo()
            self._programma_filigrana()

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
        import pymupdf
        self.title("Anteprima - " + os.path.basename(genitore.percorso_modello or "nuovo modello"))
        dimensiona(self, 820, 900)
        self.documento = pymupdf.open(percorso)
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
        ttk.Button(barra, text="Apri nel visualizzatore",
                   command=lambda: apri_con_avviso(self.percorso, self)).pack(side="right")

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


ETICHETTE_DESTINATARIO = {"0": "Destinatario", "email": "Indirizzo email", "telegram": "Chat Telegram",
                          "whatsapp": "Numero WhatsApp"}


class PannelloInvio(ttk.Frame):
    """I dati dell'invio (le righe dopo "fine" del .gui) con campi diversi per ogni canale.

    Email: oggetto, allegati in più, testo e immagini nel testo. Telegram e WhatsApp: il documento
    da mandare e il testo; con un template WhatsApp un campo per ogni valore e l'anteprima del messaggio.
    editor: la finestra con le variabili invio, uscita, config_servizio e i metodi
    _segna_modificato e _messaggio (EditorModelli o EditorExcel).
    """

    def __init__(self, genitore, editor):
        super().__init__(genitore)
        self.editor = editor
        lingua = editor.config_servizio.get("whatsapp_lingua") or "it"
        # Lo stato resta qui anche quando i campi si rifanno (cambiando canale o template)
        self.modo = tk.StringVar(value="messaggio")
        self.nome_template = tk.StringVar()
        self.lingua = tk.StringVar(value=lingua)
        self.elenco_template = []
        self.oggetto = tk.StringVar()
        self.allegati = []
        self.immagini = []
        self.manda_documento = tk.BooleanVar(value=True)
        self.documento = tk.StringVar()
        self.testo = ""
        self.valori = []
        self.campo_attivo = None
        self.dati_template = None
        self.casella_testo = None
        self.casella_valori = None
        for variabile in (self.modo, self.nome_template, self.lingua, self.oggetto, self.manda_documento,
                          self.documento):
            variabile.trace_add("write", lambda *a: self.editor._segna_modificato())
        # Senza documento creato (campo vuoto) alcune spiegazioni cambiano: i campi si rifanno
        self.senza_documento = not editor.uscita.get().strip()
        editor.uscita.trace_add("write", lambda *a: self._uscita_cambiata())
        self.columnconfigure(1, weight=1)

    def _uscita_cambiata(self):
        senza_documento = not self.editor.uscita.get().strip()
        if senza_documento != self.senza_documento:
            self.senza_documento = senza_documento
            self.ricostruisci()

    # --- Dati -----------------------------------------------------------------------

    def imposta(self, coda, whatsapp):
        """Riempie i campi con le righe dopo "fine" (e la scelta WhatsApp di invio_da_mostrare)."""
        # Le caselle mostrano ancora i dati di prima: non vanno rilette
        self.casella_testo = self.casella_valori = None
        self.modo.set(whatsapp["modo"])
        self.nome_template.set(whatsapp["template"])
        self.lingua.set(whatsapp["lingua"] or self.editor.config_servizio.get("whatsapp_lingua") or "it")
        invio = self.editor.invio.get()
        campi = modello_gui.campi_invio(invio, whatsapp["modo"], coda)
        self.oggetto.set(campi["oggetto"])
        self.allegati = campi["allegati"]
        self.immagini = campi["immagini"]
        # Il documento creato si mostra come campo vuoto: segue il nome del documento anche se cambia
        documento = campi["documento"]
        self.manda_documento.set(documento != "" if coda else invio != "whatsapp" or whatsapp["modo"] != "template")
        self.documento.set("" if documento == self.editor.uscita.get().strip() else documento)
        self.testo = campi["testo"]
        self.valori = campi["valori"]
        self.ricostruisci()

    def _leggi_caselle(self):
        """Il testo e i valori scritti nelle caselle di più righe."""
        if self.casella_testo is not None and self.casella_testo.winfo_exists():
            self.testo = self.casella_testo.get("1.0", "end-1c")
        if self.casella_valori is not None and self.casella_valori.winfo_exists():
            self.valori = [r for r in self.casella_valori.get("1.0", "end-1c").split("\n") if r.strip()]

    def righe(self):
        """Le righe da scrivere dopo "fine"."""
        self._leggi_caselle()
        documento = ""
        if self.manda_documento.get():
            documento = self.documento.get().strip() or self.editor.uscita.get().strip()
        if self._template() and self.dati_template and self.dati_template["intestazione"] != "DOCUMENT":
            documento = ""
        campi = {"oggetto": self.oggetto.get(), "allegati": self.allegati, "immagini": self.immagini,
                 "documento": documento,
                 "testo": self.testo, "valori": self.valori}
        return modello_gui.coda_invio(self.editor.invio.get(), self.modo.get(), campi)

    def scelta_whatsapp(self):
        return {"modo": self.modo.get(), "template": self.nome_template.get(), "lingua": self.lingua.get()}

    def _template(self):
        return self.editor.invio.get() == "whatsapp" and self.modo.get() == "template"

    # --- Campi --------------------------------------------------------------------

    def ricostruisci(self):
        """Rifà i campi per il canale scelto (e, con WhatsApp, per messaggio o template)."""
        self._leggi_caselle()
        for figlio in self.winfo_children():
            figlio.destroy()
        self.casella_testo = self.casella_valori = None
        self.anteprima = None
        invio = self.editor.invio.get()
        if invio not in ("email", "telegram", "whatsapp"):
            self.grid_remove()
            return
        self.grid()
        riga = 0
        if invio == "whatsapp":
            riga = self._riga_whatsapp(riga)
        if invio == "email":
            riga = self._campi_email(riga)
        elif self._template():
            self.dati_template = self._cerca_template()
            riga = self._campi_template(riga)
        else:
            riga = self._riga_documento(riga)
            riga = self._casella_testo(riga)
        ttk.Label(self, text=self._aiuto(invio), foreground="gray", wraplength=900, justify="left").grid(
            row=riga, column=1, sticky="w")

    def _etichetta(self, riga, testo):
        ttk.Label(self, text=testo).grid(row=riga, column=0, sticky="nw", pady=(4, 0), padx=(0, 6))

    def _riga_whatsapp(self, riga):
        self._etichetta(riga, "Cosa mandare")
        cornice = ttk.Frame(self)
        cornice.grid(row=riga, column=1, sticky="w", pady=(4, 0))
        for testo, valore in (("Messaggio libero", "messaggio"), ("Template approvato da Meta", "template")):
            ttk.Radiobutton(cornice, text=testo, variable=self.modo, value=valore,
                            command=self._modo_cambiato).pack(side="left", padx=(0, 10))
        if self.modo.get() == "template":
            valori = ["%s (%s)" % (t["nome"], t["lingua"]) for t in self.elenco_template]
            scelta = ttk.Combobox(cornice, textvariable=self.nome_template, values=valori, width=30)
            scelta.pack(side="left", padx=(6, 0))
            scelta.bind("<<ComboboxSelected>>", lambda e: self._template_scelto())
            # Un nome scritto a mano vale con Invio (la tendina vale appena si sceglie)
            scelta.bind("<Return>", lambda e: self.ricostruisci())
            ttk.Label(cornice, text="Lingua").pack(side="left", padx=(8, 4))
            ttk.Entry(cornice, textvariable=self.lingua, width=6).pack(side="left")
            ttk.Button(cornice, text="Scarica elenco", command=self.scarica_template).pack(side="left", padx=(8, 0))
            ttk.Button(cornice, text="Nuovo template...",
                       command=lambda: FinestraNuovoTemplate(self)).pack(side="left", padx=(6, 0))
        return riga + 1

    def _campi_email(self, riga):
        self._etichetta(riga, "Oggetto")
        ttk.Entry(self, textvariable=self.oggetto).grid(row=riga, column=1, sticky="we", pady=(4, 0))
        riga += 1
        riga = self._elenco_file(riga, "Allegati in più", self.allegati, [("Tutti i file", "*.*")])
        documento = "Il documento creato è sempre allegato" if self.editor.uscita.get().strip() else \
            "Nessun documento creato (campo Documento creato vuoto): si mandano solo testo e file indicati"
        ttk.Label(self, text=documento + ". Gli altri file vanno indicati con il percorso sul computer dove gira "
                  "il servizio.", foreground="gray").grid(row=riga, column=1, sticky="w")
        riga = self._casella_testo(riga + 1)
        riga = self._elenco_file(riga, "Immagini\nnel testo", self.immagini,
                                 [("Immagini", "*.png *.jpg *.jpeg *.gif"), ("Tutti i file", "*.*")])
        config = self.editor.config_servizio
        firma = "impostata" if (config.get("email_firma") or config.get("email_firma_immagine")) else "nessuna"
        ttk.Label(self, text="Le immagini si vedono dopo il testo, non come allegati. In fondo va la firma di "
                  "ConfigManager (testo e immagine; ora: %s)." % firma, foreground="gray").grid(row=riga, column=1,
                                                                                               sticky="w")
        return riga + 1

    def _elenco_file(self, riga, etichetta, lista, tipi):
        """Un elenco di percorsi (allegati, immagini) con i pulsanti per aggiungerne e toglierne."""
        self._etichetta(riga, etichetta)
        cornice = ttk.Frame(self)
        cornice.grid(row=riga, column=1, sticky="we", pady=(4, 0))
        elenco = tk.Listbox(cornice, height=3, activestyle="none")
        elenco.pack(side="left", fill="x", expand=True)
        for percorso in lista:
            elenco.insert("end", percorso)
        pulsanti = ttk.Frame(cornice)
        pulsanti.pack(side="left", fill="y", padx=(6, 0))
        ttk.Button(pulsanti, text="Aggiungi...",
                   command=lambda: self._aggiungi_file(elenco, lista, etichetta.replace("\n", " "), tipi)).pack(fill="x")
        ttk.Button(pulsanti, text="Rimuovi", command=lambda: self._rimuovi_file(elenco, lista)).pack(fill="x",
                                                                                                   pady=(2, 0))
        return riga + 1

    def _aggiungi_file(self, elenco, lista, titolo, tipi):
        percorsi = filedialog.askopenfilenames(parent=self.winfo_toplevel(), title=titolo, filetypes=tipi)
        for percorso in percorsi:
            percorso = os.path.normpath(percorso)
            lista.append(percorso)
            elenco.insert("end", percorso)
        if percorsi:
            self.editor._segna_modificato()

    def _rimuovi_file(self, elenco, lista):
        for indice in reversed(elenco.curselection()):
            elenco.delete(indice)
            del lista[indice]
            self.editor._segna_modificato()

    def _riga_documento(self, riga, etichetta="Documento"):
        self._etichetta(riga, etichetta)
        cornice = ttk.Frame(self)
        cornice.grid(row=riga, column=1, sticky="we", pady=(4, 0))
        ttk.Checkbutton(cornice, text="Manda il documento", variable=self.manda_documento,
                        command=self._aggiorna_anteprima).pack(side="left")
        campo = ttk.Entry(cornice, textvariable=self.documento, width=30)
        campo.pack(side="left", padx=(8, 0))
        campo.bind("<KeyRelease>", lambda e: self._aggiorna_anteprima())
        campo.bind("<FocusIn>", lambda e: self._attiva_campo(0))
        campo.bind("<FocusOut>", lambda e: self._attiva_campo(None))
        ttk.Button(cornice, text="Scegli file...", command=self._scegli_documento).pack(side="left", padx=(6, 0))
        uscita = self.editor.uscita.get().strip()
        ttk.Label(cornice, text="vuoto = il documento creato (%s); un altro file con il percorso completo" % uscita
                  if uscita else "vuoto = nessun documento; un file già sul computer con il percorso completo",
                  foreground="gray").pack(side="left", padx=(8, 0))
        return riga + 1

    def _scegli_documento(self):
        """Un file che c'è già sul computer dove gira il servizio, da mandare al posto del documento creato."""
        percorso = filedialog.askopenfilename(parent=self.winfo_toplevel(), title="File da mandare")
        if percorso:
            self.manda_documento.set(True)
            self.documento.set(os.path.normpath(percorso))
            self._aggiorna_anteprima()

    def _casella_testo(self, riga):
        self._etichetta(riga, "Testo")
        self.casella_testo = tk.Text(self, height=5, width=60, wrap="word", undo=True, font=("Segoe UI", 9))
        self.casella_testo.grid(row=riga, column=1, sticky="we", pady=(4, 0))
        self.casella_testo.insert("1.0", self.testo)
        self.casella_testo.edit_modified(False)
        self.casella_testo.bind("<<Modified>>", self._casella_modificata)
        return riga + 1

    def _casella_modificata(self, evento):
        if evento.widget.edit_modified():
            evento.widget.edit_modified(False)
            self.editor._segna_modificato()

    # --- Template WhatsApp ----------------------------------------------------------

    def _campi_template(self, riga):
        dati = self.dati_template
        if dati is None:
            # Template non nell'elenco (non ancora scaricato): documento e valori uno per riga
            riga = self._riga_documento(riga, "Documento\n(intestazione)")
            self._etichetta(riga, "Valori")
            self.casella_valori = tk.Text(self, height=4, width=60, undo=True, font=("Segoe UI", 9))
            self.casella_valori.grid(row=riga, column=1, sticky="we", pady=(4, 0))
            self.casella_valori.insert("1.0", "\n".join(self.valori))
            self.casella_valori.edit_modified(False)
            self.casella_valori.bind("<<Modified>>", self._casella_modificata)
            return riga + 1
        if dati["intestazione"] == "DOCUMENT":
            riga = self._riga_documento(riga, "Documento\n(intestazione)")
        self._etichetta(riga, "Valori")
        cornice = ttk.Frame(self)
        cornice.grid(row=riga, column=1, sticky="we", pady=(4, 0))
        cornice.columnconfigure(1, weight=1)
        campi = ttk.Frame(cornice)
        campi.grid(row=0, column=0, sticky="nw")
        self.valori += [""] * (dati["parametri"] - len(self.valori))
        self.variabili_valori = []
        for n in range(1, dati["parametri"] + 1):
            ttk.Label(campi, text="{{%d}}" % n).grid(row=n, column=0, sticky="w", pady=1)
            valore = tk.StringVar(value=self.valori[n - 1])
            campo = ttk.Entry(campi, textvariable=valore, width=28)
            campo.grid(row=n, column=1, sticky="w", padx=(6, 0), pady=1)
            campo.bind("<FocusIn>", lambda e, _n=n: self._attiva_campo(_n))
            campo.bind("<FocusOut>", lambda e: self._attiva_campo(None))
            valore.trace_add("write", lambda *a: self._valore_cambiato())
            self.variabili_valori.append(valore)
        if not dati["parametri"]:
            ttk.Label(campi, text="Questo template non ha valori da riempire", foreground="gray").grid(
                row=0, column=0)
        # L'anteprima del messaggio come arriva su WhatsApp, con i valori evidenziati
        self.anteprima = tk.Text(cornice, height=7, width=40, wrap="word", relief="flat", background="#dcf8c6",
                                 padx=8, pady=6, font=("Segoe UI", 9), cursor="arrow")
        self.anteprima.grid(row=0, column=1, sticky="nsew", padx=(12, 0))
        self.anteprima.tag_configure("titolo", font=("Segoe UI", 9, "bold"))
        self.anteprima.tag_configure("piede", foreground="#667781", font=("Segoe UI", 8))
        self.anteprima.tag_configure("campo", background="#fff3a0", font=("Segoe UI", 9, "bold"))
        self.anteprima.tag_configure("attivo", background="#ffb84d")
        self._aggiorna_anteprima()
        return riga + 1

    def _valore_cambiato(self):
        self.valori = [v.get() for v in self.variabili_valori]
        self.editor._segna_modificato()
        self._aggiorna_anteprima()

    def _attiva_campo(self, numero):
        self.campo_attivo = numero
        self._aggiorna_anteprima()

    def _aggiorna_anteprima(self):
        """Il messaggio come arriva su WhatsApp, con i valori (o i {{n}} vuoti) evidenziati."""
        from whatsapp_meta import SEGNAPOSTO_TEMPLATE
        dati, vista = self.dati_template, getattr(self, "anteprima", None)
        if dati is None or vista is None or not vista.winfo_exists():
            return
        vista.configure(state="normal")
        vista.delete("1.0", "end")
        if dati["intestazione"] == "DOCUMENT":
            documento = self.documento.get().strip() or self.editor.uscita.get().strip()
            vista.insert("end", "📄 " + (documento if self.manda_documento.get() else "(nessun documento)") + "\n",
                         ("titolo", "campo") + (("attivo",) if self.campo_attivo == 0 else ()))
        elif dati["testo_intestazione"]:
            vista.insert("end", dati["testo_intestazione"] + "\n", "titolo")
        inizio = 0
        for segnaposto in SEGNAPOSTO_TEMPLATE.finditer(dati["testo"]):
            vista.insert("end", dati["testo"][inizio:segnaposto.start()])
            numero = int(segnaposto.group(1))
            valore = self.valori[numero - 1] if numero - 1 < len(self.valori) else ""
            vista.insert("end", valore or "{{%d}}" % numero,
                         ("campo",) + (("attivo",) if self.campo_attivo == numero else ()))
            inizio = segnaposto.end()
        vista.insert("end", dati["testo"][inizio:])
        if dati["piede"]:
            vista.insert("end", "\n" + dati["piede"], "piede")
        vista.configure(state="disabled")

    def _aiuto(self, invio):
        if invio == "email":
            return "Puoi usare {nome} (il nome del file di testo, senza .txt) in tutti i campi."
        if invio == "telegram":
            return "Puoi usare {nome} in tutti i campi."
        if not self._template():
            return ("Un messaggio libero arriva solo a chi ha scritto al numero aziendale nelle ultime 24 ore "
                    "(regola di Meta): per scrivere ai clienti serve un template.")
        if self.dati_template is None:
            return ("Valori: uno per riga, per {{1}}, {{2}}... Con \"Scarica elenco\" vedi il messaggio del "
                    "template e un campo per ogni valore.")
        testo = "I valori sono evidenziati in giallo nel messaggio. Puoi usare {nome}."
        if self.dati_template["stato"] != "APPROVED":
            testo += "  Attenzione: il template non è ancora approvato (stato: %s)." % self.dati_template["stato"]
        return testo

    def _cerca_template(self):
        """I dati del template scelto, se è nell'elenco scaricato da Meta."""
        nome, lingua = self.nome_template.get().strip(), self.lingua.get().strip()
        for dati in self.elenco_template:
            if dati["nome"] == nome and (not lingua or dati["lingua"] == lingua):
                return dati
        return None

    def _template_scelto(self):
        # Nell'elenco i template sono "nome (lingua)"
        testo = self.nome_template.get()
        if testo.endswith(")") and " (" in testo:
            nome, lingua = testo[:-1].rsplit(" (", 1)
            self.lingua.set(lingua)
            self.nome_template.set(nome)
        self.ricostruisci()

    def _modo_cambiato(self):
        # La prima volta che si sceglie "Template" l'elenco si scarica da solo (se WhatsApp è configurato)
        if self.modo.get() == "template" and not self.elenco_template and \
                self.editor.config_servizio.get("whatsapp_token"):
            self.scarica_template(avvisa=False)
        self.ricostruisci()

    def whatsapp(self):
        """Il collegamento alle API di Meta e l'id dell'account WhatsApp Business."""
        from whatsapp_meta import WhatsApp
        config = self.editor.config_servizio
        messenger = WhatsApp(config.get("whatsapp_token", ""), config.get("whatsapp_phone_number_id", ""),
                             config.get("whatsapp_versione_api") or "v23.0")
        return messenger, config.get("whatsapp_account_business", "").strip() or messenger.account_business()

    def scarica_template(self, avvisa=True):
        finestra = self.winfo_toplevel()
        finestra.configure(cursor="watch")
        finestra.update_idletasks()
        try:
            messenger, account = self.whatsapp()
            self.elenco_template = messenger.elenco_template(account)
        except Exception as e:
            if avvisa:
                messagebox.showerror("Template WhatsApp", "Non riesco a scaricare i template da Meta:\n\n" + str(e),
                                     parent=finestra)
            return False
        finally:
            finestra.configure(cursor="")
        approvati = sum(t["stato"] == "APPROVED" for t in self.elenco_template)
        self.editor._messaggio("Template WhatsApp: %d, di cui %d approvati" % (len(self.elenco_template), approvati))
        self.ricostruisci()
        return True


CATEGORIE_TEMPLATE = {"Utility (fatture, avvisi, conferme)": "UTILITY", "Marketing (promozioni)": "MARKETING"}


class FinestraNuovoTemplate(tk.Toplevel):
    """Crea un template WhatsApp sull'account di Meta; Meta poi lo approva (di solito in pochi minuti)."""

    def __init__(self, pannello):
        super().__init__(pannello.winfo_toplevel())
        self.pannello = pannello
        self.title("Nuovo template WhatsApp")
        self.transient(pannello.winfo_toplevel())
        self.nome = tk.StringVar()
        self.lingua = tk.StringVar(value=pannello.lingua.get() or "it")
        self.categoria = tk.StringVar(value=list(CATEGORIE_TEMPLATE)[0])
        self.con_documento = tk.BooleanVar(value=True)
        self.esempio = tk.StringVar()
        self.piede = tk.StringVar()

        corpo = ttk.Frame(self, padding=10)
        corpo.pack(fill="both", expand=True)
        corpo.columnconfigure(1, weight=1)
        ttk.Label(corpo, text="Nome").grid(row=0, column=0, sticky="w")
        ttk.Entry(corpo, textvariable=self.nome, width=40).grid(row=0, column=1, sticky="we")
        ttk.Label(corpo, text="solo lettere minuscole, cifre e _ (es. invio_fattura)", foreground="gray").grid(
            row=1, column=1, sticky="w")
        ttk.Label(corpo, text="Lingua").grid(row=2, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(corpo, textvariable=self.lingua, width=8).grid(row=2, column=1, sticky="w", pady=(6, 0))
        ttk.Label(corpo, text="Categoria").grid(row=3, column=0, sticky="w", pady=(6, 0))
        ttk.Combobox(corpo, textvariable=self.categoria, values=list(CATEGORIE_TEMPLATE), state="readonly",
                     width=36).grid(row=3, column=1, sticky="w", pady=(6, 0))
        ttk.Checkbutton(corpo, text="Intestazione con il documento (per mandare il PDF creato)",
                        variable=self.con_documento).grid(row=4, column=1, sticky="w", pady=(6, 0))
        esempio = ttk.Frame(corpo)
        esempio.grid(row=5, column=1, sticky="we")
        ttk.Label(esempio, text="PDF di esempio per Meta").pack(side="left")
        ttk.Entry(esempio, textvariable=self.esempio).pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(esempio, text="...", width=3, command=self._scegli_esempio).pack(side="left")
        ttk.Label(corpo, text="vuoto = un PDF di esempio creato al momento", foreground="gray").grid(
            row=6, column=1, sticky="w")
        ttk.Label(corpo, text="Testo").grid(row=7, column=0, sticky="nw", pady=(6, 0))
        self.testo = tk.Text(corpo, height=6, width=60, wrap="word", undo=True)
        self.testo.grid(row=7, column=1, sticky="we", pady=(6, 0))
        self.testo.insert("1.0", "Buongiorno {{1}},\nin allegato il documento {{2}}.\nCordiali saluti")
        ttk.Label(corpo, text="{{1}}, {{2}}... sono i valori che si scrivono nei Dati dell'invio, uno per riga",
                  foreground="gray").grid(row=8, column=1, sticky="w")
        ttk.Label(corpo, text="Esempi").grid(row=9, column=0, sticky="nw", pady=(6, 0))
        self.esempi = tk.Text(corpo, height=3, width=60, undo=True)
        self.esempi.grid(row=9, column=1, sticky="we", pady=(6, 0))
        self.esempi.insert("1.0", "Mario Rossi\nFT123")
        ttk.Label(corpo, text="un valore di esempio per riga, uno per ogni {{n}}: Meta li usa per approvarlo",
                  foreground="gray").grid(row=10, column=1, sticky="w")
        ttk.Label(corpo, text="Piè di pagina").grid(row=11, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(corpo, textvariable=self.piede).grid(row=11, column=1, sticky="we", pady=(6, 0))

        pulsanti = ttk.Frame(self, padding=(10, 0, 10, 10))
        pulsanti.pack(fill="x")
        ttk.Button(pulsanti, text="Annulla", command=self.destroy).pack(side="right")
        ttk.Button(pulsanti, text="Invia a Meta per l'approvazione", command=self.crea).pack(side="right", padx=6)

    def _scegli_esempio(self):
        percorso = filedialog.askopenfilename(parent=self, title="PDF di esempio", filetypes=[("PDF", "*.pdf")])
        if percorso:
            self.esempio.set(percorso)

    def controlla(self):
        """I dati del template, oppure solleva ErroreModello con il problema."""
        import re
        from whatsapp_meta import numero_parametri
        nome = self.nome.get().strip()
        if not re.fullmatch(r"[a-z0-9_]{1,512}", nome):
            raise modello_gui.ErroreModello("Il nome può avere solo lettere minuscole, cifre e _ (es. invio_fattura)")
        testo = self.testo.get("1.0", "end-1c").strip()
        if not testo:
            raise modello_gui.ErroreModello("Scrivi il testo del messaggio")
        esempi = [e.strip() for e in self.esempi.get("1.0", "end-1c").split("\n") if e.strip()]
        parametri = numero_parametri(testo)
        if len(esempi) < parametri:
            raise modello_gui.ErroreModello("Il testo ha %d valori ({{1}}...{{%d}}): scrivi un esempio per ognuno"
                                            % (parametri, parametri))
        esempio = self.esempio.get().strip()
        if self.con_documento.get() and esempio and not os.path.isfile(esempio):
            raise modello_gui.ErroreModello("PDF di esempio non trovato: " + esempio)
        return dict(nome=nome, lingua=self.lingua.get().strip() or "it",
                    categoria=CATEGORIE_TEMPLATE[self.categoria.get()], testo=testo, esempi=esempi[:parametri],
                    piede=self.piede.get().strip(), documento=self.con_documento.get(), esempio=esempio)

    def crea(self):
        try:
            dati = self.controlla()
        except modello_gui.ErroreModello as e:
            messagebox.showerror("Nuovo template", str(e), parent=self)
            return
        cartella = tempfile.mkdtemp(prefix="template_whatsapp_")
        self.configure(cursor="watch")
        self.update_idletasks()
        try:
            esempio = None
            if dati["documento"]:
                esempio = dati["esempio"] or self._pdf_di_esempio(cartella)
            messenger, account = self.pannello.whatsapp()
            _, stato = messenger.crea_template(account, dati["nome"], dati["lingua"], dati["categoria"],
                                               dati["testo"], dati["esempi"], esempio, dati["piede"])
        except Exception as e:
            messagebox.showerror("Nuovo template", "Meta non ha accettato il template:\n\n" + str(e), parent=self)
            return
        finally:
            self.configure(cursor="")
            shutil.rmtree(cartella, ignore_errors=True)
        self.pannello.modo.set("template")
        self.pannello.nome_template.set(dati["nome"])
        self.pannello.lingua.set(dati["lingua"])
        self.pannello.scarica_template(avvisa=False)
        messagebox.showinfo("Nuovo template", "Template \"%s\" inviato a Meta (stato: %s).\n\nPuò essere usato "
                            "quando è approvato, di solito dopo pochi minuti: lo stato si vede con \"Scarica "
                            "elenco\"." % (dati["nome"], stato or "in revisione"), parent=self.master)
        self.destroy()

    def _pdf_di_esempio(self, cartella):
        import pymupdf
        percorso = os.path.join(cartella, "esempio.pdf")
        documento = pymupdf.open()
        documento.new_page().insert_text((72, 100), "Documento di esempio", fontsize=16)
        documento.save(percorso)
        documento.close()
        return percorso


# Come il servizio scrive il PDF: ogni riga è pdf.cell(..., 6, ...), cioè alta 6 mm
ALTEZZA_RIGA_PDF = 6
MARGINE_IMMAGINE = 10


class FinestraFiligrana(tk.Toplevel):
    """Il documento creato dal servizio, aggiornato a ogni modifica dell'editor.

    Mostra dove finisce il testo sul template: la riga del cursore è evidenziata e un clic
    sul documento porta il cursore dell'editor alla riga e colonna corrispondenti.
    Le posizioni dei caratteri si leggono da un secondo PDF creato su una pagina bianca grande
    come il template, così il testo del template non si confonde con quello del modello.
    """

    def __init__(self, editor):
        super().__init__(editor)
        self.editor = editor
        self.title("Filigrana - " + os.path.basename(editor.percorso_modello or "nuovo modello"))
        larghezza = min(int(720 * SCALA_FINESTRE), int(self.winfo_screenwidth() * 0.45))
        altezza = int(self.winfo_screenheight() * 0.85)
        # A destra dello schermo, per lasciare l'editor visibile
        self.geometry("%dx%d+%d+%d" % (larghezza, altezza, self.winfo_screenwidth() - larghezza - 20, 10))
        self.cartella = tempfile.mkdtemp(prefix="filigrana_modello_")
        self.documento = None
        self.pagina = 0
        # Per ogni pagina: {riga del testo (da 0): [(x0, x1) di ogni carattere, in punti]}
        self.caratteri = []
        self.righe_per_pagina = 0
        self.margine_alto = 0
        self.spostamento = 0
        self.cliccabile = False
        self.cursore = (0, 0)
        self.zoom = tk.StringVar(value="100%")
        self.griglia = tk.BooleanVar(value=True)

        barra = ttk.Frame(self, padding=6)
        barra.pack(side="top", fill="x")
        ttk.Button(barra, text="<", width=3, command=lambda: self._vai(-1)).pack(side="left")
        self.etichetta = ttk.Label(barra, width=14, anchor="center")
        self.etichetta.pack(side="left")
        ttk.Button(barra, text=">", width=3, command=lambda: self._vai(1)).pack(side="left")
        zoom = ttk.Combobox(barra, textvariable=self.zoom, values=["50%", "75%", "100%", "125%", "150%", "200%"],
                            width=6, state="readonly")
        zoom.pack(side="left", padx=10)
        zoom.bind("<<ComboboxSelected>>", lambda e: self._mostra())
        ttk.Checkbutton(barra, text="Numeri di riga", variable=self.griglia,
                        command=self._mostra).pack(side="left")
        ttk.Button(barra, text="Aggiorna", command=self.aggiorna).pack(side="right")
        self.stato = ttk.Label(self, foreground="gray", padding=(8, 2))
        self.stato.pack(side="bottom", fill="x")

        cornice = ttk.Frame(self)
        cornice.pack(fill="both", expand=True)
        self.tela = tk.Canvas(cornice, background="#808080", highlightthickness=0, cursor="crosshair")
        verticale = ttk.Scrollbar(cornice, orient="vertical", command=self.tela.yview)
        orizzontale = ttk.Scrollbar(self, orient="horizontal", command=self.tela.xview)
        verticale.pack(side="right", fill="y")
        orizzontale.pack(side="bottom", fill="x")
        self.tela.pack(side="left", fill="both", expand=True)
        self.tela.configure(yscrollcommand=verticale.set, xscrollcommand=orizzontale.set)
        self.tela.bind("<MouseWheel>", lambda e: self.tela.yview_scroll(-e.delta // 120, "units"))
        self.tela.bind("<Button-4>", lambda e: self.tela.yview_scroll(-1, "units"))
        self.tela.bind("<Button-5>", lambda e: self.tela.yview_scroll(1, "units"))
        self.tela.bind("<Button-1>", self._clic)
        self.bind("<F6>", lambda e: self.aggiorna())
        self.protocol("WM_DELETE_WINDOW", self._chiudi)
        # Chiudendo la finestra (o l'editor) si fa pulizia
        self.bind("<Destroy>", lambda e: self._pulisci() if e.widget is self else None)
        self.aggiorna()

    # --- Creazione del documento ----------------------------------------------

    def aggiorna(self):
        if self.editor.word.get():
            self.stato.configure(text="La filigrana c'è solo per i modelli PDF")
            return
        self.configure(cursor="watch")
        self.update_idletasks()
        try:
            self._crea()
        except Exception as e:
            testo = str(e) if isinstance(e, modello_gui.ErroreModello) else traceback.format_exc(limit=-1).strip()
            self.stato.configure(text="Non aggiornata: " + " ".join(testo.split()), foreground="red")
            return
        finally:
            self.configure(cursor="")
        self.pagina = max(0, min(self.pagina, self.documento.page_count - 1))
        self._mostra()

    def _crea(self):
        import pymupdf
        if self.documento is not None:
            self.documento.close()
            self.documento = None
        con_template = os.path.join(self.cartella, "template")
        senza_template = os.path.join(self.cartella, "posizioni")
        for cartella in (con_template, senza_template):
            shutil.rmtree(cartella, ignore_errors=True)
            os.makedirs(cartella)
        percorso = self.editor.crea_documento(con_template, avvisa=False)
        if self.editor.template.get().strip():
            posizioni = self.editor.crea_documento(senza_template, senza_template=True, avvisa=False)
        else:
            posizioni = percorso
        self.documento = pymupdf.open(percorso)
        with pymupdf.open(posizioni) as doc_posizioni:
            self._leggi_posizioni(doc_posizioni)
        rotazione = self.editor.rotazione.get()
        self.cliccabile = rotazione == "0"
        self.stato.configure(foreground="gray", text="Clic sul documento: il cursore dell'editor va lì" if
                             self.cliccabile else "Con la rotazione il clic sul documento non porta alla riga")

    def _leggi_posizioni(self, doc_posizioni):
        from fpdf import FPDF
        pagina_testo = doc_posizioni[0].rect
        # Le stesse misure del servizio (FPDF, in mm): margine in alto e fine della pagina
        foglio = FPDF(orientation="L" if pagina_testo.width > pagina_testo.height else "P")
        self.righe_per_pagina = int((foglio.page_break_trigger - foglio.t_margin) / ALTEZZA_RIGA_PDF + 1e-6)
        self.margine_alto = foglio.t_margin
        # Il testo viene sovrapposto al template allineando il fondo delle pagine
        self.spostamento = self.documento[0].rect.height - pagina_testo.height
        self.caratteri = []
        per_mm = 72 / 25.4
        for numero, pagina in enumerate(doc_posizioni):
            righe = {}
            for blocco in pagina.get_text("rawdict")["blocks"]:
                for linea in blocco.get("lines", []):
                    for pezzo in linea["spans"]:
                        for carattere in pezzo["chars"]:
                            x0, y0, x1, y1 = carattere["bbox"]
                            riga = int(((y0 + y1) / 2 / per_mm - self.margine_alto) // ALTEZZA_RIGA_PDF)
                            righe.setdefault(numero * self.righe_per_pagina + riga, []).append((x0, x1))
            self.caratteri.append({riga: sorted(x) for riga, x in righe.items()})

    # --- Visualizzazione --------------------------------------------------------

    def _scala(self):
        # Pixel per punto del PDF
        return int(self.zoom.get().rstrip("%")) / 100 * 96 / 72

    def _y_riga(self, riga_pagina):
        """Bordo alto della riga sulla pagina, in punti dall'alto del documento mostrato."""
        return (self.margine_alto + riga_pagina * ALTEZZA_RIGA_PDF) * 72 / 25.4 + self.spostamento

    def _vai(self, passo):
        if self.documento is not None:
            self.pagina = max(0, min(self.documento.page_count - 1, self.pagina + passo))
            self._mostra()

    def _mostra(self):
        if self.documento is None:
            return
        scala = self._scala()
        pix = self.documento[self.pagina].get_pixmap(dpi=int(72 * scala), alpha=False)
        self.immagine = tk.PhotoImage(data=base64.b64encode(pix.tobytes("png")).decode("ascii"), format="png")
        self.tela.delete("all")
        self.tela.create_image(MARGINE_IMMAGINE, MARGINE_IMMAGINE, anchor="nw", image=self.immagine)
        self.tela.configure(scrollregion=(0, 0, pix.width + 2 * MARGINE_IMMAGINE, pix.height + 2 * MARGINE_IMMAGINE))
        self.etichetta.configure(text="Pagina %d di %d" % (self.pagina + 1, self.documento.page_count))
        if self.griglia.get() and self.cliccabile:
            prima = self.pagina * self.righe_per_pagina
            for n in range(self.righe_per_pagina):
                if prima + n >= len(self.editor.righe):
                    break
                y = MARGINE_IMMAGINE + self._y_riga(n) * scala
                self.tela.create_line(MARGINE_IMMAGINE, y, MARGINE_IMMAGINE + pix.width, y, fill="#9fc5ff",
                                      dash=(2, 4))
                self.tela.create_text(MARGINE_IMMAGINE + 3, y + 1, anchor="nw", text=str(prima + n + 1),
                                      fill="#3d7bd9", font=("Arial", max(6, int(6 * scala))))
        self._disegna_cursore()

    def mostra_cursore(self, riga, colonna):
        """La riga del cursore dell'editor: va alla sua pagina e la evidenzia."""
        self.cursore = (riga, colonna)
        if self.documento is None or not self.cliccabile or not self.righe_per_pagina:
            return
        pagina = min(riga // self.righe_per_pagina, self.documento.page_count - 1)
        if pagina != self.pagina:
            self.pagina = pagina
            self._mostra()
        else:
            self._disegna_cursore()

    def _disegna_cursore(self):
        self.tela.delete("cursore")
        riga, colonna = self.cursore
        if self.documento is None or not self.cliccabile or riga // self.righe_per_pagina != self.pagina:
            return
        scala = self._scala()
        alto = MARGINE_IMMAGINE + self._y_riga(riga % self.righe_per_pagina) * scala
        basso = alto + ALTEZZA_RIGA_PDF * 72 / 25.4 * scala
        larghezza = self.documento[self.pagina].rect.width * scala
        self.tela.create_rectangle(MARGINE_IMMAGINE, alto, MARGINE_IMMAGINE + larghezza, basso, outline="#ff8c00",
                                   width=2, tags="cursore")
        x = self._x_colonna(riga, colonna)
        if x is not None:
            x = MARGINE_IMMAGINE + x * scala
            self.tela.create_line(x, alto - 2, x, basso + 2, fill="#e00000", width=2, tags="cursore")
        # La riga del cursore resta visibile
        altezza_tela = float(self.tela.cget("scrollregion").split()[3])
        vista_alto, vista_basso = self.tela.yview()
        if not (vista_alto * altezza_tela <= alto and basso <= vista_basso * altezza_tela):
            self.tela.yview_moveto(max(0.0, (alto - 100) / altezza_tela))

    def _x_colonna(self, riga, colonna):
        """Posizione orizzontale (in punti) prima del carattere della colonna, se la riga ha testo."""
        caratteri = self.caratteri[self.pagina].get(riga) if self.pagina < len(self.caratteri) else None
        if not caratteri:
            return None
        if colonna < len(caratteri):
            return caratteri[colonna][0]
        return caratteri[-1][1]

    def _clic(self, evento):
        if self.documento is None or not self.cliccabile:
            return
        scala = self._scala()
        x = (self.tela.canvasx(evento.x) - MARGINE_IMMAGINE) / scala
        y = (self.tela.canvasy(evento.y) - MARGINE_IMMAGINE) / scala
        riga_pagina = int((y - self.spostamento) / (72 / 25.4) - self.margine_alto) // ALTEZZA_RIGA_PDF
        if y < self.spostamento or not 0 <= riga_pagina < self.righe_per_pagina:
            return
        riga = self.pagina * self.righe_per_pagina + int(riga_pagina)
        if riga >= len(self.editor.righe):
            return
        caratteri = self.caratteri[self.pagina].get(riga, [])
        # La colonna del primo carattere che sta (per più di metà) a destra del clic
        colonna = sum(1 for x0, x1 in caratteri if (x0 + x1) / 2 < x)
        colonna = min(colonna, len(self.editor.righe[riga]))
        self.editor.vai_a(riga, colonna)

    def _chiudi(self):
        self.destroy()

    def _pulisci(self):
        # L'aggiornamento in attesa non serve più
        if self.editor.attesa_filigrana is not None:
            self.editor.after_cancel(self.editor.attesa_filigrana)
            self.editor.attesa_filigrana = None
        if self.documento is not None:
            self.documento.close()
            self.documento = None
        shutil.rmtree(self.cartella, ignore_errors=True)
        self.editor.finestra_filigrana = None


# Misure della pagina A4 in punti, per l'anteprima quando il PDF non c'è
PAGINA_A4 = (595, 842)


class FinestraBarcode(tk.Toplevel):
    """I .gui "barcode": barcode da mettere su un PDF che c'è già (il servizio li aggiunge appena
    trova il .gui nella cartella di lavoro, senza file di testo).

    A sinistra l'elenco dei barcode e i dati di quello scelto, a destra la pagina del PDF con i
    barcode come li mette il servizio: trascinando sulla pagina si disegna il rettangolo del
    barcode scelto, un clic su un barcode lo sceglie.
    """

    def __init__(self, editor, percorso=None):
        super().__init__(editor)
        self.editor = editor
        dimensiona(self, 1120, 760, minima=(820, 520))
        self.percorso = None
        self.barcodes = [modello_gui.barcode_vuoto()]
        self.scelto = 0
        self.modificato = False
        self.caricando = False
        self.attesa = None
        self.immagine = None
        self.inizio = None
        self.pagine = 0
        self.zoom = tk.StringVar(value="100%")
        self.campi = {nome: tk.StringVar() for nome in ("codice", "tipo", "pdf", "pagina", "x0", "y0", "x1", "y1")}
        self._crea_interfaccia()
        for variabile in self.campi.values():
            variabile.trace_add("write", lambda *a: self._campi_cambiati())
        self.protocol("WM_DELETE_WINDOW", self._chiudi)
        self.bind("<Control-s>", lambda e: self.salva())
        if percorso is None or not self.apri_file(percorso):
            self._mostra_scelto()
        self._aggiorna_titolo()

    # --- Interfaccia -------------------------------------------------------

    def _crea_interfaccia(self):
        barra = ttk.Frame(self, padding=6)
        barra.pack(side="top", fill="x")
        ttk.Button(barra, text="Nuovo", command=self.nuovo).pack(side="left")
        ttk.Button(barra, text="Apri...", command=self.apri).pack(side="left", padx=(4, 0))
        ttk.Button(barra, text="Salva", command=self.salva).pack(side="left", padx=(4, 0))
        ttk.Button(barra, text="Salva con nome...", command=lambda: self.salva(con_nome=True)).pack(side="left",
                                                                                                    padx=(4, 0))
        ttk.Button(barra, text="Manda al servizio", command=self.manda).pack(side="left", padx=(12, 0))
        zoom = ttk.Combobox(barra, textvariable=self.zoom, values=["50%", "75%", "100%", "125%", "150%"], width=6,
                            state="readonly")
        zoom.pack(side="right")
        zoom.bind("<<ComboboxSelected>>", lambda e: self._disegna())
        ttk.Label(barra, text="Zoom").pack(side="right", padx=(0, 4))

        sinistra = ttk.Frame(self, padding=(6, 0, 6, 6))
        sinistra.pack(side="left", fill="y")
        ttk.Label(sinistra, text="Barcode").pack(anchor="w")
        self.elenco = tk.Listbox(sinistra, height=8, width=44, activestyle="none", exportselection=False)
        self.elenco.pack(fill="x")
        self.elenco.bind("<<ListboxSelect>>", lambda e: self._scelto_da_elenco())
        pulsanti = ttk.Frame(sinistra)
        pulsanti.pack(fill="x", pady=(4, 8))
        ttk.Button(pulsanti, text="Aggiungi", command=self.aggiungi).pack(side="left")
        ttk.Button(pulsanti, text="Duplica", command=lambda: self.aggiungi(copia=True)).pack(side="left", padx=4)
        ttk.Button(pulsanti, text="Rimuovi", command=self.rimuovi).pack(side="left")

        dati = ttk.LabelFrame(sinistra, text="Barcode scelto", padding=8)
        dati.pack(fill="x")
        dati.columnconfigure(1, weight=1)
        ttk.Label(dati, text="Codice").grid(row=0, column=0, sticky="w")
        ttk.Entry(dati, textvariable=self.campi["codice"]).grid(row=0, column=1, columnspan=4, sticky="we", pady=2)
        ttk.Label(dati, text="Tipo").grid(row=1, column=0, sticky="w")
        # Si può anche scrivere un tipo che non è nell'elenco (python-barcode ne conosce altri)
        ttk.Combobox(dati, textvariable=self.campi["tipo"], values=list(modello_gui.TIPI_BARCODE), width=12).grid(
            row=1, column=1, sticky="w", pady=2)
        self.descrizione_tipo = ttk.Label(dati, foreground="gray")
        self.descrizione_tipo.grid(row=1, column=2, columnspan=3, sticky="w", padx=(6, 0))
        ttk.Label(dati, text="PDF").grid(row=2, column=0, sticky="w")
        ttk.Entry(dati, textvariable=self.campi["pdf"]).grid(row=2, column=1, columnspan=3, sticky="we", pady=2)
        ttk.Button(dati, text="Sfoglia...", command=self._scegli_pdf).grid(row=2, column=4, padx=(4, 0))
        ttk.Label(dati, text="Pagina").grid(row=3, column=0, sticky="w")
        ttk.Spinbox(dati, from_=1, to=9999, textvariable=self.campi["pagina"], width=6).grid(row=3, column=1,
                                                                                         sticky="w", pady=2)
        self.etichetta_pagine = ttk.Label(dati, foreground="gray")
        self.etichetta_pagine.grid(row=3, column=2, columnspan=3, sticky="w", padx=(6, 0))
        for riga, (testo, x, y) in enumerate((("In alto a sinistra", "x0", "y0"),
                                              ("In basso a destra", "x1", "y1"))):
            ttk.Label(dati, text=testo).grid(row=4 + riga, column=0, sticky="w")
            cornice = ttk.Frame(dati)
            cornice.grid(row=4 + riga, column=1, columnspan=4, sticky="w", pady=2)
            for nome in (x, y):
                ttk.Label(cornice, text=nome[0]).pack(side="left", padx=(0, 2))
                ttk.Spinbox(cornice, from_=0, to=5000, textvariable=self.campi[nome], width=6).pack(side="left",
                                                                                                padx=(0, 8))
        self.errore = ttk.Label(dati, foreground="#b00000", wraplength=330, justify="left")
        self.errore.grid(row=6, column=0, columnspan=5, sticky="w", pady=(6, 0))
        ttk.Label(sinistra, wraplength=340, justify="left", foreground="gray", text=(
            "Trascina sulla pagina per disegnare il rettangolo del barcode scelto; un clic su un barcode "
            "lo sceglie. Le misure sono in punti (72 per pollice, 1 mm = 2,83 punti) dall'angolo in alto "
            "a sinistra della pagina.\n\nIl servizio mette i barcode appena trova il .gui nella cartella di "
            "lavoro e salva il PDF lì: il PDF deve stare nella cartella di lavoro.")).pack(anchor="w", pady=(8, 0))
        self.stato = ttk.Label(sinistra, foreground="gray", wraplength=340, justify="left")
        self.stato.pack(anchor="w", pady=(8, 0))

        cornice = ttk.Frame(self)
        cornice.pack(side="left", fill="both", expand=True)
        self.tela = tk.Canvas(cornice, background="#808080", highlightthickness=0, cursor="crosshair")
        verticale = ttk.Scrollbar(cornice, orient="vertical", command=self.tela.yview)
        orizzontale = ttk.Scrollbar(cornice, orient="horizontal", command=self.tela.xview)
        verticale.pack(side="right", fill="y")
        orizzontale.pack(side="bottom", fill="x")
        self.tela.pack(side="left", fill="both", expand=True)
        self.tela.configure(yscrollcommand=verticale.set, xscrollcommand=orizzontale.set)
        self.tela.bind("<MouseWheel>", lambda e: self.tela.yview_scroll(-e.delta // 120, "units"))
        self.tela.bind("<ButtonPress-1>", self._premi)
        self.tela.bind("<B1-Motion>", self._trascina)
        self.tela.bind("<ButtonRelease-1>", self._rilascia)

    def _aggiorna_titolo(self):
        nome = os.path.basename(self.percorso) if self.percorso else "nuovo"
        self.title("Barcode su un PDF - " + nome + (" *" if self.modificato else ""))

    def _messaggio(self, testo):
        self.stato.configure(text=testo)

    # --- Dati --------------------------------------------------------------

    def errore_barcode(self, b):
        """Cosa non va in un barcode ("" se è tutto a posto)."""
        if not b["codice"].strip():
            return "Manca il codice"
        try:
            modello_gui.immagine_barcode(b["codice"], b["tipo"])
        except Exception as e:
            return "Il codice non va bene per il tipo %s: %s" % (b["tipo"] or "?", e)
        if b["a"][0] <= b["da"][0] or b["a"][1] <= b["da"][1]:
            return "Il rettangolo è vuoto: l'angolo in basso a destra deve stare a destra e sotto l'altro"
        if not b["pdf"].strip():
            return "Manca il PDF"
        return ""

    def _riga_elenco(self, n, b):
        return "%s%d. %s %s (pag. %d)" % ("⚠ " if self.errore_barcode(b) else "", n + 1, b["tipo"],
                                         b["codice"] or "?", b["pagina"] + 1)

    def _aggiorna_elenco(self):
        self.elenco.delete(0, "end")
        for n, b in enumerate(self.barcodes):
            self.elenco.insert("end", self._riga_elenco(n, b))
        self.elenco.selection_set(self.scelto)
        self.elenco.see(self.scelto)

    def _mostra_scelto(self):
        """Mette nei campi i dati del barcode scelto e ridisegna."""
        b = self.barcodes[self.scelto]
        self.caricando = True
        try:
            valori = {"codice": b["codice"], "tipo": b["tipo"], "pdf": b["pdf"], "pagina": str(b["pagina"] + 1),
                      "x0": str(b["da"][0]), "y0": str(b["da"][1]), "x1": str(b["a"][0]), "y1": str(b["a"][1])}
            for nome, valore in valori.items():
                self.campi[nome].set(valore)
        finally:
            self.caricando = False
        self._aggiorna_elenco()
        self._aggiorna_errore()
        self._disegna()

    def _campi_cambiati(self):
        if self.caricando:
            return
        b = self.barcodes[self.scelto]
        b["codice"] = self.campi["codice"].get()
        b["tipo"] = self.campi["tipo"].get().strip()
        b["pdf"] = self.campi["pdf"].get()
        # I numeri scritti a metà (es. un campo vuoto) restano quelli di prima finché non sono validi
        try:
            b["pagina"] = max(0, int(self.campi["pagina"].get()) - 1)
        except ValueError:
            pass
        try:
            b["da"] = (int(self.campi["x0"].get()), int(self.campi["y0"].get()))
            b["a"] = (int(self.campi["x1"].get()), int(self.campi["y1"].get()))
        except ValueError:
            pass
        self.modificato = True
        self._aggiorna_titolo()
        self.elenco.delete(self.scelto)
        self.elenco.insert(self.scelto, self._riga_elenco(self.scelto, b))
        self.elenco.selection_set(self.scelto)
        self._aggiorna_errore()
        # Si ridisegna quando si smette per un attimo di scrivere
        if self.attesa is not None:
            self.after_cancel(self.attesa)
        self.attesa = self.after(300, self._disegna)

    def _aggiorna_errore(self):
        b = self.barcodes[self.scelto]
        self.descrizione_tipo.configure(text=modello_gui.TIPI_BARCODE.get(b["tipo"], ""))
        self.errore.configure(text=self.errore_barcode(b))

    def _scelto_da_elenco(self):
        selezione = self.elenco.curselection()
        if selezione and selezione[0] != self.scelto:
            self.scelto = selezione[0]
            self._mostra_scelto()

    def aggiungi(self, copia=False):
        attuale = self.barcodes[self.scelto]
        if copia:
            nuovo = dict(attuale)
        else:
            # Sullo stesso PDF e pagina, un po' più in basso
            nuovo = dict(modello_gui.barcode_vuoto(), pdf=attuale["pdf"], pagina=attuale["pagina"])
        altezza = attuale["a"][1] - attuale["da"][1] + 10
        nuovo["da"] = (attuale["da"][0], attuale["da"][1] + altezza)
        nuovo["a"] = (attuale["a"][0], attuale["a"][1] + altezza)
        self.barcodes.insert(self.scelto + 1, nuovo)
        self.scelto += 1
        self.modificato = True
        self._aggiorna_titolo()
        self._mostra_scelto()

    def rimuovi(self):
        if len(self.barcodes) == 1:
            self.barcodes = [dict(modello_gui.barcode_vuoto(), pdf=self.barcodes[0]["pdf"])]
        else:
            del self.barcodes[self.scelto]
            self.scelto = min(self.scelto, len(self.barcodes) - 1)
        self.modificato = True
        self._aggiorna_titolo()
        self._mostra_scelto()

    def _scegli_pdf(self):
        attuale = self.campi["pdf"].get().strip()
        cartella = os.path.dirname(attuale) if attuale else self.editor.config_servizio.get("cartella_lavoro", "")
        percorso = filedialog.askopenfilename(parent=self, title="PDF su cui mettere il barcode",
                                              initialdir=cartella if os.path.isdir(cartella) else None,
                                              filetypes=[("PDF", "*.pdf"), ("Tutti i file", "*.*")])
        if percorso:
            self.campi["pdf"].set(os.path.normpath(percorso))

    # --- File --------------------------------------------------------------

    def _conferma_perdita(self):
        if not self.modificato:
            return True
        risposta = messagebox.askyesnocancel("Barcode", "Ci sono modifiche non salvate. Salvarle?", parent=self)
        if risposta is None:
            return False
        return self.salva() if risposta else True

    def nuovo(self):
        if not self._conferma_perdita():
            return
        self.percorso = None
        self.barcodes = [modello_gui.barcode_vuoto()]
        self.scelto = 0
        self.modificato = False
        self._mostra_scelto()
        self._aggiorna_titolo()

    def apri(self):
        if not self._conferma_perdita():
            return
        os.makedirs(CARTELLA_MODELLI, exist_ok=True)
        percorso = filedialog.askopenfilename(parent=self, title="Apri .gui barcode", initialdir=CARTELLA_MODELLI,
                                              filetypes=[("Barcode", "*.gui"), ("Tutti i file", "*.*")])
        if percorso:
            self.apri_file(percorso)

    def apri_file(self, percorso):
        try:
            _, barcodes = modello_gui.leggi_barcode(percorso)
        except (OSError, modello_gui.ErroreModello) as e:
            messagebox.showerror("Apri", "Impossibile aprire %s:\n%s" % (percorso, e), parent=self)
            return False
        self.barcodes = barcodes or [modello_gui.barcode_vuoto()]
        self.percorso = percorso
        self.scelto = 0
        self.modificato = False
        self._mostra_scelto()
        self._aggiorna_titolo()
        self._messaggio("Aperto " + percorso)
        return True

    def _controlla(self, titolo):
        """Prima di salvare: True se è tutto a posto o se si vuole salvare lo stesso."""
        problemi = ["%d. %s" % (n + 1, e) for n, e in enumerate(map(self.errore_barcode, self.barcodes)) if e]
        lavoro = self.editor.config_servizio.get("cartella_lavoro", "")
        if lavoro:
            fuori = sorted({b["pdf"] for b in self.barcodes if b["pdf"].strip() and os.path.normcase(
                os.path.dirname(os.path.abspath(b["pdf"].strip()))) != os.path.normcase(os.path.abspath(lavoro))})
            problemi += ["Il PDF non è nella cartella di lavoro (%s): il servizio non potrà salvarlo: %s"
                         % (lavoro, pdf) for pdf in fuori]
        if not problemi:
            return True
        return messagebox.askokcancel(titolo, "Attenzione:\n\n- " + "\n- ".join(problemi) + "\n\nSalvare comunque?",
                                      parent=self)

    def salva(self, con_nome=False):
        percorso = self.percorso
        if con_nome or percorso is None:
            os.makedirs(CARTELLA_MODELLI, exist_ok=True)
            percorso = filedialog.asksaveasfilename(parent=self, title="Salva .gui barcode",
                                                    initialdir=os.path.dirname(percorso) if percorso else
                                                    CARTELLA_MODELLI,
                                                    initialfile=os.path.basename(percorso) if percorso else
                                                    "barcode.gui", defaultextension=".gui",
                                                    filetypes=[("Barcode", "*.gui")])
            if not percorso:
                return False
        if not self._controlla("Salva"):
            return False
        modello_gui.scrivi_barcode(percorso, os.path.splitext(os.path.basename(percorso))[0], self.barcodes)
        self.percorso = percorso
        self.modificato = False
        self._aggiorna_titolo()
        self._messaggio("Salvato " + percorso)
        return True

    def manda(self):
        """Scrive il .gui nella cartella di lavoro: il servizio mette subito i barcode nel PDF."""
        lavoro = self.editor.config_servizio.get("cartella_lavoro", "")
        if not os.path.isdir(lavoro):
            messagebox.showerror("Manda al servizio", "Cartella di lavoro non trovata: imposta la configurazione "
                                 "con ConfigManager.py", parent=self)
            return
        nome = os.path.splitext(os.path.basename(self.percorso))[0] if self.percorso else "barcode"
        destinazione = os.path.join(lavoro, nome + ".gui")
        if not self._controlla("Manda al servizio") or not messagebox.askokcancel(
                "Manda al servizio", "Verrà creato %s\ne il servizio metterà i barcode nel PDF (il PDF viene "
                                     "modificato). Continuare?" % destinazione, parent=self):
            return
        modello_gui.scrivi_barcode(destinazione, nome, self.barcodes)
        self._messaggio("Creato " + destinazione)

    def _chiudi(self):
        if self._conferma_perdita():
            self.destroy()

    # --- Pagina ------------------------------------------------------------

    def _scala(self):
        return int(self.zoom.get().rstrip("%")) / 100

    def _stessa_pagina(self, b, scelto):
        return os.path.normcase(b["pdf"].strip()) == os.path.normcase(scelto["pdf"].strip()) and \
            b["pagina"] == scelto["pagina"]

    def _disegna(self):
        """La pagina del barcode scelto con i barcode che ci vanno, come li mette il servizio."""
        import pymupdf
        self.attesa = None
        if not self.winfo_exists():
            return
        scelto = self.barcodes[self.scelto]
        percorso = scelto["pdf"].strip()
        documento = None
        if percorso and os.path.isfile(percorso):
            try:
                documento = pymupdf.open(percorso)
            except Exception as e:
                self._messaggio("Impossibile aprire il PDF: %s" % e)
        if documento is not None and documento.page_count:
            self.pagine = documento.page_count
            indice = min(scelto["pagina"], self.pagine - 1)
            fuori = scelto["pagina"] >= self.pagine
            self.etichetta_pagine.configure(text=("di %d" % self.pagine) + ("  (la pagina non c'è!)" if fuori else ""),
                                            foreground="#b00000" if fuori else "gray")
        else:
            if documento is not None:
                documento.close()
            documento = pymupdf.open()
            documento.new_page(width=PAGINA_A4[0], height=PAGINA_A4[1])
            indice = 0
            self.etichetta_pagine.configure(text="PDF non trovato: pagina A4 bianca" if percorso else "",
                                            foreground="gray")
        try:
            pagina = documento[indice]
            for b in self.barcodes:
                if self._stessa_pagina(b, scelto) and not self.errore_barcode(b):
                    pagina.insert_image(pymupdf.Rect(*b["da"], *b["a"]),
                                        stream=modello_gui.immagine_barcode(b["codice"], b["tipo"]))
            scala = self._scala()
            pix = pagina.get_pixmap(matrix=pymupdf.Matrix(scala, scala), alpha=False)
        finally:
            documento.close()
        self.immagine = tk.PhotoImage(data=base64.b64encode(pix.tobytes("png")).decode("ascii"), format="png")
        self.tela.delete("all")
        self.tela.create_image(MARGINE_IMMAGINE, MARGINE_IMMAGINE, anchor="nw", image=self.immagine)
        self.tela.configure(scrollregion=(0, 0, pix.width + 2 * MARGINE_IMMAGINE, pix.height + 2 * MARGINE_IMMAGINE))
        for n, b in enumerate(self.barcodes):
            if not self._stessa_pagina(b, scelto):
                continue
            x0, y0 = self._da_punti(*b["da"])
            x1, y1 = self._da_punti(*b["a"])
            colore = "#ff8c00" if n == self.scelto else "#3a7bd5"
            self.tela.create_rectangle(x0, y0, x1, y1, outline=colore, width=2, dash=() if n == self.scelto else (4, 2))
            self.tela.create_text(x0 + 2, y0 - 1, anchor="sw", text=str(n + 1), fill=colore,
                                  font=("Segoe UI", 9, "bold"))

    def _da_punti(self, x, y):
        scala = self._scala()
        return MARGINE_IMMAGINE + x * scala, MARGINE_IMMAGINE + y * scala

    def _in_punti(self, evento):
        scala = self._scala()
        return (max(0, round((self.tela.canvasx(evento.x) - MARGINE_IMMAGINE) / scala)),
                max(0, round((self.tela.canvasy(evento.y) - MARGINE_IMMAGINE) / scala)))

    def _premi(self, evento):
        self.inizio = self._in_punti(evento)

    def _trascina(self, evento):
        if self.inizio is None:
            return
        fine = self._in_punti(evento)
        self.tela.delete("nuovo")
        self.tela.create_rectangle(*self._da_punti(*self.inizio), *self._da_punti(*fine), outline="#ff8c00", width=2,
                                   tags="nuovo")

    def _rilascia(self, evento):
        if self.inizio is None:
            return
        (xa, ya), (xb, yb) = self.inizio, self._in_punti(evento)
        self.inizio = None
        self.tela.delete("nuovo")
        if abs(xb - xa) < 5 or abs(yb - ya) < 5:
            # Un clic: sceglie il barcode che c'è in quel punto (l'ultimo, se sono sovrapposti)
            scelto = self.barcodes[self.scelto]
            for n in reversed(range(len(self.barcodes))):
                b = self.barcodes[n]
                if self._stessa_pagina(b, scelto) and b["da"][0] <= xb <= b["a"][0] and b["da"][1] <= yb <= b["a"][1]:
                    if n != self.scelto:
                        self.scelto = n
                        self._mostra_scelto()
                    return
            return
        self.caricando = True
        try:
            for nome, valore in (("x0", min(xa, xb)), ("y0", min(ya, yb)), ("x1", max(xa, xb)), ("y1", max(ya, yb))):
                self.campi[nome].set(str(valore))
        finally:
            self.caricando = False
        self._campi_cambiati()


if __name__ == "__main__":
    app = avvia_finestra(EditorModelli)
    if len(sys.argv) > 1:
        # Un modello che non si apre non chiude l'editor: l'errore si vede in una finestra
        app.after_idle(app.apri_modello, sys.argv[1])
    app.mainloop()
