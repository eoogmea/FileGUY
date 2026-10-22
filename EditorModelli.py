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
from piattaforma import apri_file, cartella_utente, console_sicura, leggi_testo, registra_font

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


def dimensiona(finestra, larghezza, altezza, minima=None):
    """Dimensione iniziale della finestra: quella indicata, ma al massimo l'85% dello schermo
    (sui portatili e con lo zoom di Windows lo schermo è piccolo). La finestra va al centro,
    un po' in alto per lasciare posto alla barra delle applicazioni."""
    schermo_x, schermo_y = finestra.winfo_screenwidth(), finestra.winfo_screenheight()
    larghezza = min(larghezza, int(schermo_x * 0.85))
    altezza = min(altezza, int(schermo_y * 0.85))
    finestra.geometry("%dx%d+%d+%d" % (larghezza, altezza, max(0, (schermo_x - larghezza) // 2),
                                       max(0, (schermo_y - altezza) // 3)))
    if minima:
        finestra.minsize(min(minima[0], larghezza), min(minima[1], altezza))


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
        dimensiona(self, 1060, 680, minima=(820, 460))
        try:
            self.config_servizio = carica_config()
        except Exception:
            self.config_servizio = {"cartella_lavoro": "", "cartella_template": "", "cartella_output": ""}
        self.percorso_modello = None
        self.percorso_testo = None
        self.righe = []
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
        file.add_command(label="Apri modello...", command=self.apri, accelerator="Ctrl+O")
        self.menu_recenti = tk.Menu(file, tearoff=False, postcommand=self._elenca_recenti)
        file.add_cascade(label="Modelli recenti", menu=self.menu_recenti)
        file.add_separator()
        file.add_command(label="Salva modello", command=self.salva, accelerator="Ctrl+S")
        file.add_command(label="Salva modello con nome...", command=lambda: self.salva(con_nome=True))
        file.add_separator()
        file.add_command(label="Applica a file di testo...", command=self.applica)
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
        ttk.Entry(riquadro, textvariable=self.uscita, width=24).grid(row=1, column=1, columnspan=3, sticky="we",
                                                                     pady=(4, 0))
        ttk.Label(riquadro, text="Invio").grid(row=1, column=4, sticky="w", padx=(16, 4), pady=(4, 0))
        SceltaDescritta(riquadro, self.invio, modello_gui.DESCRIZIONI_INVII, width=12).grid(row=1, column=5, sticky="w", pady=(4, 0))
        ttk.Label(riquadro, text="Destinatario").grid(row=1, column=6, sticky="w", padx=(16, 4), pady=(4, 0))
        ttk.Entry(riquadro, textvariable=self.destinatario, width=28).grid(row=1, column=7, columnspan=2,
                                                                           sticky="we", pady=(4, 0))

        ttk.Label(riquadro, text="Dati dell'invio\n(righe dopo \"fine\")").grid(row=2, column=0, sticky="nw",
                                                                             pady=(4, 0))
        self.coda = tk.Text(riquadro, height=5, width=60, font=(famiglia_tk("o"), 9), undo=True)
        self.coda.grid(row=2, column=1, columnspan=8, sticky="we", pady=(4, 0))
        self.coda.bind("<<Modified>>", self._coda_modificata)
        self.aiuto_coda = ttk.Label(riquadro, foreground="gray")
        self.aiuto_coda.grid(row=3, column=1, columnspan=8, sticky="w")
        self.invio.trace_add("write", lambda *a: self._aggiorna_aiuto_coda())
        self._aggiorna_aiuto_coda()
        riquadro.columnconfigure(8, weight=1)

    def _aggiorna_aiuto_coda(self):
        aiuti = {
            "email": "Destinatario: l'indirizzo email.  Dati dell'invio: l'oggetto, eventuali righe "
                     "\"allegato: <percorso>\", una riga vuota e il testo",
            "telegram": "Destinatario: telegram <chat>.  Dati dell'invio: il documento da mandare e il testo",
            "whatsapp": "Destinatario: whatsapp <numero>.  Dati dell'invio: il documento da mandare e il testo",
        }
        testo = aiuti.get(self.invio.get(), "")
        self.aiuto_coda.configure(text=(testo + "   -   " if testo else "") +
                                  "{nome} = nome del file di testo, senza .txt")

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
        azioni = modello_gui.AZIONI_WORD if word else modello_gui.AZIONI_PDF
        self.scelta_azione.imposta_scelte({a: modello_gui.DESCRIZIONI_AZIONI[a] for a in azioni})
        if self.azione.get() not in self.scelta_azione.codici():
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
        self._imposta_intestazione(modello_gui.nuovo_modello(self.word.get()))
        self._mostra_testo(self._leggi_testo(percorso))
        self.modificato = False
        self._aggiorna_titolo()
        return True

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
        email_aggiornata = modello_gui.aggiorna_email(intestazione)
        self._imposta_intestazione(intestazione)
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
        intestazione = self._intestazione()
        # Nella prima riga il testo con lo stesso nome del modello (applicandolo diventa quello del file)
        intestazione["testo"] = os.path.splitext(os.path.basename(percorso))[0] + ".txt"
        modello_gui.scrivi_gui(percorso, intestazione, linee)
        # Copia del testo di esempio accanto al modello, per poterlo riaprire
        copia_testo = os.path.splitext(percorso)[0] + ".txt"
        if self.testo_cambiato:
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
        if self.righe:
            self.modificato = True
            self._aggiorna_titolo()
            self._programma_filigrana()

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
        larghezza = min(720, int(self.winfo_screenwidth() * 0.45))
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


if __name__ == "__main__":
    app = EditorModelli()
    if len(sys.argv) > 1:
        app.apri_modello(sys.argv[1])
    app.mainloop()
