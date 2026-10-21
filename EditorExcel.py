"""Editor dei modelli .gui per i file Excel (azione "excel" del servizio).

Avvio: python EditorExcel.py

1. File > Nuovo da testo di esempio: si apre un file di testo con i campi separati da ";"
   come quelli che arrivano al servizio.
2. Scheda Colonne: si sceglie una colonna (clic sulla lettera o su una cella; Ctrl+clic e
   Maiusc+clic per più colonne) e se ne decide il formato dei valori (numero, valuta, data,
   ora...), il carattere (font, dimensione, grassetto, corsivo, colore), la larghezza e
   l'allineamento.
3. Scheda Righe speciali: titoli, intestazione, totali... Ogni stile ha le sue righe (es. 1-2, u1
   = ultima riga; clic sui numeri di riga per sceglierle), carattere, sfondo, allineamento e la
   scelta di lasciare i valori come testo. Vale più della formattazione della colonna.
   La tabella mostra subito le celle come le scriverà il servizio.
4. Anteprima (F5): crea il file Excel con lo stesso codice del servizio e lo apre.
5. File > Salva modello / Applica a file di testo: come nell'editor dei modelli PDF e Word.
Il formato delle linee delle colonne nel .gui è descritto in modello_gui.py e nella guida
(Guida_EditorModelli.html, menu Aiuto, F1).
"""
import importlib
import os
import shutil
import sys
import tempfile
import tkinter as tk
import traceback
from decimal import ROUND_HALF_UP, Decimal
from tkinter import filedialog, messagebox, ttk

import modello_gui
from EditorModelli import (CARTELLA_MODELLI, DIMENSIONI, MAX_ANNULLA, MAX_RECENTI, NOMI_COLORI, NOMI_FONT,
                           SceltaDescritta, apri_con_avviso, apri_guida, dimensiona, leggi_preferenze,
                           scrivi_preferenze)
from config_sicuro import carica_config
from modello_gui import (COLORI, FONT, FORMATI_EXCEL, PREDEFINITO_EXCEL, SFONDI, campi_excel,
                         colonna_excel_vuota, righe_excel, stile_riga_vuoto, valore_excel)
from piattaforma import console_sicura, leggi_testo

# Senza console (exe dell'editor) i messaggi su stdout/stderr farebbero fallire l'operazione
console_sicura()

# Righe del testo mostrate nella tabella (l'anteprima usa tutto il testo)
RIGHE_MOSTRATE = 60
NOMI_FORMATI = ["%d - %s" % (codice, descrizione) for codice, (descrizione, _) in FORMATI_EXCEL.items()]
NOMI_ALLINEAMENTI = ["automatico", "sinistra", "centro", "destra"]
MESI = ["gen", "feb", "mar", "apr", "mag", "giu", "lug", "ago", "set", "ott", "nov", "dic"]
# Formati che trasformano il testo: se un valore non si converte, in Excel resta testo
CODICI_CONVERTITI = set(modello_gui.CODICI_NUMERO + modello_gui.CODICI_DATA + modello_gui.CODICI_ORA)

NOMI_SFONDI = ["nessuno"] + [descrizione for descrizione, _ in SFONDI.values()]
# Stile delle righe speciali proposto in un modello nuovo: l'intestazione in grassetto, come testo
STILE_INTESTAZIONE = dict(stile_riga_vuoto(), righe="1", nome="Intestazione", carattere=("k", True, False, 11, "n"),
                          testo=True)
# Le righe valide si controllano con un foglio "infinito"
RIGHE_MASSIME = 10 ** 6

SFONDO = "white"
SFONDO_SELEZIONE = "#dcebfb"
SFONDO_LETTERE = "#ececec"
SFONDO_LETTERE_SELEZIONE = "#b9d6f5"
# Numero di una riga che ha uno stile speciale
SFONDO_RIGA_SPECIALE = "#e2d9f3"
# Bordo delle celle delle righe scelte nella scheda "Righe speciali"
BORDO_RIGA_SCELTA = "#2f6fc4"


def intervalli(numeri):
    """{1, 2, 3, 7} -> "1-3, 7" (il contrario di modello_gui.righe_excel)."""
    parti = []
    for numero in sorted(numeri):
        if parti and parti[-1][1] == numero - 1:
            parti[-1][1] = numero
        else:
            parti.append([numero, numero])
    return ", ".join(str(a) if a == b else "%d-%d" % (a, b) for a, b in parti)


def lettera_colonna(n):
    """0 -> A, 25 -> Z, 26 -> AA (come in Excel)."""
    lettere = ""
    n += 1
    while n > 0:
        n, resto = divmod(n - 1, 26)
        lettere = chr(65 + resto) + lettere
    return lettere


def numero_italiano(valore, decimali, migliaia):
    # Arrotondato come Excel (0,5 per eccesso); Python arrotonderebbe al pari (1234,5 -> 1234)
    arrotondato = Decimal(repr(valore)).quantize(Decimal(1).scaleb(-decimali), rounding=ROUND_HALF_UP)
    testo = (("{:,.%df}" if migliaia else "{:.%df}") % decimali).format(arrotondato)
    return testo.replace(",", "#").replace(".", ",").replace("#", ".")


def testo_mostrato(testo, codice):
    """Come Excel mostrerà la cella (con le impostazioni italiane); il valore è convertito come nel servizio."""
    valore = valore_excel(testo, codice)
    if isinstance(valore, str):
        return valore
    if isinstance(valore, float):
        if codice in (1, 2, 3, 4):
            return numero_italiano(valore, 2 if codice in (2, 4) else 0, codice in (3, 4))
        if codice in (5, 6, 7, 8, 24):
            return numero_italiano(valore, 2 if codice in (8, 24) else 0, True) + " €"
        if codice in (9, 10):
            return numero_italiano(valore * 100, 2 if codice == 10 else 0, False) + "%"
        if codice == 21:
            return ("%.1E" % valore).replace(".", ",")
        return ("%.2E" % valore).replace(".", ",")
    if hasattr(valore, "year"):
        mese = MESI[valore.month - 1]
        if codice == 13:
            return "%02d-%s-%d" % (valore.day, mese, valore.year)
        if codice == 14:
            return "%02d-%s" % (valore.day, mese)
        if codice == 15:
            return "%s %d" % (mese, valore.year)
        if codice == 20:
            return valore.strftime("%d/%m/%Y %H:%M")
        return valore.strftime("%d/%m/%Y")
    # Ora
    if codice in (16, 17):
        ora = valore.hour % 12 or 12
        secondi = ":%02d" % valore.second if codice == 17 else ""
        return "%d:%02d%s %s" % (ora, valore.minute, secondi, "AM" if valore.hour < 12 else "PM")
    if codice == 19:
        return valore.strftime("%H:%M:%S")
    if codice == 22:
        return valore.strftime("%M:%S")
    return valore.strftime("%H:%M")


def font_tk(formato, zoom=1.0):
    font, grassetto, corsivo, dimensione, _ = formato or PREDEFINITO_EXCEL
    stile = " ".join(n for n, attivo in (("bold", grassetto), ("italic", corsivo)) if attivo) or "normal"
    return (FONT[font][1], max(1, int(round(dimensione * zoom))), stile)


class GruppoCarattere(ttk.LabelFrame):
    """Font, dimensione, grassetto, corsivo e colore (di Excel predefinito: Calibri 11 nero)."""

    def __init__(self, genitore, titolo, cambiato):
        super().__init__(genitore, text=titolo, padding=(8, 4))
        self.cambiato = cambiato
        self.font = ttk.Combobox(self, values=list(NOMI_FONT), width=11, state="readonly")
        self.dimensione = ttk.Combobox(self, values=DIMENSIONI, width=4)
        self.grassetto = tk.BooleanVar()
        self.corsivo = tk.BooleanVar()
        self.colore = ttk.Combobox(self, values=list(NOMI_COLORI), width=8, state="readonly")
        ttk.Label(self, text="Carattere").grid(row=0, column=0, sticky="w")
        self.font.grid(row=0, column=1, sticky="w", padx=(4, 10))
        ttk.Label(self, text="Dimensione").grid(row=0, column=2, sticky="w")
        self.dimensione.grid(row=0, column=3, sticky="w", padx=(4, 0))
        ttk.Checkbutton(self, text="Grassetto", variable=self.grassetto, command=cambiato).grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))
        ttk.Checkbutton(self, text="Corsivo", variable=self.corsivo, command=cambiato).grid(
            row=1, column=1, sticky="e", padx=(0, 10), pady=(4, 0))
        ttk.Label(self, text="Colore").grid(row=1, column=2, sticky="w", pady=(4, 0))
        self.colore.grid(row=1, column=3, sticky="w", padx=(4, 0), pady=(4, 0))
        for voce in (self.font, self.colore, self.dimensione):
            voce.bind("<<ComboboxSelected>>", lambda e: cambiato())
        self.dimensione.bind("<Return>", lambda e: cambiato())
        self.dimensione.bind("<FocusOut>", lambda e: cambiato())

    def imposta(self, formato):
        font, grassetto, corsivo, dimensione, colore = formato or PREDEFINITO_EXCEL
        self.font.set(FONT[font][0])
        self.dimensione.set(str(dimensione))
        self.grassetto.set(grassetto)
        self.corsivo.set(corsivo)
        self.colore.set(COLORI[colore][0])

    def valore(self, predefinito_vuoto=True):
        """Il formato scelto; con predefinito_vuoto None se è quello di Excel (Calibri 11 nero).
        ValueError se la dimensione è sbagliata."""
        dimensione = self.dimensione.get().strip()
        if not dimensione.isdigit() or not 1 <= int(dimensione) <= 99:
            raise ValueError("Dimensione non valida: deve essere tra 1 e 99")
        formato = (NOMI_FONT[self.font.get()], self.grassetto.get(), self.corsivo.get(), int(dimensione),
                   NOMI_COLORI[self.colore.get()])
        return None if predefinito_vuoto and formato == PREDEFINITO_EXCEL else formato

    def abilita(self, attivo):
        for figlio in self.winfo_children():
            figlio.state(["!disabled"] if attivo else ["disabled"])
        if attivo:
            # I menu a tendina restano di sola scelta
            self.font.state(["readonly"])
            self.colore.state(["readonly"])


class EditorExcel(tk.Tk):
    def __init__(self):
        super().__init__()
        dimensiona(self, 1100, 690, minima=(880, 560))
        try:
            self.config_servizio = carica_config()
        except Exception:
            self.config_servizio = {"cartella_lavoro": "", "cartella_template": "", "cartella_output": ""}
        self.percorso_modello = None
        self.percorso_testo = None
        self.testo_esempio = ""
        # Righe del testo divise in campi, e una colonna del .gui per ogni colonna
        self.righe = []
        self.colonne = []
        self.selezionate = []
        # Righe speciali: stili (nell'ordine del .gui), quello scelto nell'elenco, righe scelte nella tabella
        self.stili = []
        self.stile_corrente = None
        self.righe_selezionate = set()
        # Riga del foglio (da 1) -> stile, ricalcolata quando cambiano gli stili
        self.stile_della_riga = {}
        self.numeri_riga = []
        self.modificato = False
        self.da_annullare = []
        self.da_ripetere = []
        self.colonna_copiata = None
        self.celle = []
        self.lettere = []
        # Mentre si mostrano i valori di una colonna nei campi, i loro eventi non sono modifiche
        self._caricamento = False
        self.preferenze = leggi_preferenze()

        self.uscita = tk.StringVar(value=modello_gui.SEGNAPOSTO + ".xlsx")
        self.invio = tk.StringVar(value="0")
        self.destinatario = tk.StringVar()
        self.formato = tk.StringVar()
        self.larghezza = tk.StringVar()
        self.allineamento = tk.StringVar()
        self.nome_stile = tk.StringVar()
        self.righe_stile = tk.StringVar()
        self.carattere_proprio = tk.BooleanVar()
        self.sfondo_stile = tk.StringVar()
        self.allineamento_stile = tk.StringVar()
        self.testo_stile = tk.BooleanVar()

        self._crea_menu()
        self._crea_documento()
        self.stato = ttk.Label(self, text="File > Nuovo da testo di esempio per iniziare", foreground="gray",
                               padding=(8, 2))
        self.stato.pack(side="bottom", fill="x")
        self._crea_centro()
        for variabile in (self.uscita, self.invio, self.destinatario):
            variabile.trace_add("write", lambda *a: self._segna_modificato())
        self._mostra_proprieta()
        self._mostra_stile()
        self._aggiorna_titolo()
        self.protocol("WM_DELETE_WINDOW", self._esci)

    # --- Interfaccia -------------------------------------------------------

    def _crea_menu(self):
        menu = tk.Menu(self)
        file = tk.Menu(menu, tearoff=False)
        file.add_command(label="Nuovo da testo di esempio...", command=self.nuovo, accelerator="Ctrl+N")
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
        modifica.add_command(label="Annulla", command=self.annulla, accelerator="Ctrl+Z")
        modifica.add_command(label="Ripeti", command=self.ripeti, accelerator="Ctrl+Y")
        modifica.add_separator()
        modifica.add_command(label="Copia formato colonna", command=self.copia_colonna,
                             accelerator="Ctrl+Maiusc+C")
        modifica.add_command(label="Incolla formato colonna", command=self.incolla_colonna,
                             accelerator="Ctrl+Maiusc+V")
        modifica.add_command(label="Formato predefinito", command=self.colonna_predefinita)
        modifica.add_separator()
        modifica.add_command(label="Seleziona tutte le colonne", command=self.seleziona_tutte,
                             accelerator="Ctrl+A")
        menu.add_cascade(label="Modifica", menu=modifica)
        menu.add_command(label="Anteprima (F5)", command=self.anteprima)
        aiuto = tk.Menu(menu, tearoff=False)
        aiuto.add_command(label="Guida all'uso", command=lambda: apri_guida(self), accelerator="F1")
        aiuto.add_command(label="Informazioni", command=self._informazioni)
        menu.add_cascade(label="Aiuto", menu=aiuto)
        self.configure(menu=menu)
        self.bind("<Control-n>", lambda e: self.nuovo())
        self.bind("<Control-o>", lambda e: self.apri())
        self.bind("<Control-s>", lambda e: self.salva())
        self.bind("<F5>", lambda e: self.anteprima())
        self.bind("<F1>", lambda e: apri_guida(self))
        # Questi tasti valgono per le colonne solo se non si sta scrivendo in un campo
        tasti = {("<Control-z>", "<Control-Z>"): self.annulla,
                 ("<Control-y>", "<Control-Y>", "<Control-Shift-Z>"): self.ripeti,
                 ("<Control-Shift-C>",): self.copia_colonna, ("<Control-Shift-V>",): self.incolla_colonna,
                 ("<Control-a>",): self.seleziona_tutte}
        for sequenze, comando in tasti.items():
            for sequenza in sequenze:
                self.bind(sequenza, lambda e, c=comando: None if self._scrivendo() else c())

    def _scrivendo(self):
        return isinstance(self.focus_get(), (tk.Entry, ttk.Entry, tk.Text))

    def _crea_documento(self):
        riquadro = ttk.LabelFrame(self, text="Documento", padding=(8, 4))
        riquadro.pack(side="top", fill="x", padx=6, pady=(6, 0))
        ttk.Label(riquadro, text="Documento creato").grid(row=0, column=0, sticky="w")
        ttk.Entry(riquadro, textvariable=self.uscita, width=24).grid(row=0, column=1, sticky="w")
        ttk.Label(riquadro, text="Invio").grid(row=0, column=2, sticky="w", padx=(16, 4))
        SceltaDescritta(riquadro, self.invio, modello_gui.DESCRIZIONI_INVII, width=12).grid(
            row=0, column=3, sticky="w")
        ttk.Label(riquadro, text="Destinatario").grid(row=0, column=4, sticky="w", padx=(16, 4))
        ttk.Entry(riquadro, textvariable=self.destinatario, width=28).grid(row=0, column=5, sticky="we")
        ttk.Label(riquadro, text="Dati dell'invio\n(righe dopo \"fine\")").grid(row=1, column=0, sticky="nw",
                                                                             pady=(4, 0))
        self.coda = tk.Text(riquadro, height=3, width=60, font=("Consolas", 9), undo=True)
        self.coda.grid(row=1, column=1, columnspan=5, sticky="we", pady=(4, 0))
        self.coda.bind("<<Modified>>", self._coda_modificata)
        ttk.Label(riquadro, text="{nome} = nome del file di testo, senza .txt", foreground="gray").grid(
            row=2, column=1, columnspan=5, sticky="w")
        riquadro.columnconfigure(5, weight=1)

    def _crea_centro(self):
        centro = ttk.Frame(self)
        centro.pack(side="top", fill="both", expand=True, padx=6, pady=6)

        # A sinistra due schede: le proprietà delle colonne selezionate e le righe speciali
        self.schede = ttk.Notebook(centro)
        self.schede.pack(side="left", fill="y", padx=(0, 6))
        scheda_colonne = ttk.Frame(self.schede, padding=4)
        scheda_righe = ttk.Frame(self.schede, padding=4)
        self.schede.add(scheda_colonne, text="Colonne")
        self.schede.add(scheda_righe, text="Righe speciali")
        self._crea_scheda_righe(scheda_righe)
        self.pannello = ttk.LabelFrame(scheda_colonne, text="Colonna", padding=8)
        self.pannello.pack(fill="both", expand=True)
        ttk.Label(self.pannello, text="Formato dei valori").pack(anchor="w")
        self.scelta_formato = ttk.Combobox(self.pannello, textvariable=self.formato, values=NOMI_FORMATI,
                                           width=44, state="readonly")
        self.scelta_formato.pack(anchor="w", fill="x")
        self.scelta_formato.bind("<<ComboboxSelected>>", lambda e: self._formato_cambiato())
        self.avviso = ttk.Label(self.pannello, text="", foreground="#b05000", wraplength=330, justify="left")
        self.avviso.pack(anchor="w", pady=(2, 6))

        self.gruppo_celle = GruppoCarattere(self.pannello, "Carattere delle celle",
                                            lambda: self._carattere_cambiato("carattere", self.gruppo_celle))
        self.gruppo_celle.pack(fill="x", pady=(0, 6))

        altro = ttk.Frame(self.pannello)
        altro.pack(fill="x")
        ttk.Label(altro, text="Larghezza").grid(row=0, column=0, sticky="w")
        voce_larghezza = ttk.Spinbox(altro, from_=1, to=255, textvariable=self.larghezza, width=6,
                                     command=self._larghezza_cambiata)
        voce_larghezza.grid(row=0, column=1, sticky="w", padx=(4, 4))
        voce_larghezza.bind("<Return>", lambda e: self._larghezza_cambiata())
        voce_larghezza.bind("<FocusOut>", lambda e: self._larghezza_cambiata())
        ttk.Label(altro, text="caratteri (vuoto = automatica)", foreground="gray").grid(row=0, column=2,
                                                                                       sticky="w")
        ttk.Label(altro, text="Allineamento").grid(row=1, column=0, sticky="w", pady=(6, 0))
        scelta_allineamento = ttk.Combobox(altro, textvariable=self.allineamento, values=NOMI_ALLINEAMENTI,
                                           width=11, state="readonly")
        scelta_allineamento.grid(row=1, column=1, columnspan=2, sticky="w", padx=(4, 0), pady=(6, 0))
        scelta_allineamento.bind("<<ComboboxSelected>>", lambda e: self._allineamento_cambiato())

        pulsanti = ttk.Frame(self.pannello)
        pulsanti.pack(fill="x", pady=(12, 0))
        ttk.Button(pulsanti, text="Copia formato", command=self.copia_colonna).pack(side="left")
        ttk.Button(pulsanti, text="Incolla formato", command=self.incolla_colonna).pack(side="left", padx=(4, 0))
        ttk.Button(pulsanti, text="Predefinito", command=self.colonna_predefinita).pack(side="left", padx=(4, 0))
        ttk.Button(self.pannello, text="Anteprima in Excel (F5)", command=self.anteprima).pack(fill="x",
                                                                                            pady=(12, 0))
        ttk.Label(self.pannello, text="Clic su una colonna per sceglierla; Ctrl+clic o Maiusc+clic per più "
                                      "colonne insieme. Le righe speciali (scheda accanto) valgono più della "
                                      "colonna.", foreground="gray", wraplength=330,
                  justify="left").pack(anchor="w", side="bottom")

        # A destra la tabella con le celle come le scriverà il servizio
        tabella = ttk.Frame(centro)
        tabella.pack(side="left", fill="both", expand=True)
        self.tela = tk.Canvas(tabella, background="#d4d4d4", highlightthickness=0, takefocus=1)
        verticale = ttk.Scrollbar(tabella, orient="vertical", command=self.tela.yview)
        orizzontale = ttk.Scrollbar(tabella, orient="horizontal", command=self.tela.xview)
        self.tela.configure(yscrollcommand=verticale.set, xscrollcommand=orizzontale.set)
        verticale.pack(side="right", fill="y")
        orizzontale.pack(side="bottom", fill="x")
        self.tela.pack(side="left", fill="both", expand=True)
        self.griglia = tk.Frame(self.tela, background="#c8c8c8")
        self.tela.create_window((0, 0), window=self.griglia, anchor="nw")
        self.griglia.bind("<Configure>", lambda e: self.tela.configure(scrollregion=self.tela.bbox("all")))
        # Le celle hanno il tag "CellaExcel": clic e rotellina valgono per tutte
        self.bind_class("CellaExcel", "<Button-1>", self._clic_cella)
        self.bind_class("CellaExcel", "<MouseWheel>",
                        lambda e: self.tela.yview_scroll(-1 if e.delta > 0 else 1, "units"))
        self.bind_class("CellaExcel", "<Button-4>", lambda e: self.tela.yview_scroll(-1, "units"))
        self.bind_class("CellaExcel", "<Button-5>", lambda e: self.tela.yview_scroll(1, "units"))
        self.tela.bind("<Left>", lambda e: self._sposta_selezione(-1, e))
        self.tela.bind("<Right>", lambda e: self._sposta_selezione(1, e))
        self.tela.bind("<Up>", lambda e: self._sposta_riga(-1, e))
        self.tela.bind("<Down>", lambda e: self._sposta_riga(1, e))
        # La scheda aperta decide se i clic sulla tabella scelgono colonne o righe
        self.schede.bind("<<NotebookTabChanged>>", lambda e: self._scheda_cambiata())
        # I numeri di riga hanno il tag "NumeroRigaExcel": clic per scegliere le righe speciali
        self.bind_class("NumeroRigaExcel", "<Button-1>", self._clic_numero_riga)
        for evento in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.bind_class("NumeroRigaExcel", evento, self.bind_class("CellaExcel", evento))

    def _crea_scheda_righe(self, scheda):
        ttk.Label(scheda, text="Clic su una riga della tabella: se ha uno stile lo modifichi qui sotto, "
                               "altrimenti premi \"Nuovo\". Vale lo stile più in basso nell'elenco.",
                  wraplength=340, justify="left",
                  foreground="gray").pack(anchor="w", pady=(0, 6))
        elenco = ttk.Frame(scheda)
        elenco.pack(fill="x")
        self.lista_stili = tk.Listbox(elenco, height=4, exportselection=False, activestyle="none")
        self.lista_stili.pack(side="left", fill="x", expand=True)
        self.lista_stili.bind("<<ListboxSelect>>", lambda e: self._scegli_stile())
        pulsanti = ttk.Frame(elenco)
        pulsanti.pack(side="left", fill="y", padx=(6, 0))
        for testo, comando in (("Nuovo", self.nuovo_stile), ("Elimina", self.elimina_stile),
                               ("Su", lambda: self.sposta_stile(-1)), ("Giù", lambda: self.sposta_stile(1))):
            ttk.Button(pulsanti, text=testo, width=8, command=comando).pack(pady=(0, 3))

        self.riquadro_stile = ttk.LabelFrame(scheda, text="Stile", padding=8)
        self.riquadro_stile.pack(fill="both", expand=True, pady=(8, 0))
        r = self.riquadro_stile
        ttk.Label(r, text="Nome").grid(row=0, column=0, sticky="w")
        voce_nome = ttk.Entry(r, textvariable=self.nome_stile, width=26)
        voce_nome.grid(row=0, column=1, columnspan=2, sticky="we", padx=(4, 0))
        voce_nome.bind("<Return>", lambda e: self._nome_stile_cambiato())
        voce_nome.bind("<FocusOut>", lambda e: self._nome_stile_cambiato())
        ttk.Label(r, text="Righe").grid(row=1, column=0, sticky="w", pady=(6, 0))
        voce_righe = ttk.Entry(r, textvariable=self.righe_stile, width=18)
        voce_righe.grid(row=1, column=1, sticky="we", padx=(4, 0), pady=(6, 0))
        voce_righe.bind("<Return>", lambda e: self._righe_stile_cambiate())
        voce_righe.bind("<FocusOut>", lambda e: self._righe_stile_cambiate())
        ttk.Button(r, text="Righe scelte", command=self.usa_righe_scelte).grid(row=1, column=2, sticky="w",
                                                                              padx=(4, 0), pady=(6, 0))
        ttk.Label(r, text="es. 1-3, 10, u1   (u1 = ultima riga, u2 = penultima). Oppure fai clic sulle "
                          "righe della tabella e poi \"Righe scelte\".", foreground="gray", wraplength=320,
                  justify="left").grid(row=2, column=0, columnspan=3, sticky="w")
        ttk.Checkbutton(r, text="Carattere proprio (senza: quello della colonna)", variable=self.carattere_proprio,
                        command=self._carattere_stile_cambiato).grid(row=3, column=0, columnspan=3, sticky="w",
                                                                    pady=(8, 0))
        self.gruppo_stile = GruppoCarattere(r, "Carattere", self._carattere_stile_cambiato)
        self.gruppo_stile.grid(row=4, column=0, columnspan=3, sticky="we", pady=(2, 0))
        ttk.Label(r, text="Sfondo").grid(row=5, column=0, sticky="w", pady=(8, 0))
        scelta_sfondo = ttk.Combobox(r, textvariable=self.sfondo_stile, values=NOMI_SFONDI, width=16,
                                     state="readonly")
        scelta_sfondo.grid(row=5, column=1, columnspan=2, sticky="w", padx=(4, 0), pady=(8, 0))
        scelta_sfondo.bind("<<ComboboxSelected>>", lambda e: self._sfondo_stile_cambiato())
        ttk.Label(r, text="Allineamento").grid(row=6, column=0, sticky="w", pady=(6, 0))
        scelta_allineamento = ttk.Combobox(r, textvariable=self.allineamento_stile, values=NOMI_ALLINEAMENTI,
                                           width=16, state="readonly")
        scelta_allineamento.grid(row=6, column=1, columnspan=2, sticky="w", padx=(4, 0), pady=(6, 0))
        scelta_allineamento.bind("<<ComboboxSelected>>", lambda e: self._allineamento_stile_cambiato())
        ttk.Checkbutton(r, text="Valori come testo (non diventano numeri o date)", variable=self.testo_stile,
                        command=self._testo_stile_cambiato).grid(row=7, column=0, columnspan=3, sticky="w",
                                                                pady=(8, 0))
        r.columnconfigure(1, weight=1)

    # --- Tabella -------------------------------------------------------------

    def _cella(self, testo, riga, colonna, **opzioni):
        # Il bordo (highlight) segna le righe scelte senza coprire il loro sfondo
        etichetta = tk.Label(self.griglia, text=testo, padx=4, pady=1, borderwidth=0, highlightthickness=1,
                             **opzioni)
        # riga: numero della riga del foglio (0 = le lettere delle colonne)
        etichetta.riga = riga
        etichetta.colonna = colonna
        etichetta.bindtags(("CellaExcel",) + etichetta.bindtags())
        etichetta.grid(row=riga, column=colonna + 1, sticky="nsew", padx=(0, 1), pady=(0, 1))
        return etichetta

    def _crea_tabella(self):
        for figlio in self.griglia.winfo_children():
            figlio.destroy()
        self.lettere = []
        self.celle = []
        self.numeri_riga = []
        mostrate = self.righe[:RIGHE_MOSTRATE]
        tk.Label(self.griglia, text="", background=SFONDO_LETTERE).grid(row=0, column=0, sticky="nsew",
                                                                       padx=(0, 1), pady=(0, 1))
        for c in range(len(self.colonne)):
            self.lettere.append(self._cella(lettera_colonna(c), 0, c, background=SFONDO_LETTERE,
                                            font=("Segoe UI", 9, "bold")))
        for r, campi in enumerate(mostrate):
            numero = tk.Label(self.griglia, text=str(r + 1), background=SFONDO_LETTERE, padx=4,
                              font=("Segoe UI", 8))
            numero.riga = r + 1
            numero.bindtags(("NumeroRigaExcel",) + numero.bindtags())
            numero.grid(row=r + 1, column=0, sticky="nsew", padx=(0, 1), pady=(0, 1))
            self.numeri_riga.append(numero)
            self.celle.append([self._cella("", r + 1, c) for c in range(len(self.colonne))])
        if len(self.righe) > len(mostrate):
            ttk.Label(self.griglia, text="... altre %d righe (nell'anteprima ci sono tutte)"
                      % (len(self.righe) - len(mostrate)), foreground="gray").grid(
                row=len(mostrate) + 1, column=1, columnspan=max(1, len(self.colonne)), sticky="w", pady=4)
        self._aggiorna_tutto()

    def _ricalcola_stili(self):
        # Come il servizio: le righe di ogni stile, l'ultimo stile vince; righe scritte male = nessuna riga
        self.stile_della_riga = {}
        for stile in self.stili:
            try:
                for numero in righe_excel(stile["righe"], len(self.righe)):
                    self.stile_della_riga[numero] = stile
            except ValueError:
                pass

    def _aggiorna_tutto(self):
        self._ricalcola_stili()
        for c in range(len(self.colonne)):
            self._aggiorna_colonna(c)
        self._aggiorna_numeri_riga()

    def _aggiorna_numeri_riga(self):
        for etichetta in self.numeri_riga:
            if etichetta.riga in self.righe_selezionate and self._modalita_righe():
                sfondo = SFONDO_LETTERE_SELEZIONE
            elif etichetta.riga in self.stile_della_riga:
                sfondo = SFONDO_RIGA_SPECIALE
            else:
                sfondo = SFONDO_LETTERE
            etichetta.configure(background=sfondo)

    def _modalita_righe(self):
        """Con la scheda "Righe speciali" aperta i clic sulla tabella scelgono righe, non colonne."""
        return self.schede.index("current") == 1

    def _scheda_cambiata(self):
        self._aggiorna_tutto()
        if self._modalita_righe():
            self._messaggio("Righe speciali: fai clic su una riga della tabella per sceglierla (Ctrl+clic e "
                            "Maiusc+clic per più righe, frecce su e giù per spostarti)")
        else:
            self._messaggio("Colonne: fai clic su una colonna per sceglierla (Ctrl+clic e Maiusc+clic per più "
                            "colonne)")

    def _clic_numero_riga(self, evento):
        # Il clic sul numero di riga sceglie la riga anche dalla scheda Colonne
        if not self._modalita_righe():
            self.schede.select(1)
        self.tela.focus_set()
        self.seleziona_righe(evento.widget.riga, evento.state)

    def seleziona_righe(self, riga, stato=0):
        """Sceglie la riga (con Maiusc fino a questa, con Ctrl aggiunta o tolta) e mostra il suo stile."""
        if stato & 0x1 and self.righe_selezionate:
            # Maiusc: dalla riga scelta più vicina a questa
            vicina = min(self.righe_selezionate, key=lambda n: abs(n - riga))
            self.righe_selezionate |= set(range(min(vicina, riga), max(vicina, riga) + 1))
        elif stato & 0x4:
            self.righe_selezionate ^= {riga}
        else:
            self.righe_selezionate = {riga}
        scelte = intervalli(self.righe_selezionate) or "nessuna"
        stili = {id(self.stile_della_riga[n]): self.stile_della_riga[n]
                 for n in self.righe_selezionate if n in self.stile_della_riga}
        senza_stile = [n for n in self.righe_selezionate if n not in self.stile_della_riga]
        if len(stili) == 1 and not senza_stile:
            # Tutte le righe scelte hanno lo stesso stile: si apre per modificarlo
            stile = next(iter(stili.values()))
            self.stile_corrente = next(i for i, s in enumerate(self.stili) if s is stile)
            self._messaggio("Righe %s: stile \"%s\" (righe %s). Modificalo nel pannello a sinistra."
                            % (scelte, stile["nome"] or "senza nome", stile["righe"]))
        elif not stili:
            self.stile_corrente = None
            self._messaggio("Righe %s: nessuno stile. Premi \"Nuovo\" per crearne uno con queste righe." % scelte)
        else:
            self.stile_corrente = None
            self._messaggio("Righe %s: hanno stili diversi. Scegline una sola, oppure premi \"Nuovo\" per dare "
                            "a tutte uno stile nuovo." % scelte)
        self._aggiorna_tutto()
        self._mostra_stile()

    def _sposta_riga(self, passo, evento):
        if not self._modalita_righe() or not self.celle:
            return
        corrente = max(self.righe_selezionate) if passo > 0 and self.righe_selezionate else \
            min(self.righe_selezionate or {1})
        nuova = max(1, min(len(self.celle), corrente + passo))
        self.seleziona_righe(nuova, evento.state & 0x1)

    def _valore_grezzo(self, riga, colonna):
        campi = self.righe[riga]
        return campi[colonna] if colonna < len(campi) else ""

    def _aggiorna_colonna(self, c):
        if c >= len(self.lettere):
            return
        colonna = self.colonne[c]
        # Nella scheda delle righe speciali si vedono le righe scelte, non le colonne
        modalita_righe = self._modalita_righe()
        selezionata = c in self.selezionate and not modalita_righe
        sfondo_lettera = SFONDO_LETTERE_SELEZIONE if selezionata else SFONDO_LETTERE
        self.lettere[c].configure(background=sfondo_lettera, highlightbackground=sfondo_lettera)
        vuoto = stile_riga_vuoto()
        stili = [self.stile_della_riga.get(r + 1, vuoto) for r in range(len(self.celle))]
        # Valore che il servizio scriverà nella cella: testo nelle righe "come testo", altrimenti convertito
        valori = [self._valore_grezzo(r, c) if stili[r]["testo"]
                  else valore_excel(self._valore_grezzo(r, c), colonna["formato"]) for r in range(len(self.celle))]
        testi = [self._valore_grezzo(r, c) if stili[r]["testo"]
                 else testo_mostrato(self._valore_grezzo(r, c), colonna["formato"]) for r in range(len(self.celle))]
        larghezza = colonna["larghezza"] or min(60, max([len(t) for t in testi] + [3]) + 2)
        for r, etichetta in enumerate(row[c] for row in self.celle):
            stile = stili[r]
            # La riga speciale vale più della colonna
            formato = stile["carattere"] or colonna["carattere"]
            colore = (formato or PREDEFINITO_EXCEL)[4]
            allineamento = stile["allineamento"] or colonna["allineamento"]
            if allineamento:
                ancora = {"sinistra": "w", "centro": "center", "destra": "e"}[allineamento]
            else:
                # Come Excel: numeri e date a destra, testo a sinistra
                ancora = "w" if isinstance(valori[r], str) else "e"
            if stile["sfondo"]:
                sfondo = "#" + SFONDI[stile["sfondo"]][1]
            else:
                sfondo = SFONDO_SELEZIONE if selezionata else SFONDO
            if colore == "w" and not stile["sfondo"]:
                # Il bianco si vede solo su uno sfondo scuro
                sfondo = "#6f8fb0" if selezionata else "#8a8a8a"
            riga_scelta = modalita_righe and r + 1 in self.righe_selezionate
            etichetta.configure(text=testi[r], font=font_tk(formato), foreground="#" + COLORI[colore][1],
                                background=sfondo, anchor=ancora, width=larghezza,
                                highlightbackground=BORDO_RIGA_SCELTA if riga_scelta else sfondo)

    def _clic_cella(self, evento):
        c = evento.widget.colonna
        self.tela.focus_set()
        if self._modalita_righe():
            if evento.widget.riga > 0:
                self.seleziona_righe(evento.widget.riga, evento.state)
                return
            # Clic sulla lettera di una colonna: si torna alle colonne
            self.schede.select(0)
        if evento.state & 0x1 and self.selezionate:
            # Maiusc: dalla colonna scelta prima a questa
            inizio = self.selezionate[0]
            self.seleziona(list(range(min(inizio, c), max(inizio, c) + 1)), mantieni_primo=inizio)
        elif evento.state & 0x4:
            # Ctrl: aggiunge o toglie la colonna
            scelte = [x for x in self.selezionate if x != c] if c in self.selezionate else self.selezionate + [c]
            self.seleziona(scelte or [c])
        else:
            self.seleziona([c])

    def _sposta_selezione(self, passo, evento):
        if not self.colonne or self._modalita_righe():
            return
        corrente = self.selezionate[-1] if self.selezionate else 0
        nuova = max(0, min(len(self.colonne) - 1, corrente + passo))
        if evento.state & 0x1:
            scelte = self.selezionate if nuova in self.selezionate else self.selezionate + [nuova]
            self.seleziona(scelte)
        else:
            self.seleziona([nuova])

    def seleziona(self, colonne, mantieni_primo=None):
        vecchie = set(self.selezionate)
        self.selezionate = [c for c in colonne if 0 <= c < len(self.colonne)]
        if mantieni_primo is not None and mantieni_primo in self.selezionate:
            # La prima selezionata resta il punto di partenza di Maiusc+clic
            self.selezionate.remove(mantieni_primo)
            self.selezionate.insert(0, mantieni_primo)
        for c in vecchie | set(self.selezionate):
            self._aggiorna_colonna(c)
        self._mostra_proprieta()

    def seleziona_tutte(self):
        self.seleziona(list(range(len(self.colonne))))

    # --- Proprietà della colonna --------------------------------------------

    def _mostra_proprieta(self):
        """Mostra nei campi la prima colonna selezionata."""
        self._caricamento = True
        try:
            if not self.selezionate:
                self.pannello.configure(text="Colonna")
                self.avviso.configure(text="")
                return
            colonna = self.colonne[self.selezionate[0]]
            nomi = ", ".join(self._nome_colonna(c) for c in self.selezionate[:4])
            if len(self.selezionate) > 4:
                nomi += " e altre %d" % (len(self.selezionate) - 4)
            if len(self.selezionate) > 1:
                # Le scelte valgono per tutte; i campi mostrano quelle della prima
                nomi = "Colonne " + nomi + " - si vedono i valori di " + lettera_colonna(self.selezionate[0])
            else:
                nomi = "Colonna " + nomi
            self.pannello.configure(text=nomi)
            self.formato.set(NOMI_FORMATI[list(FORMATI_EXCEL).index(colonna["formato"])])
            self.gruppo_celle.imposta(colonna["carattere"])
            self.larghezza.set(str(colonna["larghezza"]) if colonna["larghezza"] else "")
            self.allineamento.set(colonna["allineamento"] or "automatico")
            self.avviso.configure(text=self._avviso_conversione())
        finally:
            self._caricamento = False

    def _nome_colonna(self, c):
        # Il nome è il testo della prima riga "come testo" (di solito l'intestazione)
        righe_testo = sorted(n for n, s in self.stile_della_riga.items() if s["testo"])
        intestazione = self._valore_grezzo(righe_testo[0] - 1, c).strip() if righe_testo else ""
        return lettera_colonna(c) + (" (" + intestazione[:20] + ")" if intestazione else "")

    def _avviso_conversione(self):
        """Avvisa se alcuni valori non sono numeri o date: in Excel resteranno testo."""
        messaggi = []
        for c in self.selezionate:
            codice = self.colonne[c]["formato"]
            if codice not in CODICI_CONVERTITI:
                continue
            # Le righe speciali "come testo" (intestazione, titoli) non si convertono
            sbagliati = [self._valore_grezzo(r, c) for r in range(len(self.righe))
                         if not self.stile_della_riga.get(r + 1, stile_riga_vuoto())["testo"]
                         and self._valore_grezzo(r, c).strip() != ""
                         and isinstance(valore_excel(self._valore_grezzo(r, c), codice), str)]
            if sbagliati:
                messaggi.append("%s: %d valori non si possono convertire e resteranno testo (es. \"%s\")"
                                % (lettera_colonna(c), len(sbagliati), sbagliati[0][:20]))
        return "\n".join(messaggi)

    def _imposta_colonne(self, chiave, valore):
        """Cambia una proprietà di tutte le colonne selezionate (con annulla)."""
        if self._caricamento or not self.selezionate:
            return
        if all(self.colonne[c][chiave] == valore for c in self.selezionate):
            return
        self._prima_di_cambiare()
        for c in self.selezionate:
            self.colonne[c][chiave] = valore
        self._dopo_cambio(self.selezionate)

    def _formato_cambiato(self):
        self._imposta_colonne("formato", int(self.formato.get().split(" - ")[0]))
        self.avviso.configure(text=self._avviso_conversione())

    def _carattere_cambiato(self, chiave, gruppo):
        if self._caricamento:
            return
        try:
            formato = gruppo.valore()
        except ValueError as e:
            self._messaggio(str(e))
            self._mostra_proprieta()
            return
        self._imposta_colonne(chiave, formato)

    def _larghezza_cambiata(self):
        if self._caricamento:
            return
        testo = self.larghezza.get().strip()
        if testo == "":
            self._imposta_colonne("larghezza", None)
        elif testo.isdigit() and 1 <= int(testo) <= 255:
            self._imposta_colonne("larghezza", int(testo))
        else:
            self._messaggio("Larghezza non valida: da 1 a 255 caratteri, oppure vuota (automatica)")
            self._mostra_proprieta()

    def _allineamento_cambiato(self):
        valore = self.allineamento.get()
        self._imposta_colonne("allineamento", None if valore == "automatico" else valore)

    def _dopo_cambio(self, colonne):
        self._segna_modificato()
        for c in colonne:
            self._aggiorna_colonna(c)
        self._mostra_proprieta()

    # --- Annulla, copia formato ------------------------------------------------

    def _istantanea(self):
        return {"colonne": [dict(c) for c in self.colonne], "stili": [dict(s) for s in self.stili]}

    def _prima_di_cambiare(self):
        self.da_annullare.append(self._istantanea())
        del self.da_annullare[:-MAX_ANNULLA]
        self.da_ripetere = []

    def _ripristina(self, istantanea):
        self.colonne = istantanea["colonne"]
        self.stili = istantanea["stili"]
        if self.stile_corrente is not None and self.stile_corrente >= len(self.stili):
            self.stile_corrente = len(self.stili) - 1 if self.stili else None
        self._dopo_cambio_stili()

    def annulla(self):
        if not self.da_annullare:
            self._messaggio("Niente da annullare")
            return
        self.da_ripetere.append(self._istantanea())
        self._ripristina(self.da_annullare.pop())

    def ripeti(self):
        if not self.da_ripetere:
            self._messaggio("Niente da ripetere")
            return
        self.da_annullare.append(self._istantanea())
        self._ripristina(self.da_ripetere.pop())

    def copia_colonna(self):
        if not self.selezionate:
            return
        self.colonna_copiata = dict(self.colonne[self.selezionate[0]])
        self._messaggio("Formato della colonna %s copiato: scegli le colonne e usa Incolla formato"
                        % lettera_colonna(self.selezionate[0]))

    def incolla_colonna(self):
        if self.colonna_copiata is None:
            self._messaggio("Prima copia il formato di una colonna")
            return
        if not self.selezionate:
            return
        self._prima_di_cambiare()
        for c in self.selezionate:
            self.colonne[c] = dict(self.colonna_copiata)
        self._dopo_cambio(self.selezionate)

    def colonna_predefinita(self):
        if not self.selezionate:
            return
        self._prima_di_cambiare()
        for c in self.selezionate:
            self.colonne[c] = colonna_excel_vuota()
        self._dopo_cambio(self.selezionate)

    # --- Righe speciali ----------------------------------------------------------

    def _dopo_cambio_stili(self):
        self._segna_modificato()
        self._aggiorna_tutto()
        self._mostra_proprieta()
        self._mostra_stile()

    def _mostra_stile(self):
        """Aggiorna l'elenco degli stili e i campi dello stile scelto."""
        self._caricamento = True
        try:
            self.lista_stili.delete(0, "end")
            for stile in self.stili:
                self.lista_stili.insert("end", "%s   [%s]" % (stile["nome"] or "(senza nome)", stile["righe"] or "-"))
            attivo = self.stile_corrente is not None
            for figlio in self.riquadro_stile.winfo_children():
                if isinstance(figlio, ttk.Widget) and figlio is not self.gruppo_stile:
                    figlio.state(["!disabled"] if attivo else ["disabled"])
            if not attivo:
                self.gruppo_stile.abilita(False)
                for variabile in (self.nome_stile, self.righe_stile, self.sfondo_stile, self.allineamento_stile):
                    variabile.set("")
                return
            self.lista_stili.selection_set(self.stile_corrente)
            self.lista_stili.see(self.stile_corrente)
            stile = self.stili[self.stile_corrente]
            self.riquadro_stile.configure(text="Stile: " + (stile["nome"] or "senza nome"))
            self.nome_stile.set(stile["nome"])
            self.righe_stile.set(stile["righe"])
            self.carattere_proprio.set(stile["carattere"] is not None)
            self.gruppo_stile.imposta(stile["carattere"])
            self.gruppo_stile.abilita(stile["carattere"] is not None)
            self.sfondo_stile.set(SFONDI[stile["sfondo"]][0] if stile["sfondo"] else "nessuno")
            self.allineamento_stile.set(stile["allineamento"] or "automatico")
            self.testo_stile.set(stile["testo"])
        finally:
            self._caricamento = False

    def _scegli_stile(self):
        scelta = self.lista_stili.curselection()
        if scelta:
            self.stile_corrente = scelta[0]
            stile = self.stili[self.stile_corrente]
            # Le righe dello stile diventano quelle scelte nella tabella
            try:
                self.righe_selezionate = righe_excel(stile["righe"], len(self.righe))
            except ValueError:
                self.righe_selezionate = set()
            self._aggiorna_tutto()
            self._mostra_stile()

    def nuovo_stile(self):
        self._prima_di_cambiare()
        stile = stile_riga_vuoto()
        stile["nome"] = "Stile %d" % (len(self.stili) + 1)
        # Parte dalle righe scelte nella tabella, e in grassetto
        stile["righe"] = intervalli(self.righe_selezionate)
        stile["carattere"] = ("k", True, False, 11, "n")
        self.stili.append(stile)
        self.stile_corrente = len(self.stili) - 1
        self._dopo_cambio_stili()

    def elimina_stile(self):
        if self.stile_corrente is None:
            return
        self._prima_di_cambiare()
        del self.stili[self.stile_corrente]
        self.stile_corrente = min(self.stile_corrente, len(self.stili) - 1) if self.stili else None
        self._dopo_cambio_stili()

    def sposta_stile(self, passo):
        """Su/giù nell'elenco: se una riga è in più stili vale quello più in basso."""
        if self.stile_corrente is None:
            return
        nuovo = self.stile_corrente + passo
        if not 0 <= nuovo < len(self.stili):
            return
        self._prima_di_cambiare()
        self.stili[self.stile_corrente], self.stili[nuovo] = self.stili[nuovo], self.stili[self.stile_corrente]
        self.stile_corrente = nuovo
        self._dopo_cambio_stili()

    def _imposta_stile(self, chiave, valore):
        if self._caricamento or self.stile_corrente is None:
            return
        if self.stili[self.stile_corrente][chiave] == valore:
            return
        self._prima_di_cambiare()
        self.stili[self.stile_corrente][chiave] = valore
        self._dopo_cambio_stili()

    def _nome_stile_cambiato(self):
        self._imposta_stile("nome", " ".join(self.nome_stile.get().split()))

    def _righe_stile_cambiate(self):
        if self._caricamento or self.stile_corrente is None:
            return
        testo = self.righe_stile.get().strip()
        try:
            righe_excel(testo, RIGHE_MASSIME)
        except ValueError as e:
            self._messaggio(str(e))
            self._mostra_stile()
            return
        self._imposta_stile("righe", testo)

    def usa_righe_scelte(self):
        if not self.righe_selezionate:
            self._messaggio("Prima fai clic sulle righe della tabella (Ctrl+clic e Maiusc+clic per più "
                            "righe)")
            return
        if self.stile_corrente is None:
            self.nuovo_stile()
            return
        self._imposta_stile("righe", intervalli(self.righe_selezionate))

    def _carattere_stile_cambiato(self):
        if self._caricamento or self.stile_corrente is None:
            return
        if not self.carattere_proprio.get():
            self._imposta_stile("carattere", None)
            return
        try:
            # Nella riga anche Calibri 11 nero è una scelta: senza carattere proprio vale quello della colonna
            formato = self.gruppo_stile.valore(predefinito_vuoto=False)
        except ValueError as e:
            self._messaggio(str(e))
            self._mostra_stile()
            return
        self._imposta_stile("carattere", formato)

    def _sfondo_stile_cambiato(self):
        nomi = {descrizione: nome for nome, (descrizione, _) in SFONDI.items()}
        self._imposta_stile("sfondo", nomi.get(self.sfondo_stile.get()))

    def _allineamento_stile_cambiato(self):
        valore = self.allineamento_stile.get()
        self._imposta_stile("allineamento", None if valore == "automatico" else valore)

    def _testo_stile_cambiato(self):
        self._imposta_stile("testo", self.testo_stile.get())

    # --- Lettura e scrittura del modello ------------------------------------

    def _mostra_testo(self, testo, colonne=None, stili=None):
        self.testo_esempio = testo
        # Una riga del foglio per ogni riga del testo, come nel servizio (l'a capo finale non fa una riga)
        righe_testo = testo.split("\n")
        if righe_testo and righe_testo[-1] == "":
            righe_testo.pop()
        self.righe = [campi_excel(riga) for riga in righe_testo]
        numero = max([len(r) for r in self.righe] + [len(colonne or [])])
        self.colonne = [dict(c) for c in (colonne or [])] + \
                       [colonna_excel_vuota() for _ in range(numero - len(colonne or []))]
        self.selezionate = [0] if self.colonne else []
        # Modello nuovo: la prima riga è l'intestazione, in grassetto e come testo
        self.stili = [dict(s) for s in stili] if stili is not None else [dict(STILE_INTESTAZIONE)]
        self.stile_corrente = 0 if self.stili else None
        self.righe_selezionate = set()
        self.da_annullare = []
        self.da_ripetere = []
        self._crea_tabella()
        self._mostra_proprieta()
        self._mostra_stile()
        self.tela.xview_moveto(0)
        self.tela.yview_moveto(0)
        if not self.colonne:
            self._messaggio("Il testo non ha campi separati da \";\": non è un testo per Excel")

    def _intestazione(self):
        coda = self.coda.get("1.0", "end-1c")
        return {"testo": modello_gui.SEGNAPOSTO + ".txt", "azione": "excel", "uscita": self.uscita.get().strip(),
                "template": "", "rotazione": "0", "invio": self.invio.get(),
                "destinatario": self.destinatario.get().strip(), "coda": coda.split("\n") if coda else []}

    def _imposta_intestazione(self, intestazione):
        self.uscita.set(intestazione["uscita"])
        self.invio.set(intestazione["invio"] if intestazione["invio"] in modello_gui.INVII else "0")
        self.destinatario.set(intestazione["destinatario"])
        self.coda.delete("1.0", "end")
        self.coda.insert("1.0", "\n".join(intestazione["coda"]))
        self.coda.edit_modified(False)

    def nuovo(self):
        if not self._conferma_perdita():
            return
        percorso = filedialog.askopenfilename(parent=self, title="Testo di esempio (campi separati da ;)",
                                              filetypes=[("File di testo", "*.txt"), ("Tutti i file", "*.*")])
        if not percorso:
            return
        self.percorso_testo = percorso
        self.percorso_modello = None
        self._imposta_intestazione({"uscita": modello_gui.SEGNAPOSTO + ".xlsx", "invio": "0", "destinatario": "",
                                    "coda": []})
        self._mostra_testo(leggi_testo(percorso))
        self.modificato = False
        self._aggiorna_titolo()

    def apri(self):
        if not self._conferma_perdita():
            return
        os.makedirs(CARTELLA_MODELLI, exist_ok=True)
        percorso = filedialog.askopenfilename(parent=self, title="Apri modello Excel", initialdir=CARTELLA_MODELLI,
                                              filetypes=[("Modelli", "*.gui"), ("Tutti i file", "*.*")])
        if percorso:
            self.apri_modello(percorso)

    def apri_modello(self, percorso, percorso_testo=None):
        intestazione, linee = modello_gui.leggi_gui(percorso)
        if intestazione["azione"] != "excel":
            messagebox.showinfo("Apri modello", "Non è un modello Excel: aprilo con EditorModelli.", parent=self)
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
        self._imposta_intestazione(intestazione)
        stili, colonne = modello_gui.leggi_excel(linee)
        self._mostra_testo(leggi_testo(percorso_testo), colonne, stili)
        self.percorso_modello = percorso
        self.percorso_testo = percorso_testo
        self.modificato = False
        self._aggiorna_titolo()
        self._aggiungi_recente(percorso)

    def salva(self, con_nome=False):
        if not self.righe:
            self._messaggio("Prima apri un testo di esempio")
            return False
        percorso = self.percorso_modello
        if con_nome or percorso is None:
            os.makedirs(CARTELLA_MODELLI, exist_ok=True)
            percorso = filedialog.asksaveasfilename(parent=self, title="Salva modello Excel",
                                                    initialdir=CARTELLA_MODELLI, defaultextension=".gui",
                                                    filetypes=[("Modelli", "*.gui")])
            if not percorso:
                return False
        lavoro = self.config_servizio.get("cartella_lavoro", "")
        if lavoro and os.path.normcase(os.path.abspath(os.path.dirname(percorso))) == \
                os.path.normcase(os.path.abspath(lavoro)):
            messagebox.showerror("Modello", "Non salvare il modello nella cartella di lavoro: il servizio "
                                 "proverebbe a elaborarlo. Usa ad esempio " + CARTELLA_MODELLI, parent=self)
            return False
        modello_gui.scrivi_gui(percorso, self._intestazione(), modello_gui.linee_excel(self.stili, self.colonne))
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
                                 "ConfigManager", parent=self)
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

    # --- Anteprima -------------------------------------------------------------

    def crea_documento(self, cartella):
        """Crea il file Excel con il codice del servizio; restituisce il percorso."""
        intestazione = self._intestazione()
        intestazione.update(testo="anteprima.txt", uscita="anteprima.xlsx", invio="0", destinatario="", coda=[])
        with open(os.path.join(cartella, "anteprima.txt"), "w", encoding="utf-8") as f:
            f.write(self.testo_esempio)
        percorso_gui = os.path.join(cartella, "anteprima.gui")
        modello_gui.scrivi_gui(percorso_gui, intestazione, modello_gui.linee_excel(self.stili, self.colonne))
        with open(percorso_gui, "r") as f:
            guida = f.readlines()
        config = dict(self.config_servizio, cartella_lavoro=cartella, cartella_output=cartella)
        servizio = importlib.import_module("SingleGUY")
        servizio.excel(guida, config)
        return os.path.join(cartella, "anteprima.xlsx")

    def anteprima(self):
        if not self.righe:
            self._messaggio("Prima apri un testo di esempio")
            return
        self.configure(cursor="watch")
        self.update_idletasks()
        try:
            percorso = self.crea_documento(tempfile.mkdtemp(prefix="anteprima_excel_"))
        except Exception:
            messagebox.showerror("Anteprima", "Errore nel creare il file Excel:\n\n" + traceback.format_exc(limit=-3),
                                 parent=self)
            return
        finally:
            self.configure(cursor="")
        apri_con_avviso(percorso, self)

    # --- Modelli recenti, stato, chiusura -------------------------------------

    def _aggiungi_recente(self, percorso):
        percorso = os.path.abspath(percorso)
        recenti = [p for p in self.preferenze.get("recenti_excel", [])
                   if os.path.normcase(p) != os.path.normcase(percorso)]
        self.preferenze["recenti_excel"] = ([percorso] + recenti)[:MAX_RECENTI]
        scrivi_preferenze(self.preferenze)

    def _elenca_recenti(self):
        self.menu_recenti.delete(0, "end")
        recenti = [p for p in self.preferenze.get("recenti_excel", []) if os.path.isfile(p)]
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

    def _informazioni(self):
        messagebox.showinfo("Informazioni", "Editor modelli Excel di FileGUY\n\nPrepara i modelli .gui per i file "
                            "Excel: formato dei valori, carattere, larghezza e allineamento di ogni colonna.\n\n"
                            "Guida completa: Aiuto > Guida all'uso (F1)", parent=self)

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
        self.title("Editor modelli Excel - " + nome + (" *" if self.modificato else ""))

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


if __name__ == "__main__":
    app = EditorExcel()
    if len(sys.argv) > 1:
        app.apri_modello(sys.argv[1])
    app.mainloop()
