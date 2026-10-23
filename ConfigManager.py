"""Interfaccia grafica per gestire la configurazione cifrata di FileGUY (config.dat).

Avvio: python ConfigManager.py (oppure l'eseguibile creato con PyInstaller, nella stessa
cartella del servizio).
"""
import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

import config_sicuro
from config_sicuro import CAMPI, ConfigError
from piattaforma import SCALA_FINESTRE, avvia_finestra, console_sicura, scala_finestra

# Senza console (exe di ConfigManager) i messaggi su stdout/stderr farebbero fallire l'operazione
console_sicura()


# (campo, etichetta, tipo) - tipo: "cartella", "file", "immagine", "stampante", "segreto", "testo",
# "testo_lungo" (più righe) o "spunta" (valore "1" se spuntata, vuoto se no)
SEZIONI = [
    ("Cartelle", [
        ("cartella_lavoro", "Cartella di lavoro (.txt e .gui)", "cartella"),
        ("cartella_template", "Cartella dei template PDF", "cartella"),
        ("cartella_output", "Cartella di output", "cartella"),
    ]),
    ("Stampa e log", [
        ("stampante", "Stampante (vuoto = predefinita)", "stampante"),
        ("file_log", "File di log (vuoto = FileGUY.log accanto al programma)", "file"),
    ]),
    ("Email", [
        ("email_mittente", "Indirizzo mittente", "testo"),
        ("email_password", "Password (per Gmail: password per le app)", "segreto"),
        ("email_smtp_server", "Server SMTP (SSL)", "testo"),
        ("email_smtp_porta", "Porta SMTP", "testo"),
        ("email_firma", "Firma in fondo a ogni email", "testo_lungo"),
        ("email_firma_html", "Firma scritta in HTML (es. <b>Nome</b><br><a href=\"...\">sito</a>)", "spunta"),
        ("email_firma_immagine", "Immagine della firma (logo, firma scansionata)", "immagine"),
    ]),
    ("Telegram", [
        ("telegram_token", "Token del bot", "segreto"),
    ]),
    ("WhatsApp (API Cloud di Meta)", [
        ("whatsapp_token", "Access token", "segreto"),
        ("whatsapp_phone_number_id", "Phone number ID", "testo"),
        ("whatsapp_versione_api", "Versione API Graph", "testo"),
        ("whatsapp_lingua", "Lingua predefinita dei template", "testo"),
        ("whatsapp_account_business", "ID account WhatsApp Business (vuoto = dal token)", "testo"),
    ]),
]


class ConfigManager(tk.Tk):
    def __init__(self):
        super().__init__()
        scala_finestra(self)
        self.title("FileGUY - Configurazione")
        self.resizable(True, False)
        self.minsize(int(640 * SCALA_FINESTRE), 0)

        self.variabili = {campo: tk.StringVar() for campo in CAMPI}
        self.tutto_il_pc = tk.BooleanVar(value=False)
        self.salvato = {}
        self._percorso_importato = None

        self._crea_interfaccia()
        self._carica_iniziale()
        self.protocol("WM_DELETE_WINDOW", self._chiudi)

    # --- Interfaccia -------------------------------------------------------

    def _crea_interfaccia(self):
        principale = ttk.Frame(self, padding=12)
        principale.pack(fill="both", expand=True)
        principale.columnconfigure(0, weight=1)

        riga = 0
        for titolo, campi in SEZIONI:
            riquadro = ttk.LabelFrame(principale, text=titolo, padding=8)
            riquadro.grid(row=riga, column=0, sticky="ew", pady=(0, 8))
            riquadro.columnconfigure(1, weight=1)
            for r, (campo, etichetta, tipo) in enumerate(campi):
                self._crea_campo(riquadro, r, campo, etichetta, tipo)
            if titolo.startswith("WhatsApp"):
                ttk.Button(riquadro, text="Invia messaggio di prova...", command=self._prova_whatsapp) \
                    .grid(row=len(campi), column=1, sticky="w", pady=(6, 0))
            riga += 1

        ttk.Checkbutton(
            principale, variable=self.tutto_il_pc,
            text="Consenti la lettura a tutti gli utenti di questo PC "
                 "(necessario se il servizio gira con un altro utente)",
        ).grid(row=riga, column=0, sticky="w", pady=(0, 8))
        riga += 1

        pulsanti = ttk.Frame(principale)
        pulsanti.grid(row=riga, column=0, sticky="ew")
        ttk.Button(pulsanti, text="Importa config.txt...", command=self._importa).pack(side="left")
        ttk.Button(pulsanti, text="Chiudi", command=self._chiudi).pack(side="right")
        ttk.Button(pulsanti, text="Salva", command=self._salva).pack(side="right", padx=(0, 6))
        riga += 1

        self.stato = ttk.Label(principale, text="", foreground="gray")
        self.stato.grid(row=riga, column=0, sticky="w", pady=(8, 0))

    def _crea_campo(self, genitore, riga, campo, etichetta, tipo):
        if tipo == "spunta":
            ttk.Checkbutton(genitore, text=etichetta, variable=self.variabili[campo], onvalue="1",
                            offvalue="").grid(row=riga, column=1, sticky="w", pady=2)
            return
        ttk.Label(genitore, text=etichetta).grid(row=riga, column=0, sticky="w", padx=(0, 8), pady=2)
        if tipo == "stampante":
            # Si può anche scrivere il nome; l'elenco si legge solo quando si apre la tendina
            voce = ttk.Combobox(genitore, textvariable=self.variabili[campo], width=48)
            voce.configure(postcommand=lambda: voce.configure(values=[""] + self._stampanti()))
        elif tipo == "testo_lungo":
            voce = self._casella_lunga(genitore, self.variabili[campo])
        else:
            voce = ttk.Entry(genitore, textvariable=self.variabili[campo], width=50,
                             show="•" if tipo == "segreto" else "")
        voce.grid(row=riga, column=1, sticky="ew", pady=2)
        if tipo == "cartella":
            ttk.Button(genitore, text="Sfoglia...", command=lambda: self._scegli_cartella(campo)) \
                .grid(row=riga, column=2, padx=(6, 0))
        elif tipo == "file":
            ttk.Button(genitore, text="Sfoglia...", command=lambda: self._scegli_file_log(campo)) \
                .grid(row=riga, column=2, padx=(6, 0))
        elif tipo == "immagine":
            ttk.Button(genitore, text="Sfoglia...", command=lambda: self._scegli_immagine(campo)) \
                .grid(row=riga, column=2, padx=(6, 0))
        elif tipo == "segreto":
            mostra = tk.BooleanVar(value=False)
            ttk.Checkbutton(genitore, text="Mostra", variable=mostra,
                            command=lambda: voce.configure(show="" if mostra.get() else "•")) \
                .grid(row=riga, column=2, padx=(6, 0))

    def _scegli_cartella(self, campo):
        iniziale = self.variabili[campo].get()
        cartella = filedialog.askdirectory(parent=self, initialdir=iniziale if os.path.isdir(iniziale) else None)
        if cartella:
            # Il formato del sistema (con "\" su Windows, "/" su Linux), come li compone il servizio
            self.variabili[campo].set(os.path.normpath(cartella))

    def _casella_lunga(self, genitore, variabile):
        """Casella di più righe legata alla variabile del campo (in tutti e due i sensi)."""
        casella = tk.Text(genitore, height=3, width=50, wrap="word", undo=True, font=("Segoe UI", 9))

        def da_casella(evento):
            if casella.edit_modified():
                casella.edit_modified(False)
                variabile.set(casella.get("1.0", "end-1c"))

        def da_variabile(*argomenti):
            if casella.get("1.0", "end-1c") != variabile.get():
                casella.delete("1.0", "end")
                casella.insert("1.0", variabile.get())
                casella.edit_modified(False)
        casella.bind("<<Modified>>", da_casella)
        variabile.trace_add("write", da_variabile)
        return casella

    def _scegli_immagine(self, campo):
        attuale = self.variabili[campo].get()
        percorso = filedialog.askopenfilename(
            parent=self, title="Immagine della firma", initialdir=os.path.dirname(attuale) if attuale else None,
            filetypes=[("Immagini", "*.png *.jpg *.jpeg *.gif"), ("Tutti i file", "*.*")])
        if percorso:
            self.variabili[campo].set(os.path.normpath(percorso))

    def _scegli_file_log(self, campo):
        attuale = self.variabili[campo].get()
        percorso = filedialog.asksaveasfilename(
            parent=self, title="File di log", defaultextension=".log", confirmoverwrite=False,
            initialdir=os.path.dirname(attuale) if attuale else config_sicuro.CARTELLA_PROGRAMMA,
            initialfile=os.path.basename(attuale) if attuale else "FileGUY.log",
            filetypes=[("File di log", "*.log"), ("Tutti i file", "*.*")])
        if percorso:
            self.variabili[campo].set(os.path.normpath(percorso))

    def _stampanti(self):
        try:
            from stampa_pdf import elenca_stampanti
            return elenca_stampanti()
        except Exception as e:
            self._messaggio("Impossibile leggere l'elenco delle stampanti: " + str(e))
            return []

    # --- Dati --------------------------------------------------------------

    def _valori(self):
        return {campo: var.get().strip() for campo, var in self.variabili.items()}

    def _imposta(self, config):
        for campo, var in self.variabili.items():
            var.set(config.get(campo, CAMPI[campo]))

    def _carica_iniziale(self):
        if config_sicuro.config_esiste():
            try:
                self._imposta(config_sicuro.carica_config())
                self.salvato = self._valori()
                self._messaggio("Configurazione caricata da " + config_sicuro.FILE_CONFIG)
                return
            except ConfigError as e:
                messagebox.showerror("Configurazione", str(e), parent=self)
        self._imposta(CAMPI)
        self.salvato = self._valori()
        if os.path.isfile(config_sicuro.FILE_CONFIG_VECCHIO):
            if messagebox.askyesno("Configurazione",
                                   "Non esiste ancora una configurazione cifrata, ma è presente il vecchio "
                                   "config.txt.\n\nVuoi importarlo?", parent=self):
                self._importa(config_sicuro.FILE_CONFIG_VECCHIO)
                return
        self._messaggio("Nuova configurazione: compila i campi e premi Salva")

    def _importa(self, percorso=None):
        if percorso is None:
            percorso = filedialog.askopenfilename(
                parent=self, title="Importa il vecchio config.txt", initialdir=config_sicuro.CARTELLA_PROGRAMMA,
                filetypes=[("File di testo", "*.txt"), ("Tutti i file", "*.*")])
            if not percorso:
                return
        try:
            vecchio = config_sicuro.importa_config_vecchio(percorso)
        except OSError as e:
            messagebox.showerror("Importa", "Impossibile leggere " + percorso + ":\n" + str(e), parent=self)
            return
        # Mantiene i campi che il vecchio file non aveva (SMTP, WhatsApp...)
        attuali = self._valori()
        for campo in config_sicuro.RIGHE_CONFIG_VECCHIO.values():
            attuali[campo] = vecchio[campo]
        self._imposta(attuali)
        self._percorso_importato = percorso
        self._messaggio("Importato " + percorso + ": controlla i dati e premi Salva")

    def _valida(self, config):
        errori = []
        for campo in ("cartella_lavoro", "cartella_output"):
            if config[campo] == "":
                errori.append("La " + campo.replace("_", " ") + " è obbligatoria")
        for campo in ("cartella_lavoro", "cartella_template", "cartella_output"):
            if config[campo] != "" and not os.path.isdir(config[campo]):
                errori.append("La cartella non esiste: " + config[campo])
        if config["file_log"] != "" and not os.path.isdir(os.path.dirname(os.path.abspath(config["file_log"]))):
            errori.append("La cartella del file di log non esiste: " + os.path.dirname(config["file_log"]))
        if config["email_firma_immagine"] != "" and not os.path.isfile(config["email_firma_immagine"]):
            errori.append("L'immagine della firma non esiste: " + config["email_firma_immagine"])
        if not config["email_smtp_porta"].isdigit():
            errori.append("La porta SMTP deve essere un numero")
        return errori

    def _salva(self):
        config = self._valori()
        errori = self._valida(config)
        if errori:
            if not messagebox.askyesno("Salva", "Attenzione:\n\n- " + "\n- ".join(errori)
                                       + "\n\nSalvare comunque?", parent=self):
                return False
        try:
            config_sicuro.salva_config(config, tutto_il_pc=self.tutto_il_pc.get())
        except OSError as e:
            messagebox.showerror("Salva", "Impossibile salvare la configurazione:\n" + str(e), parent=self)
            return False
        self.salvato = config
        self._messaggio("Configurazione cifrata salvata in " + config_sicuro.FILE_CONFIG)
        self._proponi_eliminazione_vecchio()
        return True

    def _proponi_eliminazione_vecchio(self):
        # Il vecchio config.txt contiene le password in chiaro: dopo l'importazione va eliminato
        percorso = self._percorso_importato
        if percorso is None or not os.path.isfile(percorso):
            return
        if messagebox.askyesno("Vecchio config.txt",
                               "La configurazione è ora cifrata.\n\nIl file " + percorso
                               + " contiene ancora le password in chiaro. Vuoi eliminarlo?", parent=self):
            try:
                os.remove(percorso)
                self._messaggio("Configurazione salvata, " + percorso + " eliminato")
            except OSError as e:
                messagebox.showerror("Elimina", "Impossibile eliminare il file:\n" + str(e), parent=self)
        self._percorso_importato = None

    def _prova_whatsapp(self):
        config = self._valori()
        if config["whatsapp_token"] == "" or config["whatsapp_phone_number_id"] == "":
            messagebox.showwarning("WhatsApp", "Inserisci prima Access token e Phone number ID", parent=self)
            return
        numero = simpledialog.askstring(
            "Messaggio di prova",
            "Numero del destinatario con prefisso internazionale (es. 393331234567).\n"
            "Verrà inviato il template \"hello_world\" fornito da Meta.", parent=self)
        if not numero:
            return
        try:
            from whatsapp_meta import WhatsApp, WhatsAppError
            import requests
        except ImportError as e:
            messagebox.showerror("WhatsApp", "Libreria mancante: " + str(e) + "\nInstalla con: pip install requests",
                                 parent=self)
            return
        self.configure(cursor="watch")
        self.update_idletasks()
        try:
            messenger = WhatsApp(config["whatsapp_token"], config["whatsapp_phone_number_id"],
                                 config["whatsapp_versione_api"])
            id_messaggio = messenger.invia_template(numero, "hello_world", "en_US")
            messagebox.showinfo("WhatsApp", "Messaggio inviato.\nID: " + id_messaggio, parent=self)
        except (WhatsAppError, requests.RequestException) as e:
            messagebox.showerror("WhatsApp", str(e), parent=self)
        finally:
            self.configure(cursor="")

    def _messaggio(self, testo):
        self.stato.configure(text=testo)

    def _chiudi(self):
        if self._valori() != self.salvato:
            risposta = messagebox.askyesnocancel("Chiudi", "Ci sono modifiche non salvate. Salvarle?", parent=self)
            if risposta is None:
                return
            if risposta and not self._salva():
                return
        self.destroy()


if __name__ == "__main__":
    avvia_finestra(ConfigManager).mainloop()
