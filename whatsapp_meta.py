"""Invio di messaggi WhatsApp tramite la Cloud API di Meta (graph.facebook.com).

Servono, da https://developers.facebook.com (app di tipo Business con il prodotto WhatsApp):
- un access token (meglio quello permanente di un "utente di sistema" del Business Manager)
- il "Phone number ID" del numero mittente

Attenzione alla regola delle 24 ore di Meta: un messaggio libero (testo o documento) arriva
solo se il destinatario ha scritto al numero aziendale nelle ultime 24 ore. Altrimenti bisogna
usare un template approvato su WhatsApp Manager.
"""
import mimetypes
import os
import re

import requests

GRAPH_URL = "https://graph.facebook.com"


class WhatsAppError(Exception):
    pass


def normalizza_numero(numero):
    """Porta il numero nel formato richiesto da Meta: solo cifre con prefisso internazionale."""
    n = numero.strip()
    for c in " -./()":
        n = n.replace(c, "")
    if n.startswith("+"):
        n = n[1:]
    elif n.startswith("00"):
        n = n[2:]
    if not n.isdigit():
        raise WhatsAppError("Numero WhatsApp non valido: " + numero)
    return n


def leggi_destinatario(linea):
    """Numero, template e lingua dalla linea 7 del .gui: "whatsapp <numero> [<template> [<lingua>]]".

    La parola "whatsapp" iniziale si può omettere. Il numero può essere scritto con spazi.
    """
    parti = linea.split()
    if parti and parti[0].lower() == "whatsapp":
        parti = parti[1:]
    numero = ""
    while parti and parti[0].strip("+-./()0123456789") == "":
        numero += parti.pop(0)
    return numero, parti[0] if parti else "", parti[1] if len(parti) > 1 else ""


def leggi_riga_template(riga):
    """(nome, lingua) da una riga "template: <nome> [<lingua>]", altrimenti None.

    Il nome si può scrivere anche tra virgolette: template: "fattura" it
    """
    testo = riga.strip()
    if not testo.lower().startswith(PREFISSO_TEMPLATE):
        return None
    parti = testo[len(PREFISSO_TEMPLATE):].replace('"', " ").split()
    return (parti[0] if parti else "", parti[1] if len(parti) > 1 else "")


PREFISSO_TEMPLATE = "template:"
# Parametri {{1}}, {{2}}... nel testo di un template
SEGNAPOSTO_TEMPLATE = re.compile(r"\{\{\s*(\d+)\s*\}\}")


def numero_parametri(testo):
    numeri = [int(n) for n in SEGNAPOSTO_TEMPLATE.findall(testo or "")]
    return max(numeri) if numeri else 0


class WhatsApp:
    def __init__(self, token, phone_number_id, versione_api="v23.0", timeout=60):
        if token == "" or phone_number_id == "":
            raise WhatsAppError("Token o Phone number ID di WhatsApp non configurati")
        self.graph_url = GRAPH_URL + "/" + versione_api
        self.base_url = self.graph_url + "/" + phone_number_id
        self.token = token
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers["Authorization"] = "Bearer " + token

    # --- Template (WhatsApp Manager) -------------------------------------------

    def _dati_token(self):
        risposta = self.session.get(self.graph_url + "/debug_token",
                                    params={"input_token": self.token, "access_token": self.token},
                                    timeout=self.timeout)
        return self._controlla(risposta).get("data", {})

    def account_business(self):
        """L'id dell'account WhatsApp Business (WABA) a cui il token dà accesso.

        Lo ricava dai permessi del token; serve per leggere e creare i template.
        """
        for permesso in self._dati_token().get("granular_scopes", []):
            if permesso.get("scope") == "whatsapp_business_management" and permesso.get("target_ids"):
                return permesso["target_ids"][0]
        raise WhatsAppError("Il token non dà accesso a un account WhatsApp Business (serve il permesso "
                            "whatsapp_business_management): scrivi l'ID dell'account WhatsApp Business in "
                            "ConfigManager")

    def elenco_template(self, account):
        """I template dell'account: lista di dict con nome, lingua, stato, categoria, testo
        del corpo, tipo di intestazione (es. DOCUMENT, "" se non c'è) e numero di parametri."""
        risultato = []
        url = self.graph_url + "/" + account + "/message_templates"
        parametri = {"fields": "name,language,status,category,components", "limit": 100}
        while url:
            dati = self._controlla(self.session.get(url, params=parametri, timeout=self.timeout))
            for t in dati.get("data", []):
                componenti = {c.get("type", "").upper(): c for c in t.get("components", [])}
                corpo = componenti.get("BODY", {}).get("text", "")
                risultato.append({
                    "nome": t.get("name", ""), "lingua": t.get("language", ""), "stato": t.get("status", ""),
                    "categoria": t.get("category", ""), "testo": corpo,
                    "intestazione": componenti.get("HEADER", {}).get("format", ""),
                    "parametri": numero_parametri(corpo),
                })
            # Le pagine successive hanno già i parametri nell'indirizzo
            url = dati.get("paging", {}).get("next")
            parametri = None
        return sorted(risultato, key=lambda t: (t["nome"], t["lingua"]))

    def _carica_esempio(self, percorso):
        """Carica un file di esempio per l'intestazione di un nuovo template (Resumable Upload
        API dell'app); restituisce l'handle da mettere nell'esempio."""
        app = self._dati_token().get("app_id")
        if not app:
            raise WhatsAppError("Non riesco a sapere l'app del token per caricare il documento di esempio")
        tipo_mime = mimetypes.guess_type(percorso)[0] or "application/pdf"
        with open(percorso, "rb") as f:
            contenuto = f.read()
        sessione = self._controlla(self.session.post(
            self.graph_url + "/" + app + "/uploads",
            params={"file_name": os.path.basename(percorso), "file_length": len(contenuto), "file_type": tipo_mime},
            timeout=self.timeout))["id"]
        risposta = requests.post(self.graph_url + "/" + sessione, data=contenuto, timeout=self.timeout,
                                 headers={"Authorization": "OAuth " + self.token, "file_offset": "0"})
        return self._controlla(risposta)["h"]

    def crea_template(self, account, nome, lingua, categoria, testo, esempi=(), documento_esempio=None,
                      piede=""):
        """Crea un template (va poi approvato da Meta: di solito minuti, a volte ore).

        testo: corpo con i parametri {{1}}, {{2}}...; esempi: un valore di esempio per ogni parametro
        (Meta li vuole per l'approvazione). documento_esempio: con un percorso il template ha
        un'intestazione Documento (per mandare il PDF creato). Restituisce (id, stato).
        """
        componenti = []
        if documento_esempio:
            componenti.append({"type": "HEADER", "format": "DOCUMENT",
                               "example": {"header_handle": [self._carica_esempio(documento_esempio)]}})
        corpo = {"type": "BODY", "text": testo}
        if numero_parametri(testo):
            corpo["example"] = {"body_text": [list(esempi)]}
        componenti.append(corpo)
        if piede.strip():
            componenti.append({"type": "FOOTER", "text": piede.strip()})
        risposta = self.session.post(self.graph_url + "/" + account + "/message_templates", timeout=self.timeout,
                                     json={"name": nome, "language": lingua, "category": categoria,
                                           "components": componenti})
        dati = self._controlla(risposta)
        return dati.get("id", ""), dati.get("status", "")

    # --- Messaggi ----------------------------------------------------------------

    def _controlla(self, risposta):
        try:
            dati = risposta.json()
        except ValueError:
            dati = {}
        if not risposta.ok or "error" in dati:
            errore = dati.get("error", {})
            messaggio = errore.get("message", risposta.text)
            dettagli = errore.get("error_data", {}).get("details", "")
            raise WhatsAppError("Errore API WhatsApp " + str(risposta.status_code) + " (codice "
                                + str(errore.get("code", "?")) + "): " + messaggio
                                + (" - " + dettagli if dettagli else ""))
        return dati

    def _invia(self, destinatario, tipo, contenuto):
        payload = {
            "messaging_product": "whatsapp",
            "recipient_type": "individual",
            "to": normalizza_numero(destinatario),
            "type": tipo,
            tipo: contenuto,
        }
        risposta = self.session.post(self.base_url + "/messages", json=payload, timeout=self.timeout)
        # Restituisce l'id del messaggio (wamid...)
        return self._controlla(risposta)["messages"][0]["id"]

    def carica_media(self, percorso):
        """Carica un file sui server di Meta e restituisce l'id da usare nei messaggi."""
        tipo_mime = mimetypes.guess_type(percorso)[0] or "application/octet-stream"
        with open(percorso, "rb") as f:
            risposta = self.session.post(
                self.base_url + "/media",
                data={"messaging_product": "whatsapp", "type": tipo_mime},
                files={"file": (os.path.basename(percorso), f, tipo_mime)},
                timeout=self.timeout,
            )
        return self._controlla(risposta)["id"]

    def invia_testo(self, destinatario, testo):
        return self._invia(destinatario, "text", {"body": testo, "preview_url": False})

    def invia_documento(self, destinatario, percorso, didascalia=""):
        documento = {"id": self.carica_media(percorso), "filename": os.path.basename(percorso)}
        if didascalia != "":
            documento["caption"] = didascalia
        return self._invia(destinatario, "document", documento)

    def invia_template(self, destinatario, nome, lingua, parametri=(), documento=None):
        """Invia un template approvato.

        parametri: testi che sostituiscono {{1}}, {{2}}, ... nel corpo del template
        documento: percorso di un file da usare come intestazione (il template deve avere
                   un'intestazione di tipo Documento)
        """
        componenti = []
        if documento:
            componenti.append({
                "type": "header",
                "parameters": [{
                    "type": "document",
                    "document": {"id": self.carica_media(documento), "filename": os.path.basename(documento)},
                }],
            })
        if parametri:
            componenti.append({
                "type": "body",
                "parameters": [{"type": "text", "text": p} for p in parametri],
            })
        template = {"name": nome, "language": {"code": lingua}}
        if componenti:
            template["components"] = componenti
        return self._invia(destinatario, "template", template)
