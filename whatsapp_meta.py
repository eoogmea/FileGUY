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


class WhatsApp:
    def __init__(self, token, phone_number_id, versione_api="v23.0", timeout=60):
        if token == "" or phone_number_id == "":
            raise WhatsAppError("Token o Phone number ID di WhatsApp non configurati")
        self.base_url = GRAPH_URL + "/" + versione_api + "/" + phone_number_id
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers["Authorization"] = "Bearer " + token

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
