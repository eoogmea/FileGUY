# Criticità da verificare con i prossimi file di prova

Punti emersi durante i test del 23/10/2026: da rileggere quando arrivano file veri del gestionale
(fatture, DDT, listini) o si prova su altri computer. Per ognuno: cosa può succedere, come
controllarlo e cosa si potrebbe fare.

## Testi PDF e Word

1. **Formati legati al numero di riga.** Il formato vale per la riga N del testo, qualunque cosa
   contenga. Se una fattura ha meno (o più) righe dell'esempio, tutte le righe dopo prendono il
   formato sbagliato (nella prova con 3 fatture: totali spostati, scadenza non più in rosso).
   - Controllare: le fatture vere hanno sempre la stessa struttura (righe vuote al posto degli
     articoli mancanti, totali sempre alla stessa riga)?
   - Se no: regola per le ultime N righe di ogni pagina (come `u1` di Excel) oppure formato scelto
     da una parola della riga ("TOTALE", "Imponibile").

2. **Salto pagina (`\f`): i formati ripartono dalla prima linea.** Va bene per più fatture in un
   file. Ma un modello fatto su un esempio di 2 pagine diverse (prima pagina con intestazione,
   seconda di continuazione) ora applica alla seconda pagina i formati della prima.
   - Controllare: ci sono documenti di più pagine con pagine diverse tra loro (es. DDT lunghi)?
   - Se sì: un'opzione nel .gui per non ripartire al salto pagina, o formati per "pagina successiva".

3. **`\r` da solo (ritorno carrello senza a capo).** I file del gestionale hanno a fine pagina
   `\r \r\f` (es. `C:\prova\ModelliGUI\prova.txt`). Il programma tratta ogni `\r` da solo come
   un a capo, quindi aggiunge una riga " " in più per pagina. Editor e servizio leggono allo stesso
   modo, quindi i modelli restano coerenti, ma il conteggio delle righe non è quello della stampante.
   - Controllare con un file vero: le righe dopo la prima pagina sono allineate ai formati?
   - Se no: in `piattaforma.leggi_testo` ignorare i `\r` da soli invece di farli diventare a capo
     (cambia il numero di righe: i modelli già fatti su quei file vanno ricontrollati).

4. **Codici all'inizio del file.** `prova.txt` comincia con `\05\r\10` (scritti come testo, forse
   comandi per la stampante: carattere compresso, interlinea…). Ora finiscono nel documento come
   testo, su due righe.
   - Chiedere cosa sono; se sono comandi per la stampante si possono togliere o usare (es. per la
     dimensione del carattere).

5. **Testo sopra la grafica del template.** Nell'esempio `02_fattura_su_template` il testo non è
   allineato a `fat.pdf` e si accavalla. Solo un problema dell'esempio, ma con i template veri va
   controllata la posizione (margini, interlinea di 6 mm, 44 righe per pagina A4).

## Excel

6. **Numeri con un solo punto e senza virgola.** `1.000.000` ora è un milione, ma `1.234` viene
   letto come 1,234 (decimale all'inglese), non come milleduecentotrentaquattro.
   - Controllare come il gestionale scrive le migliaia senza decimali.

7. **Larghezza automatica delle colonne.** È una stima (testo come lo mostra il formato, dimensione
   del carattere, grassetto e corsivo). Con caratteri molto larghi o font sostituiti su Linux
   potrebbe ancora uscire `###`.
   - Controllare: `python test/test_excel_resa.py` su Linux (segnala le celle con `###`) e le
     immagini `test/_run/excel_resa/excel_resa*.png`, comprese le pagine con date e ore, che non
     abbiamo ancora visto.

8. **Testo atteso da LibreOffice non ancora verificato.** In `test/test_excel_resa.py` il testo
   che Calc dovrebbe mostrare (es. `1,23E+03`, `1:30 PM`, `20-ott-2026`) è scritto senza averlo
   provato: alla prima esecuzione su Linux alcune differenze potrebbero essere solo di scrittura e
   non errori veri.

## Stampa

9. **Stampanti PDF virtuali dal servizio.** Con l'azione "stampa" su "Microsoft Print to PDF"
   il servizio (back.exe) fa comparire la finestra "Salva con nome" e resta fermo finché qualcuno
   non risponde: gli altri file della cartella di lavoro aspettano.
   - Se serve stampare su PDF in automatico: un campo in ConfigManager con la cartella dove salvare
     le stampe delle stampanti virtuali (`stampa_pdf` lo supporta già con `file_uscita`).

## Invii

10. **.gui di solo invio senza .txt.** Si elabora appena compare nella cartella di lavoro. Se il
    gestionale scrive prima il .gui e poi il .txt (per un solo invio con il testo), il .txt arriva
    dopo e resta lì.
    - Consigliato: scrivere il .gui con un altro nome (es. `.tmp`) e poi rinominarlo, e per i soli
      invii non mandare il .txt.

11. **WhatsApp.** Il messaggio libero arriva solo se il destinatario ha scritto al numero aziendale
    nelle ultime 24 ore; per gli altri serve un template approvato, con la lingua giusta (errore
    132001 = template non trovato in quella lingua). L'errore "API access blocked" (codice 200)
    vuol dire account o app bloccati su Meta, non un problema del programma.

12. **Firma HTML nelle email.** Provata con Gmail. Da guardare anche in Outlook e su telefono
    (tabelle, immagini, link); la versione solo testo si ricava togliendo i tag.

## Linux

13. **Finestre all'80%.** Fatto senza poterlo vedere su uno schermo Linux: se sono ancora troppo
    grandi o troppo piccole si cambia `SCALA_FINESTRE` in `piattaforma.py`.

## Eseguibili

14. **Exe in C:\prova da ricompilare.** Sono stati creati prima delle correzioni del salto pagina
    e della larghezza delle colonne Excel (commit 8b83bf83): prima dei prossimi test
    `python installa.py --compila` e copiarli in C:\prova.
