# Dashboard dichiarazioni di origine

## Avvio

Usare `DASHBOARD_DICH_ORIGINE.bat`.

La dashboard usa il runtime Python portatile incluso nella cartella `runtime`: non e
necessario installare Python sul computer. All'avvio viene sincronizzata nella cache
locale `%LOCALAPPDATA%\NordLaser\DichiarazioniOrigine` per evitare limitazioni
legate all'esecuzione diretta dalla rete.

Si apre automaticamente su una porta locale libera senza lasciare aperta la
finestra CMD. La cartella puo essere avviata direttamente dal percorso di rete.
Il log diagnostico si trova in
`%LOCALAPPDATA%\NordLaser\DichiarazioniOrigine\dashboard.log`.

## Clienti

L'elenco clienti filtrabile e nel file `clients.json`.

Formato:

```json
[
  {
    "code": "CODICE_CLIENTE",
    "name": "Ragione sociale cliente",
    "address": ""
  }
]
```

La ricerca cliente filtra sia per codice sia per ragione sociale.

## Import codici

Incollare una riga per articolo:

```text
H9313468 COMPRESSOR CANOPY FIXING HANGER PAINTED
H9313469;DESCRIZIONE ARTICOLO
H9313470	DESCRIZIONE ARTICOLO
```

La colonna "Gruppo paese" viene compilata nel PDF con un solo asterisco per riga; il dettaglio dei gruppi resta sotto la tabella.

## PDF

Il nome file viene generato automaticamente:

```text
AAAAMMGG_HHMMSS - Dichiarazione origine CLIENTE.pdf
AAAAMMGG_HHMMSS - Origin declaration CLIENTE.pdf
```

La cartella di destinazione puo essere digitata oppure selezionata con il pulsante
cartella accanto al campo.

## Registro stampe

Ogni PDF generato viene registrato nel file condiviso
`registro_attestazioni.csv`, nella cartella NAS della dashboard. Il registro e
unico per tutte le postazioni e contiene data e ora, utente Windows, lingua,
cliente, validita, numero di articoli e percorso del PDF.

Il pulsante `Log stampe`, accanto ai pulsanti di generazione PDF, apre la tabella
completa. Dalla finestra del registro sono disponibili le esportazioni Excel
`.xlsx` e testo `.txt`.
