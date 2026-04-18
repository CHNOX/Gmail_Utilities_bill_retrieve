# Gmail Utilities Bill Retrieve

Script Python che recupera tutte le email ricevute da un mittente specifico (**Enel Energia** – `enelenergia@enel.sandsiv.com`) tramite l'API Gmail e le esporta in un file Excel dettagliato.

## Funzionalità

- Autenticazione OAuth2 con Google (flusso locale, token salvato per i successivi accessi)
- Ricerca di tutte le email dal mittente `enelenergia@enel.sandsiv.com`
- Ordinamento per data decrescente (più recenti prima)
- Export Excel con:
  - Numero progressivo
  - Data e ora di ricezione
  - Oggetto dell'email
  - Anteprima del testo
- Intestazioni con tema cromatico Enel, righe alternate, filtri automatici e riga bloccata

## Prerequisiti

- Python 3.7+
- Dipendenze:

```bash
pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib openpyxl
```

## Configurazione Google Cloud

1. Vai su [https://console.cloud.google.com/](https://console.cloud.google.com/)
2. Crea o seleziona un progetto
3. Abilita **Gmail API** (API e servizi › Libreria)
4. Crea credenziali › **ID client OAuth** › tipo *App desktop*
5. Scarica il JSON e rinominalo **`credentials.json`** nella stessa cartella dello script
6. Nella schermata consenso OAuth aggiungi lo scope:  
   `https://www.googleapis.com/auth/gmail.readonly`  
   e il tuo indirizzo Gmail come *utente di test*

> `credentials.json` e `token.json` sono esclusi dal repository tramite `.gitignore`.

## Utilizzo

```bash
python Gmail_Utilities_bill_retrieve.py
```

La prima volta si aprirà il browser per autorizzare l'accesso.  
Verrà generato un file `enel_emails_YYYYMMDD_HHMMSS.xlsx` nella stessa cartella.

## Output

| Colonna   | Contenuto                        |
|-----------|----------------------------------|
| #         | Numero progressivo               |
| Data      | Data e ora ricezione (gg/mm/aaaa)|
| Oggetto   | Oggetto dell'email               |
| Anteprima | Testo breve del corpo email      |
