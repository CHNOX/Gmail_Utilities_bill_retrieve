# Gmail Utilities Bill Retrieve

Recupera **tutte le email** da uno o più mittenti configurati in `settings.json` e genera un file **Excel** con data, oggetto, anteprima e — per i mittenti che lo prevedono — i valori estratti automaticamente dal corpo dell'email (es. *TOTALE DA PAGARE*).

I mittenti e le chiavi di estrazione si configurano liberamente senza toccare il codice.

## Cosa produce

Il file Excel generato (`enel_emails_YYYYMMDD_HHMMSS.xlsx`) contiene le colonne fisse più una colonna aggiuntiva per ogni chiave di estrazione definita in `settings.json`:

| # | Data | Mittente | Oggetto | Anteprima | TOTALE DA PAGARE |
|---|------|----------|---------|-----------|-----------------|
| 1 | 15/04/2026 08:32 | noreply.enelenergia@enel.com | La tua bolletta di aprile | Gentile cliente… | € 87,43 |
| 2 | 14/03/2026 09:10 | enelenergia@enel.sandsiv.com | Riepilogo consumi marzo | In allegato trovi… | |
| … | … | … | … | … | … |

- Ordinate dalla email più recente alla meno recente
- Colonne di estrazione aggiunte dinamicamente in base a `settings.json`
- Prima riga bloccata e filtri automatici attivi
- Tema cromatico Enel (rosso), righe alternate per leggibilità

---

## Requisiti

- **Python 3.7 o superiore**
- Un account **Google / Gmail**
- Accesso a **Google Cloud Console** (gratuito)

---

## Configurazione settings.json

Il file `settings.json` definisce quali mittenti monitorare e quali valori estrarre dal corpo di ogni email.

```json
{
  "senders": [
    {
      "email": "enelenergia@enel.sandsiv.com",
      "extract_keys": []
    },
    {
      "email": "noreply.enelenergia@enel.com",
      "extract_keys": ["TOTALE DA PAGARE"]
    }
  ]
}
```

**Campi:**

| Campo | Tipo | Descrizione |
|---|---|---|
| `email` | stringa | Indirizzo del mittente da cercare |
| `extract_keys` | array di stringhe | Etichette da cercare nel corpo dell'email; lascia `[]` per non estrarre nulla |

**Regole:**
- Ogni chiave in `extract_keys` genera una colonna aggiuntiva nel file Excel
- Se la stessa chiave appare in più mittenti, viene creata **una sola colonna** condivisa
- Per i mittenti senza chiavi lo script usa solo i metadati (più veloce)
- Puoi aggiungere quanti mittenti e chiavi vuoi senza modificare il codice

**Esempio con più mittenti e chiavi:**

```json
{
  "senders": [
    {
      "email": "enelenergia@enel.sandsiv.com",
      "extract_keys": []
    },
    {
      "email": "noreply.enelenergia@enel.com",
      "extract_keys": ["TOTALE DA PAGARE"]
    },
    {
      "email": "fatture@altrofornitore.it",
      "extract_keys": ["TOTALE DA PAGARE", "NUMERO FATTURA"]
    }
  ]
}
```

Questo produrrebbe un Excel con le colonne: `#`, `Data`, `Mittente`, `Oggetto`, `Anteprima`, `TOTALE DA PAGARE`, `NUMERO FATTURA`.

---

## Installazione

### 1. Installa Python

Scarica e installa Python da [python.org](https://www.python.org/downloads/) (versione 3.7+).

Verifica l'installazione:
```bash
python --version
```

### 2. Clona il repository

```bash
git clone https://github.com/CHNOX/Gmail_Utilities_bill_retrieve.git
cd Gmail_Utilities_bill_retrieve
```

### 3. Installa le dipendenze

Con il file `requirements.txt` incluso nel progetto:

```bash
pip install -r requirements.txt
```

Oppure manualmente:

```bash
pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib openpyxl
```

**Librerie installate:**

| Libreria | Versione minima | Scopo |
|---|---|---|
| `google-api-python-client` | >= 2.0.0 | Client per le API di Google |
| `google-auth-httplib2` | >= 0.1.0 | Trasporto HTTP per l'autenticazione |
| `google-auth-oauthlib` | >= 0.5.0 | Flusso OAuth2 per desktop |
| `openpyxl` | >= 3.0.0 | Generazione file Excel `.xlsx` |

---

## Configurazione Google Cloud Console

> Questa è la parte più lunga ma va fatta **una sola volta**.
> Tempo stimato: circa 10 minuti.

### Passo 1 — Accedi a Google Cloud Console

Vai su [console.cloud.google.com](https://console.cloud.google.com) con il tuo account Google.

![Google Cloud Console - Home](docs/screenshots/01_home.png)

---

### Passo 2 — Crea un nuovo progetto

1. Clicca sul selettore di progetto in alto (accanto al logo Google Cloud)

   ![Seleziona progetto](docs/screenshots/02_seleziona_progetto.png)

2. Nella finestra che si apre, clicca **"Nuovo progetto"** in alto a destra

   ![Nuovo progetto](docs/screenshots/03_nuovo_progetto.png)

3. Inserisci un nome (es. `API Gmail`) e clicca **"Crea"**

   ![Crea progetto](docs/screenshots/04_crea_progetto.png)

4. Attendi qualche secondo: il progetto verrà creato e selezionato automaticamente

   ![Progetto creato](docs/screenshots/05_progetto_creato.png)

---

### Passo 3 — Abilita la Gmail API

1. Dal menu laterale vai su **"API e servizi"** → **"Libreria"**

2. Nella libreria troverai la **Gmail API** nella sezione Google Workspace (oppure cercala nella barra di ricerca)

   ![Libreria API - Gmail API](docs/screenshots/06_libreria_gmail_api.png)

3. Clicca su **"Gmail API"** → poi clicca **"Abilita"**

   ![Abilita Gmail API](docs/screenshots/07_abilita_gmail_api.png)

4. Verrai reindirizzato alla pagina dei dettagli dell'API, dove vedrai lo stato **"Abilitata"**

   ![Gmail API abilitata](docs/screenshots/08_gmail_api_abilitata.png)

---

### Passo 4 — Configura la schermata di consenso OAuth

Questa schermata viene mostrata all'utente quando lo script richiede l'accesso a Gmail.

1. Dal menu laterale vai su **"API e servizi"** → **"Schermata consenso OAuth"**
   *(oppure naviga in "Google Auth Platform" → "Panoramica")*

2. Se non hai ancora configurato nulla, clicca **"Inizia"**

   ![OAuth - Panoramica](docs/screenshots/09_oauth_panoramica.png)

3. **Sezione "Informazioni sull'app"**: inserisci:
   - **Nome applicazione**: es. `API Gmail Read`
   - **Email per assistenza utenti**: la tua email Gmail

   Clicca **"Avanti"**

   ![OAuth - Info app](docs/screenshots/10_oauth_info_app.png)

4. **Sezione "Pubblico"**: seleziona **"Esterno"**

   > Con "Esterno" l'app funziona in modalità *Test* e solo gli utenti che aggiungi esplicitamente potranno usarla. È la scelta corretta per uso personale.

   Clicca **"Avanti"**

   ![OAuth - Pubblico Esterno](docs/screenshots/11_oauth_esterno.png)

5. **Sezione "Dati di contatto"**: inserisci la tua email Gmail e clicca **"Avanti"**

   ![OAuth - Dati contatto](docs/screenshots/12_oauth_dati_contatto.png)

6. Clicca **"Crea"** per completare la configurazione di base

   ![OAuth - Fine](docs/screenshots/13_oauth_fine.png)

---

### Passo 5 — Aggiungi lo scope (permesso di lettura Gmail)

1. Nel menu laterale di Google Auth Platform, clicca **"Accesso ai dati"**

2. Clicca **"Aggiungi o rimuovi ambiti"**

   ![Accesso ai dati](docs/screenshots/14_accesso_dati.png)

3. Nel pannello che si apre, cerca e seleziona:

   ```
   https://www.googleapis.com/auth/gmail.readonly
   ```

   Metti la spunta su **Gmail API** → `.../auth/gmail.readonly` → clicca **"Aggiorna"**

   ![Seleziona scope gmail.readonly](docs/screenshots/15_scope_gmail_readonly.png)

> **Nota:** questo scope garantisce accesso in **sola lettura**. Lo script non può modificare, spostare né eliminare alcuna email.

---

### Passo 6 — Aggiungi te stesso come utente di test

1. Nel menu laterale clicca **"Pubblico"**

2. Nella sezione **"Utenti di prova"**, clicca **"+ Add users"**

   ![Pubblico - Add users](docs/screenshots/16_pubblico_add_users.png)

3. Inserisci il tuo indirizzo Gmail e clicca **"Salva"**

   ![Aggiungi utente di test](docs/screenshots/17_aggiungi_utente.png)

> Finché l'app è in modalità *Test*, solo gli indirizzi aggiunti qui potranno autenticarsi.

---

### Passo 7 — Crea le credenziali OAuth (file credentials.json)

1. Nel menu laterale clicca **"Client"**

2. Clicca **"+ Crea client"**

   ![Client - Crea client](docs/screenshots/18_client_crea.png)

3. Nella pagina "Crea ID client OAuth":
   - **Tipo di applicazione**: seleziona **"Applicazione desktop"**
   - **Nome**: es. `Client_Desktop`
   - Clicca **"Crea"**

   ![Crea OAuth Client](docs/screenshots/19_crea_oauth_client.png)

4. Si aprirà un popup con le credenziali create. Clicca **"Scarica JSON"**

   ![Client creato - Scarica JSON](docs/screenshots/20_client_creato_scarica_json.png)

   > **Attenzione:** salva il `Client secret` o scarica subito il JSON — non sarà più visibile dopo aver chiuso questa finestra.

5. **Rinomina** il file scaricato da `client_secret_XXXXX.json` in:
   ```
   credentials.json
   ```

6. **Copia** `credentials.json` nella stessa cartella dello script `Gmail_Utilities_bill_retrieve.py`

---

## Utilizzo

### Prima esecuzione

```bash
python Gmail_Utilities_bill_retrieve.py
```

1. Lo script legge `settings.json` e mostra i mittenti e le chiavi configurati
2. Si aprirà **automaticamente il browser** con la schermata di consenso Google
3. Google potrebbe mostrare l'avviso **"App non verificata"** — è normale per app personali in fase di test:
   - Clicca **"Avanzate"**
   - Clicca **"Vai a [nome app] (non sicuro)"**
4. Concedi i permessi di lettura email
5. Il browser mostrerà `The authentication flow has completed` — puoi chiuderlo
6. Lo script recupera le email e mostra il progresso nel terminale

### Esecuzioni successive

Il token di accesso viene salvato in `token.json`. Lo script si avvierà direttamente senza aprire il browser.

### Output nel terminale

```
============================================================
   Gmail Utilities Bill Retrieve
   • enelenergia@enel.sandsiv.com  [—]
   • noreply.enelenergia@enel.com  [TOTALE DA PAGARE]
============================================================

Account: tuo@gmail.com

Ricerca email da: enelenergia@enel.sandsiv.com
  Trovati 18 messaggi — scaricamento dettagli...
  18/18 elaborati

Ricerca email da: noreply.enelenergia@enel.com
  Chiavi da estrarre: TOTALE DA PAGARE
  Trovati 12 messaggi — scaricamento dettagli...
  12/12 elaborati

Recupero completato: 30 email totali.

File Excel salvato: enel_emails_20260418_103215.xlsx

Prime 10 email (più recenti):
------------------------------------------------------------
   1. [18/04/2026] La tua bolletta di aprile 2026  →  TOTALE DA PAGARE: € 87,43
   2. [15/04/2026] Aggiornamento contatore
   3. [15/03/2026] La tua bolletta di marzo 2026   →  TOTALE DA PAGARE: € 74,10
   ...

File generato: enel_emails_20260418_103215.xlsx
```

---

## Struttura file

```
Gmail_Utilities_bill_retrieve/
├── Gmail_Utilities_bill_retrieve.py   # Script principale — contiene tutta la logica
├── settings.json                      # ← CONFIGURAZIONE: mittenti e chiavi da estrarre
│                                      #   Modifica questo file per aggiungere mittenti/chiavi
├── requirements.txt                   # Elenco delle dipendenze Python da installare con pip
├── README.md                          # Questa guida
│
├── credentials.json                   # ← DA SCARICARE da Google Cloud Console (vedi Passo 7)
│                                      #   Contiene Client ID e Client Secret dell'app OAuth
│                                      #   ⚠ NON committare su GitHub
│
├── token.json                         # ← GENERATO AUTOMATICAMENTE al primo accesso
│                                      #   Contiene il token OAuth dell'utente autenticato
│                                      #   ⚠ NON committare su GitHub
│
└── enel_emails_YYYYMMDD_HHMMSS.xlsx   # ← FILE DI OUTPUT generato ad ogni esecuzione
                                       #   Il nome include data e ora del recupero
```

> **Sicurezza:** `credentials.json` e `token.json` contengono credenziali private.
> Non caricarli mai su GitHub — sono già esclusi tramite `.gitignore`.

---

## Risoluzione problemi

### "File 'credentials.json' non trovato"
Il file `credentials.json` non è nella stessa cartella dello script. Spostalo o verificane il nome.

### "Token scaduto"
Il token viene rinnovato automaticamente. Se l'errore persiste, cancella `token.json` e riesegui.

### "Accesso bloccato: questa app non è verificata"
Segui i passaggi del [Passo 6](#passo-6--aggiungi-te-stesso-come-utente-di-test): assicurati di aver aggiunto il tuo indirizzo Gmail tra gli **utenti di test**.

### "Error 403: access_denied"
L'indirizzo Gmail con cui ti stai autenticando non è nella lista degli utenti di test. Aggiungilo dalla sezione **"Pubblico"** in Google Auth Platform.

### "Nessun mittente configurato in settings.json"
Il file `settings.json` esiste ma l'array `senders` è vuoto. Aggiungi almeno un mittente.

### Nessuna email trovata per un mittente
Verifica che il mittente sia scritto correttamente in `settings.json` e che abbia effettivamente inviato email alla casella autenticata. Puoi cercarlo manualmente in Gmail con `from:indirizzo@esempio.com`.

### La chiave non viene trovata / la colonna è vuota
Il testo estratto dall'email potrebbe avere un formato diverso dal previsto. Lo script usa tre strategie di ricerca progressive (importo con €, numero decimale, primo token). Se nessuna funziona, verifica il layout reale dell'email e aggiusta la chiave in `settings.json` in modo che corrisponda esattamente al testo presente nel corpo.

### Voglio usare un account Gmail diverso
Cancella `token.json` e riesegui lo script: verrà chiesta una nuova autenticazione.

---

## Note tecniche

- Lo script usa **solo** lo scope `gmail.readonly`: non può modificare, spostare né eliminare email
- I dati vengono elaborati **localmente** sul tuo computer, non inviati a server esterni
- L'API Gmail ha un limite di 1 miliardo di unità quota al giorno — ampiamente sufficiente per uso personale
- Lo script processa i messaggi in batch da 500 per ottimizzare le chiamate API
- Per i mittenti **senza** `extract_keys` lo script usa `format=metadata` (solo intestazioni): più veloce e meno quota consumata
- Per i mittenti **con** `extract_keys` lo script usa `format=full` per scaricare il corpo completo e applicare l'estrazione
- L'estrazione di un valore avviene in tre passaggi: ① importo con simbolo €, ② numero decimale generico, ③ primo token dopo la chiave

---

## Licenza

MIT
