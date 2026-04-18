"""
Gmail Utilities Bill Retrieve
==============================
Recupera tutte le email ricevute dai mittenti Enel:
  - enelenergia@enel.sandsiv.com
  - noreply.enelenergia@enel.com  (estrae anche "TOTALE DA PAGARE")

e genera un file Excel con i dettagli (data, mittente, oggetto, anteprima,
importo da pagare).

PREREQUISITI:
1. Python 3.7+
2. Installa le dipendenze:
       pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib openpyxl

3. Vai su https://console.cloud.google.com/
   - Crea un nuovo progetto (o seleziona uno esistente)
   - Abilita "Gmail API" (API e servizi > Libreria > cerca "Gmail API" > Abilita)
   - Vai su "API e servizi > Credenziali"
   - Clicca "Crea credenziali > ID client OAuth"
   - Tipo applicazione: "App desktop"
   - Scarica il file JSON e rinominalo "credentials.json"
   - Mettilo nella stessa cartella di questo script

4. (Solo la prima volta) Vai su "API e servizi > Schermata consenso OAuth"
   - Seleziona "Esterno" e compila i campi obbligatori
   - In "Ambiti" aggiungi: https://www.googleapis.com/auth/gmail.readonly
   - In "Utenti di test" aggiungi il tuo indirizzo Gmail

5. Esegui:
       python Gmail_Utilities_bill_retrieve.py

   La prima volta si aprirà il browser per autorizzare l'accesso.
   Le credenziali vengono salvate in "token.json" per gli accessi successivi.
"""

import base64
import os
import re
import sys
from datetime import datetime
from html.parser import HTMLParser

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

CREDENTIALS_FILE = "credentials.json"
TOKEN_FILE = "token.json"
BATCH_SIZE = 500

SENDER_SANDSIV = "enelenergia@enel.sandsiv.com"
SENDER_NOREPLY = "noreply.enelenergia@enel.com"


# ---------------------------------------------------------------------------
# Autenticazione
# ---------------------------------------------------------------------------

def authenticate():
    """Gestisce l'autenticazione OAuth2 con Google."""
    creds = None

    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("Token scaduto, rinnovo in corso...")
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                print(f"ERRORE: File '{CREDENTIALS_FILE}' non trovato!")
                print("Scaricalo dalla Google Cloud Console (vedi istruzioni nel file).")
                sys.exit(1)
            print("Apertura browser per autorizzazione...")
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())
        print("Autenticazione completata e salvata.")

    return creds


# ---------------------------------------------------------------------------
# Utilità di parsing
# ---------------------------------------------------------------------------

def parse_date(date_str):
    """Converte la stringa data dell'header in oggetto datetime."""
    if not date_str:
        return None
    date_str = re.sub(r'\s*\([^)]*\)\s*$', '', date_str.strip())
    formats = [
        "%a, %d %b %Y %H:%M:%S %z",
        "%a, %d %b %Y %H:%M:%S",
        "%d %b %Y %H:%M:%S %z",
        "%d %b %Y %H:%M:%S",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    return None


class _HTMLStripper(HTMLParser):
    """Estrae solo il testo da una stringa HTML."""
    def __init__(self):
        super().__init__()
        self._parts = []

    def handle_data(self, data):
        self._parts.append(data)

    def get_text(self):
        return " ".join(self._parts)


def _html_to_text(html_content):
    stripper = _HTMLStripper()
    stripper.feed(html_content)
    return stripper.get_text()


def extract_body_text(payload):
    """
    Estrae ricorsivamente il testo dal payload del messaggio Gmail.
    Preferisce text/plain; ricade su text/html se non disponibile.
    """
    mime_type = payload.get("mimeType", "")
    parts = payload.get("parts", [])

    if mime_type == "text/plain":
        raw = payload.get("body", {}).get("data", "")
        if raw:
            return base64.urlsafe_b64decode(raw + "==").decode("utf-8", errors="replace")

    if mime_type == "text/html":
        raw = payload.get("body", {}).get("data", "")
        if raw:
            html = base64.urlsafe_b64decode(raw + "==").decode("utf-8", errors="replace")
            return _html_to_text(html)

    # multipart: cerca prima text/plain, poi text/html tra le parti
    plain_text = ""
    html_text = ""
    for part in parts:
        pt = part.get("mimeType", "")
        if pt == "text/plain":
            raw = part.get("body", {}).get("data", "")
            if raw:
                plain_text = base64.urlsafe_b64decode(raw + "==").decode("utf-8", errors="replace")
        elif pt == "text/html":
            raw = part.get("body", {}).get("data", "")
            if raw:
                html_content = base64.urlsafe_b64decode(raw + "==").decode("utf-8", errors="replace")
                html_text = _html_to_text(html_content)
        elif part.get("parts"):
            # multipart annidato
            nested = extract_body_text(part)
            if nested:
                plain_text = plain_text or nested

    return plain_text or html_text


def extract_totale_da_pagare(text):
    """
    Cerca 'TOTALE DA PAGARE' nel testo e restituisce l'importo associato.
    Formato atteso: es. "123,45 €" oppure "€ 123,45" oppure "123.45€"
    """
    if not text:
        return ""

    # Normalizza spazi multipli
    text = re.sub(r"\s+", " ", text)

    patterns = [
        # "TOTALE DA PAGARE ... € 123,45" oppure "TOTALE DA PAGARE ... 123,45 €"
        r"TOTALE\s+DA\s+PAGARE[^0-9€$]*?([\d]{1,5}[.,][\d]{2})\s*€",
        r"TOTALE\s+DA\s+PAGARE[^0-9€$]*?€\s*([\d]{1,5}[.,][\d]{2})",
        # fallback senza simbolo €
        r"TOTALE\s+DA\s+PAGARE[^0-9]*?([\d]{1,5}[.,][\d]{2})",
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            amount = match.group(1).strip()
            # Normalizza separatore decimale in formato italiano (virgola)
            return f"€ {amount}"

    return ""


# ---------------------------------------------------------------------------
# Recupero email
# ---------------------------------------------------------------------------

def _list_message_ids(service, sender):
    """Restituisce tutti gli ID messaggi del mittente dato."""
    ids = []
    page_token = None
    while True:
        params = {"userId": "me", "maxResults": BATCH_SIZE, "q": f"from:{sender}"}
        if page_token:
            params["pageToken"] = page_token
        results = service.users().messages().list(**params).execute()
        ids.extend(results.get("messages", []))
        page_token = results.get("nextPageToken")
        if not page_token:
            break
    return ids


def fetch_all_emails(service):
    """
    Recupera email da entrambi i mittenti Enel.
    Per noreply.enelenergia@enel.com scarica il corpo completo
    per estrarre il totale da pagare.
    """
    all_emails = []

    for sender in [SENDER_SANDSIV, SENDER_NOREPLY]:
        needs_body = sender == SENDER_NOREPLY
        print(f"\nRicerca email da: {sender}")

        msg_ids = _list_message_ids(service, sender)
        print(f"  Trovati {len(msg_ids)} messaggi — scaricamento dettagli...")

        for i, msg_info in enumerate(msg_ids, 1):
            try:
                fmt = "full" if needs_body else "metadata"
                kwargs = {"userId": "me", "id": msg_info["id"], "format": fmt}
                if not needs_body:
                    kwargs["metadataHeaders"] = ["From", "Subject", "Date"]

                msg = service.users().messages().get(**kwargs).execute()

                headers = {
                    h["name"].lower(): h["value"]
                    for h in msg.get("payload", {}).get("headers", [])
                }

                totale = ""
                if needs_body:
                    body_text = extract_body_text(msg.get("payload", {}))
                    totale = extract_totale_da_pagare(body_text)

                parsed_date = parse_date(headers.get("date", ""))

                all_emails.append({
                    "id": msg_info["id"],
                    "sender_key": sender,
                    "from": headers.get("from", sender),
                    "subject": headers.get("subject", "(nessun oggetto)"),
                    "date_raw": headers.get("date", ""),
                    "date": parsed_date,
                    "snippet": msg.get("snippet", ""),
                    "totale_da_pagare": totale,
                })

                if i % 50 == 0 or i == len(msg_ids):
                    print(f"  {i}/{len(msg_ids)} messaggi elaborati")

            except Exception as e:
                print(f"  Errore nel messaggio {msg_info['id']}: {e}")

    # Ordina per data decrescente
    all_emails.sort(key=lambda x: x["date"] or datetime.min, reverse=True)
    print(f"\nRecupero completato: {len(all_emails)} email totali.\n")
    return all_emails


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------

def create_excel_report(emails, output_file):
    """Crea il file Excel con i dettagli delle email."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Email Enel"

    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", fgColor="C0392B")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    data_font = Font(name="Arial", size=10)
    amount_font = Font(name="Arial", size=10, bold=True, color="1A5276")
    center_top = Alignment(horizontal="center", vertical="top")
    left_wrap = Alignment(horizontal="left", vertical="top", wrap_text=True)
    right_top = Alignment(horizontal="right", vertical="top")
    thin_border = Border(
        left=Side(style="thin", color="D0D0D0"),
        right=Side(style="thin", color="D0D0D0"),
        top=Side(style="thin", color="D0D0D0"),
        bottom=Side(style="thin", color="D0D0D0"),
    )
    alt_fill = PatternFill("solid", fgColor="FDECEA")

    # Colonne: #, Data, Mittente, Oggetto, Anteprima, Totale da pagare
    col_defs = [
        ("#",               5,  center_top),
        ("Data",           22,  center_top),
        ("Mittente",       30,  left_wrap),
        ("Oggetto",        50,  left_wrap),
        ("Anteprima",      70,  left_wrap),
        ("Totale da pagare", 18, right_top),
    ]

    for col, (title, width, _) in enumerate(col_defs, 1):
        cell = ws.cell(row=1, column=col, value=title)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
        ws.column_dimensions[get_column_letter(col)].width = width

    ws.row_dimensions[1].height = 20

    for idx, email in enumerate(emails, 1):
        row = idx + 1
        fill = alt_fill if idx % 2 == 0 else None

        date_value = email["date"]
        date_display = (
            date_value.strftime("%d/%m/%Y %H:%M") if date_value else email["date_raw"]
        )

        cells_data = [
            (idx,                       col_defs[0][2], data_font),
            (date_display,              col_defs[1][2], data_font),
            (email["sender_key"],       col_defs[2][2], data_font),
            (email["subject"],          col_defs[3][2], data_font),
            (email["snippet"],          col_defs[4][2], data_font),
            (email["totale_da_pagare"], col_defs[5][2], amount_font),
        ]

        for col, (value, align, font) in enumerate(cells_data, 1):
            c = ws.cell(row=row, column=col, value=value)
            c.font = font
            c.alignment = align
            c.border = thin_border
            if fill:
                c.fill = fill

        ws.row_dimensions[row].height = 40

    ws.freeze_panes = "A2"
    last_row = len(emails) + 1
    ws.auto_filter.ref = f"A1:F{last_row}"

    wb.save(output_file)
    print(f"File Excel salvato: {output_file}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"enel_emails_{timestamp}.xlsx"

    print("=" * 60)
    print("   Gmail Utilities Bill Retrieve")
    print(f"   Mittenti: {SENDER_SANDSIV}")
    print(f"             {SENDER_NOREPLY}")
    print("=" * 60)

    creds = authenticate()
    service = build("gmail", "v1", credentials=creds)

    profile = service.users().getProfile(userId="me").execute()
    print(f"\nAccount: {profile.get('emailAddress', 'N/A')}")

    emails = fetch_all_emails(service)

    if not emails:
        print("Nessuna email trovata.")
        return

    create_excel_report(emails, output_file)

    print("\nPrime 10 email (più recenti):")
    print("-" * 60)
    for i, email in enumerate(emails[:10], 1):
        date_str = email["date"].strftime("%d/%m/%Y") if email["date"] else "N/D"
        totale = f"  →  {email['totale_da_pagare']}" if email["totale_da_pagare"] else ""
        print(f"  {i:2}. [{date_str}] {email['subject'][:40]}{totale}")

    print(f"\nFile generato: {output_file}")


if __name__ == "__main__":
    main()
