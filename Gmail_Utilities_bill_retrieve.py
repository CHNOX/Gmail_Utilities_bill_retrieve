"""
Gmail Utilities Bill Retrieve
==============================
Recupera tutte le email ricevute dal mittente enelenergia@enel.sandsiv.com
e genera un file Excel con i dettagli (data, oggetto, anteprima).

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

import os
import sys
import re
from datetime import datetime

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

CREDENTIALS_FILE = "credentials.json"
TOKEN_FILE = "token.json"
TARGET_SENDER = "enelenergia@enel.sandsiv.com"
BATCH_SIZE = 500


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


def parse_date(date_str):
    """Converte la stringa data dell'header in oggetto datetime."""
    if not date_str:
        return None
    # Rimuove eventuale timezone testuale es. "(UTC)", "(CET)"
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


def fetch_emails_from_sender(service, sender):
    """Recupera tutte le email del mittente specificato."""
    emails = []
    page_token = None
    batch_num = 0

    print(f"\nRicerca email da: {sender}\n")

    while True:
        batch_num += 1
        params = {
            "userId": "me",
            "maxResults": BATCH_SIZE,
            "q": f"from:{sender}",
        }
        if page_token:
            params["pageToken"] = page_token

        results = service.users().messages().list(**params).execute()
        messages = results.get("messages", [])

        if not messages:
            break

        for msg_info in messages:
            try:
                msg = service.users().messages().get(
                    userId="me",
                    id=msg_info["id"],
                    format="metadata",
                    metadataHeaders=["From", "Subject", "Date"],
                ).execute()

                headers = {
                    h["name"].lower(): h["value"]
                    for h in msg.get("payload", {}).get("headers", [])
                }

                snippet = msg.get("snippet", "")
                date_raw = headers.get("date", "")
                subject = headers.get("subject", "(nessun oggetto)")
                from_hdr = headers.get("from", sender)

                parsed_date = parse_date(date_raw)

                emails.append({
                    "id": msg_info["id"],
                    "from": from_hdr,
                    "subject": subject,
                    "date_raw": date_raw,
                    "date": parsed_date,
                    "snippet": snippet,
                })

            except Exception as e:
                print(f"  Errore nel messaggio {msg_info['id']}: {e}")

        print(f"  Batch {batch_num}: +{len(messages)} messaggi (totale recuperati: {len(emails)})")

        page_token = results.get("nextPageToken")
        if not page_token:
            break

    # Ordina per data decrescente (più recenti prima)
    emails.sort(key=lambda x: x["date"] or datetime.min, reverse=True)

    print(f"\nRecupero completato: {len(emails)} email trovate.\n")
    return emails


def create_excel_report(emails, output_file):
    """Crea il file Excel con i dettagli delle email."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Email Enel"

    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", fgColor="C0392B")  # Rosso Enel
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    data_font = Font(name="Arial", size=10)
    date_align = Alignment(horizontal="center", vertical="top")
    wrap_align = Alignment(horizontal="left", vertical="top", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin", color="D0D0D0"),
        right=Side(style="thin", color="D0D0D0"),
        top=Side(style="thin", color="D0D0D0"),
        bottom=Side(style="thin", color="D0D0D0"),
    )
    alt_fill = PatternFill("solid", fgColor="FDECEA")

    col_headers = ["#", "Data", "Oggetto", "Anteprima"]
    col_widths = [5, 22, 55, 80]

    for col, (title, width) in enumerate(zip(col_headers, col_widths), 1):
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
            (idx, date_align),
            (date_display, date_align),
            (email["subject"], wrap_align),
            (email["snippet"], wrap_align),
        ]

        for col, (value, align) in enumerate(cells_data, 1):
            c = ws.cell(row=row, column=col, value=value)
            c.font = data_font
            c.alignment = align
            c.border = thin_border
            if fill:
                c.fill = fill

        ws.row_dimensions[row].height = 40

    # Blocca prima riga
    ws.freeze_panes = "A2"

    # Filtro automatico
    last_row = len(emails) + 1
    ws.auto_filter.ref = f"A1:D{last_row}"

    wb.save(output_file)
    print(f"File Excel salvato: {output_file}")


def main():
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"enel_emails_{timestamp}.xlsx"

    print("=" * 55)
    print("   Gmail Utilities Bill Retrieve")
    print(f"   Mittente: {TARGET_SENDER}")
    print("=" * 55)

    creds = authenticate()
    service = build("gmail", "v1", credentials=creds)

    profile = service.users().getProfile(userId="me").execute()
    email_address = profile.get("emailAddress", "N/A")
    print(f"\nAccount: {email_address}")

    emails = fetch_emails_from_sender(service, TARGET_SENDER)

    if not emails:
        print("Nessuna email trovata per questo mittente.")
        return

    create_excel_report(emails, output_file)

    print("\nPrime 10 email (più recenti):")
    print("-" * 55)
    for i, email in enumerate(emails[:10], 1):
        date_str = (
            email["date"].strftime("%d/%m/%Y") if email["date"] else "N/D"
        )
        subject = email["subject"][:45]
        print(f"  {i:2}. [{date_str}] {subject}")

    print(f"\nFile generato: {output_file}")


if __name__ == "__main__":
    main()
