"""
Gmail Utilities Bill Retrieve
==============================
Legge la configurazione da settings.json e per ogni mittente:
  - recupera tutte le email
  - estrae i valori delle chiavi configurate (es. "TOTALE DA PAGARE")

Genera un file Excel con una colonna per ogni chiave di estrazione.

PREREQUISITI:
1. Python 3.7+
2. pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib openpyxl

3. Scarica credentials.json da Google Cloud Console (vedi README).

4. Configura settings.json con i mittenti e le chiavi da estrarre.
"""

import base64
import json
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
TOKEN_FILE        = "token.json"
SETTINGS_FILE     = "settings.json"
BATCH_SIZE        = 500


# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------

def load_settings():
    """
    Carica settings.json. Struttura attesa:
    {
      "senders": [
        { "email": "...", "extract_keys": ["CHIAVE 1", "CHIAVE 2"] },
        ...
      ]
    }
    """
    if not os.path.exists(SETTINGS_FILE):
        print(f"ERRORE: File '{SETTINGS_FILE}' non trovato.")
        print("Crea il file con la lista dei mittenti e delle chiavi da estrarre.")
        sys.exit(1)

    with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
        settings = json.load(f)

    senders = settings.get("senders", [])
    if not senders:
        print(f"ERRORE: Nessun mittente configurato in '{SETTINGS_FILE}'.")
        sys.exit(1)

    return senders


def collect_all_keys(senders_config):
    """Restituisce la lista ordinata di tutte le chiavi uniche tra i mittenti."""
    seen = {}
    for s in senders_config:
        for key in s.get("extract_keys", []):
            seen[key] = None  # dict per preservare l'ordine di inserimento
    return list(seen)


# ---------------------------------------------------------------------------
# Autenticazione
# ---------------------------------------------------------------------------

def authenticate():
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
                sys.exit(1)
            print("Apertura browser per autorizzazione...")
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())
        print("Autenticazione completata e salvata.")

    return creds


# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------

def parse_date(date_str):
    if not date_str:
        return None
    date_str = re.sub(r'\s*\([^)]*\)\s*$', '', date_str.strip())
    for fmt in (
        "%a, %d %b %Y %H:%M:%S %z",
        "%a, %d %b %Y %H:%M:%S",
        "%d %b %Y %H:%M:%S %z",
        "%d %b %Y %H:%M:%S",
    ):
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    return None


class _HTMLStripper(HTMLParser):
    def __init__(self):
        super().__init__()
        self._parts = []

    def handle_data(self, data):
        self._parts.append(data)

    def get_text(self):
        return " ".join(self._parts)


def _html_to_text(html_content):
    s = _HTMLStripper()
    s.feed(html_content)
    return s.get_text()


def extract_body_text(payload):
    """Estrae ricorsivamente il testo dal payload Gmail (plain > html > multipart)."""
    mime = payload.get("mimeType", "")

    if mime == "text/plain":
        raw = payload.get("body", {}).get("data", "")
        if raw:
            return base64.urlsafe_b64decode(raw + "==").decode("utf-8", errors="replace")

    if mime == "text/html":
        raw = payload.get("body", {}).get("data", "")
        if raw:
            return _html_to_text(
                base64.urlsafe_b64decode(raw + "==").decode("utf-8", errors="replace")
            )

    plain, html = "", ""
    for part in payload.get("parts", []):
        pt = part.get("mimeType", "")
        raw = part.get("body", {}).get("data", "")
        if pt == "text/plain" and raw:
            plain = base64.urlsafe_b64decode(raw + "==").decode("utf-8", errors="replace")
        elif pt == "text/html" and raw:
            html = _html_to_text(
                base64.urlsafe_b64decode(raw + "==").decode("utf-8", errors="replace")
            )
        elif part.get("parts"):
            nested = extract_body_text(part)
            plain = plain or nested

    return plain or html


def extract_value_for_key(text, key):
    """
    Cerca `key` nel testo e restituisce il valore associato.
    Strategia:
      1. Importo con simbolo €  (es. "123,45 €" o "€ 123,45")
      2. Numero decimale generico (es. "123,45" o "123.45")
      3. Prima parola/token non vuoto dopo la chiave
    """
    if not text:
        return ""

    normalized = re.sub(r"\s+", " ", text)
    escaped_key = re.escape(key)

    # 1. importo con €
    for pattern in (
        rf"{escaped_key}[^0-9€]*?([\d]{{1,6}}[.,][\d]{{2}})\s*€",
        rf"{escaped_key}[^€]*?€\s*([\d]{{1,6}}[.,][\d]{{2}})",
    ):
        m = re.search(pattern, normalized, re.IGNORECASE)
        if m:
            return f"€ {m.group(1).strip()}"

    # 2. numero decimale senza €
    m = re.search(
        rf"{escaped_key}[^0-9]*?([\d]{{1,6}}[.,][\d]{{2}})",
        normalized, re.IGNORECASE
    )
    if m:
        return m.group(1).strip()

    # 3. primo token non vuoto dopo la chiave
    m = re.search(
        rf"{escaped_key}\s*[:\-]?\s*(\S{{1,60}})",
        normalized, re.IGNORECASE
    )
    if m:
        return m.group(1).strip()

    return ""


# ---------------------------------------------------------------------------
# Recupero email
# ---------------------------------------------------------------------------

def _list_message_ids(service, sender):
    ids, page_token = [], None
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


def fetch_all_emails(service, senders_config):
    """Recupera email per ogni mittente configurato ed estrae le chiavi richieste."""
    all_emails = []

    for sender_cfg in senders_config:
        sender     = sender_cfg["email"]
        keys       = sender_cfg.get("extract_keys", [])
        needs_body = bool(keys)

        print(f"\nRicerca email da: {sender}")
        if keys:
            print(f"  Chiavi da estrarre: {', '.join(keys)}")

        msg_ids = _list_message_ids(service, sender)
        print(f"  Trovati {len(msg_ids)} messaggi — scaricamento dettagli...")

        for i, msg_info in enumerate(msg_ids, 1):
            try:
                fmt    = "full" if needs_body else "metadata"
                kwargs = {"userId": "me", "id": msg_info["id"], "format": fmt}
                if not needs_body:
                    kwargs["metadataHeaders"] = ["From", "Subject", "Date"]

                msg = service.users().messages().get(**kwargs).execute()

                headers = {
                    h["name"].lower(): h["value"]
                    for h in msg.get("payload", {}).get("headers", [])
                }

                extracted = {}
                if needs_body:
                    body_text = extract_body_text(msg.get("payload", {}))
                    for key in keys:
                        extracted[key] = extract_value_for_key(body_text, key)

                all_emails.append({
                    "sender_email": sender,
                    "from":         headers.get("from", sender),
                    "subject":      headers.get("subject", "(nessun oggetto)"),
                    "date_raw":     headers.get("date", ""),
                    "date":         parse_date(headers.get("date", "")),
                    "snippet":      msg.get("snippet", ""),
                    "extracted":    extracted,   # { "CHIAVE": "valore", ... }
                })

                if i % 50 == 0 or i == len(msg_ids):
                    print(f"  {i}/{len(msg_ids)} elaborati")

            except Exception as e:
                print(f"  Errore nel messaggio {msg_info['id']}: {e}")

    all_emails.sort(key=lambda x: x["date"] or datetime.min, reverse=True)
    print(f"\nRecupero completato: {len(all_emails)} email totali.\n")
    return all_emails


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------

def create_excel_report(emails, all_keys, output_file):
    """
    Crea il file Excel.
    Colonne fisse: #, Data, Mittente, Oggetto, Anteprima
    Colonne dinamiche: una per ogni chiave in all_keys
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Email Enel"

    header_font  = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    header_fill  = PatternFill("solid", fgColor="C0392B")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    data_font    = Font(name="Arial", size=10)
    value_font   = Font(name="Arial", size=10, bold=True, color="1A5276")
    center_top   = Alignment(horizontal="center", vertical="top")
    left_wrap    = Alignment(horizontal="left",   vertical="top", wrap_text=True)
    right_top    = Alignment(horizontal="right",  vertical="top")
    thin_border  = Border(
        left=Side(style="thin", color="D0D0D0"), right=Side(style="thin", color="D0D0D0"),
        top=Side(style="thin", color="D0D0D0"),  bottom=Side(style="thin", color="D0D0D0"),
    )
    alt_fill = PatternFill("solid", fgColor="FDECEA")

    # Colonne fisse
    fixed_cols = [
        ("#",         5,  center_top, data_font),
        ("Data",     22,  center_top, data_font),
        ("Mittente", 30,  left_wrap,  data_font),
        ("Oggetto",  50,  left_wrap,  data_font),
        ("Anteprima",70,  left_wrap,  data_font),
    ]
    # Colonne dinamiche (una per chiave)
    dynamic_cols = [(key, 20, right_top, value_font) for key in all_keys]
    all_cols = fixed_cols + dynamic_cols

    for col, (title, width, *_) in enumerate(all_cols, 1):
        cell = ws.cell(row=1, column=col, value=title)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
        ws.column_dimensions[get_column_letter(col)].width = width

    ws.row_dimensions[1].height = 20

    for idx, email in enumerate(emails, 1):
        row  = idx + 1
        fill = alt_fill if idx % 2 == 0 else None

        date_val = email["date"]
        date_str = date_val.strftime("%d/%m/%Y %H:%M") if date_val else email["date_raw"]

        row_values = [
            (idx,                  center_top, data_font),
            (date_str,             center_top, data_font),
            (email["sender_email"],left_wrap,  data_font),
            (email["subject"],     left_wrap,  data_font),
            (email["snippet"],     left_wrap,  data_font),
        ]
        for key in all_keys:
            row_values.append(
                (email["extracted"].get(key, ""), right_top, value_font)
            )

        for col, (value, align, font) in enumerate(row_values, 1):
            c = ws.cell(row=row, column=col, value=value)
            c.font = font
            c.alignment = align
            c.border = thin_border
            if fill:
                c.fill = fill

        ws.row_dimensions[row].height = 40

    ws.freeze_panes = "A2"
    last_col = get_column_letter(len(all_cols))
    ws.auto_filter.ref = f"A1:{last_col}{len(emails) + 1}"

    wb.save(output_file)
    print(f"File Excel salvato: {output_file}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    senders_config = load_settings()
    all_keys       = collect_all_keys(senders_config)
    timestamp      = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file    = f"enel_emails_{timestamp}.xlsx"

    print("=" * 60)
    print("   Gmail Utilities Bill Retrieve")
    for s in senders_config:
        keys_str = ", ".join(s.get("extract_keys", [])) or "—"
        print(f"   • {s['email']}  [{keys_str}]")
    print("=" * 60)

    creds   = authenticate()
    service = build("gmail", "v1", credentials=creds)

    profile = service.users().getProfile(userId="me").execute()
    print(f"\nAccount: {profile.get('emailAddress', 'N/A')}")

    emails = fetch_all_emails(service, senders_config)

    if not emails:
        print("Nessuna email trovata.")
        return

    create_excel_report(emails, all_keys, output_file)

    print("\nPrime 10 email (più recenti):")
    print("-" * 60)
    for i, email in enumerate(emails[:10], 1):
        date_str  = email["date"].strftime("%d/%m/%Y") if email["date"] else "N/D"
        extras    = "  →  " + "  |  ".join(
            f"{k}: {v}" for k, v in email["extracted"].items() if v
        ) if email["extracted"] else ""
        print(f"  {i:2}. [{date_str}] {email['subject'][:40]}{extras}")

    print(f"\nFile generato: {output_file}")


if __name__ == "__main__":
    main()
