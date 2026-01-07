import requests
import fitz  # PyMuPDF
import re
from datetime import datetime
import os
import json
import pickle
import base64
from base64 import b64decode
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from email.mime.text import MIMEText

# --- Konfiguration ---
ISIN = "IE0007UPSEA3"
PDF_URL = "https://lt.morningstar.com/1c6qh1t6k9/snapshotpdf/default.aspx?Id=0P0001RHV3&LanguageId=en-GB"
PDF_FILE = "ms_snapshot.pdf"
USD_TO_EUR = 0.929
ALERT_THRESHOLD = 4.2

# Trade Republic Werte (Beispiel)
TR_PRICE_EUR = 4.33
TR_YIELD = 4.23

# Gmail
SCOPES = ['https://www.googleapis.com/auth/gmail.send']
EMAIL_TO = "tamer.demirel@googlemail.com"

# Alert-Datei, um mehrfaches Mailen zu verhindern
ALERT_FILE = ".alert_sent"

# --- Gmail Credentials aus Env lesen ---
creds_json_b64 = os.getenv("GMAIL_CREDENTIALS")
token_pickle_b64 = os.getenv("GMAIL_TOKEN")

# Dateien lokal für PyGmail erzeugen (falls nötig)
CREDENTIALS_JSON = "credentials.json"
TOKEN_PICKLE = "token.pickle"

if creds_json_b64:
    creds_json = json.loads(b64decode(creds_json_b64))
    with open(CREDENTIALS_JSON, "w") as f:
        json.dump(creds_json, f)

if token_pickle_b64:
    token_bytes = b64decode(token_pickle_b64)
    with open(TOKEN_PICKLE, "wb") as f:
        f.write(token_bytes)

# --- PDF herunterladen ---
def download_pdf(url, filename):
    try:
        r = requests.get(url)
        if r.status_code == 200:
            with open(filename, "wb") as f:
                f.write(r.content)
            return True
    except Exception as e:
        print("Fehler beim Download:", e)
    return False

# --- PDF Text extrahieren ---
def extract_pdf_text(filename):
    try:
        doc = fitz.open(filename)
        text = ""
        for page in doc:
            text += page.get_text()
        return text
    except Exception as e:
        print("PDF konnte nicht gelesen werden:", e)
        return ""

# --- PDF parsen ---
def parse_morningstar_snapshot(text):
    results = {}

    # NAV / Preis
    m = re.search(r"NAV\s*\((?:\d{1,2}\s\w+\s\d{4})\)\s*([\d\.]+)\s*(USD|EUR)?", text)
    if m:
        price = float(m.group(1))
        currency = m.group(2)
        price_eur = round(price * USD_TO_EUR, 2) if currency == "USD" else price
        results["NAV/Preis USD"] = f"{price} {currency}"
        results["NAV/Preis EUR"] = f"{price_eur} EUR"
    else:
        results["NAV/Preis USD"] = results["NAV/Preis EUR"] = "nicht gefunden"

    # 12 Month Yield
    m = re.search(r"12 Month Yield\s*([\d\.]+)%", text)
    results["12M Yield"] = float(m.group(1)) if m else None

    # TER / Kosten
    m = re.search(r"Ongoing Cost\s*([\d\.]+)%", text)
    results["Kosten/TER"] = f"{m.group(1)}%" if m else "nicht gefunden"

    # TR Werte
    results["TR Preis EUR"] = f"{TR_PRICE_EUR} EUR"
    results["TR Yield"] = TR_YIELD
    return results

# --- Gmail senden via OAuth2 ---
def send_email(subject, body):
    creds = None
    if os.path.exists(TOKEN_PICKLE):
        creds = Credentials.from_authorized_user_file(TOKEN_PICKLE, SCOPES)
    else:
        flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_JSON, SCOPES)
        creds = flow.run_local_server(port=0)
        with open(TOKEN_PICKLE, 'w') as token:
            token.write(creds.to_json())
    service = build('gmail', 'v1', credentials=creds)
    message = MIMEText(body)
    message['to'] = EMAIL_TO
    message['from'] = EMAIL_TO
    message['subject'] = subject
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    service.users().messages().send(userId='me', body={'raw': raw}).execute()
    print("Email gesendet!")

# --- Hauptfunktion ---
def main():
    print("\n=== Morningstar PDF Scraper + Gmail OAuth2 –", datetime.now(), "===")

    # Prüfen, ob Alert schon gesendet wurde
    alert_sent = os.path.exists(ALERT_FILE)

    if download_pdf(PDF_URL, PDF_FILE):
        print(f"PDF heruntergeladen als '{PDF_FILE}'")
        text = extract_pdf_text(PDF_FILE)
        data = parse_morningstar_snapshot(text)

        print("\nErkannte Daten:")
        for k, v in data.items():
            print(f"{k}: {v}")

        # Alert prüfen
        if data["12M Yield"] and data["12M Yield"] > ALERT_THRESHOLD and not alert_sent:
            send_email(
                subject=f"iBond Alert: Yield über {ALERT_THRESHOLD}%",
                body=f"Yield ist aktuell {data['12M Yield']}% für {ISIN}\n\n"
                     f"TR Yield: {data['TR Yield']}%\n"
                     f"TR Preis: {data['TR Preis EUR']}\n"
                     f"NAV Preis: {data['NAV/Preis EUR']}"
            )
            # Alert merken
            with open(ALERT_FILE, "w") as f:
                f.write("sent")
    else:
        print("PDF Download fehlgeschlagen. Bitte URL prüfen.")

    print("\n========================================\n")

if __name__ == "__main__":
    main()
