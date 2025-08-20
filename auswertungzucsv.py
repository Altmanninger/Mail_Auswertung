import imaplib
import email
from email.header import decode_header
import re
import csv
import sys
import os
from dotenv import load_dotenv

# Umgebungsvariablen aus .env laden
load_dotenv()

IMAP_SERVER = os.getenv('IMAP_SERVER')
EMAIL_ACCOUNT = os.getenv('EMAIL_ACCOUNT')
PASSWORD = os.getenv('PASSWORD')
FOLDER = os.getenv('FOLDER', 'INBOX')  # Default: INBOX

# Extraktionsmuster erweitert für alle Mail-Typ-Varianten
patterns = {
    'Mail-Typ': '',  # Wird später gesetzt
    'Anmeldedatum': r'(?:ANMELDUNG\s*AKTUALISIERUNG|ANMELDUNG|AKTUALISIERUNG|ABMELDUNG) am:\s*([^\n]+)',
    'Name': r'Name:\s*([^\n]+)',
    'Geburtsdatum': r'Geburtsdatum:\s*([^\n]+)',
    'Verantwortliche': r'Verantwortliche bei Jugendlichen:\s*([^\n]+)',
    'Adresse': r'Adresse:\s*([^\n]+)',
    'Telefon-Mobil': r'Telefon-Mobil:\s*([^\n,]*)',
    'Telefon-Festnetz': r'Telefon-Festnetz:\s*([^\n]*)',
    'E-Mail': r'E-Mail:\s*([^\n]+)',
    'Bemerkungen': r'Bemerkung\(en\):\s*(.*?)(?=Datenschutz-Grundverordnung:|$)',
    'DSGVO': r'Datenschutz-Grundverordnung:\s*([^\n]*)',
    'Turnstunde 1': r'1\.Turnstunde:\s*([^\n]*)',
    'Turnstunde 2': r'2\.Turnstunde:\s*([^\n]*)',
    'Turnstunde 3': r'3\.Turnstunde:\s*([^\n]*)',
}

def remove_html_tags(text):
    # Entfernt alle HTML-Tags mit einem Regex
    return re.sub(r'<[^>]+>', '', text)

def extract_fields(body, mail_type):
    data = {}
    data['Mail-Typ'] = mail_type  # Mail-Typ hinzufügen
    for key, pattern in patterns.items():
        if key == 'Mail-Typ':
            continue  # Bereits gesetzt
        if key == 'Bemerkungen':
            # Für Bemerkungen DOTALL verwenden, damit . auch \n matcht
            match = re.search(pattern, body, re.DOTALL)
        else:
            match = re.search(pattern, body)
        value = match.group(1).strip() if match else ""
        value = remove_html_tags(value)  # HTML-Tags entfernen
        # Mehrfache Leerzeichen und Zeilenwechsel durch einzelne Leerzeichen ersetzen
        if key == 'Bemerkungen':
            value = re.sub(r'\s+', ' ', value).strip()
        data[key] = value
    return data

def get_body(msg):
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            cdispo = str(part.get('Content-Disposition'))
            # Nur Textteile, keine Anhänge
            if ctype == 'text/plain' and 'attachment' not in cdispo:
                charset = part.get_content_charset() or 'utf-8'
                try:
                    return part.get_payload(decode=True).decode(charset, errors='replace')
                except Exception:
                    continue
    else:
        charset = msg.get_content_charset() or 'utf-8'
        try:
            return msg.get_payload(decode=True).decode(charset, errors='replace')
        except Exception:
            return ""
    return ""

def decode_subject(subject):
    decoded_parts = decode_header(subject)
    subject_parts = []
    for part, encoding in decoded_parts:
        if isinstance(part, bytes):
            try:
                subject_parts.append(part.decode(encoding if encoding else 'utf-8'))
            except Exception:
                subject_parts.append(part.decode('utf-8', errors='replace'))
        else:
            subject_parts.append(part)
    return ''.join(subject_parts)

def determine_mail_type(subject):
    """Bestimmt den Mail-Typ basierend auf dem Betreff"""
    subject_upper = subject.upper()
    if 'ANMELDUNG' in subject_upper and 'AKTUALISIERUNG' in subject_upper:
        return 'ANMELDUNG AKTUALISIERUNG'
    elif 'ANMELDUNG' in subject_upper:
        return 'ANMELDUNG'
    elif 'AKTUALISIERUNG' in subject_upper:
        return 'AKTUALISIERUNG'
    elif 'ABMELDUNG' in subject_upper:
        return 'ABMELDUNG'
    else:
        return 'UNBEKANNT'

def main():
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL_ACCOUNT, PASSWORD)
        mail.select(FOLDER)
    except Exception as e:
        print(f"Fehler beim Verbinden mit dem Mailserver: {e}")
        sys.exit(1)

    # E-Mails mit ANMELDUNG, AKTUALISIERUNG oder ABMELDUNG im Betreff
    status, messages = mail.search(None, '(OR (OR (SUBJECT "ANMELDUNG") (SUBJECT "AKTUALISIERUNG")) (SUBJECT "ABMELDUNG"))')
    if status != "OK":
        print("Fehler beim Suchen der Mails.")
        return

    email_ids = messages[0].split()
    print(f"{len(email_ids)} Mails gefunden.")

    fields = list(patterns.keys())
    rows = []

    for num in email_ids:
        status, msg_data = mail.fetch(num, "(RFC822)")
        if status != "OK":
            print(f"Fehler beim Abrufen der Mail {num}")
            continue
        msg = email.message_from_bytes(msg_data[0][1])

        # Betreff dekodieren und Mail-Typ bestimmen
        subject = decode_subject(msg.get("Subject", ""))
        mail_type = determine_mail_type(subject)

        body = get_body(msg)
        if body:
            data = extract_fields(body, mail_type)
            rows.append(data)
        else:
            print(f"Keine lesbare Mail bei ID {num}")

    if not rows:
        print("Keine passenden Mails gefunden oder keine Daten extrahiert.")
        return

    # Schreibe alles in die CSV
    with open('anmeldungen.csv', 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fields)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)

    print("Fertig! Alle Daten in 'anmeldungen.csv' gespeichert.")
    mail.logout()

if __name__ == "__main__":
    main()