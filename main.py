import imaplib
import email
from openpyxl import Workbook
from datetime import datetime
import chardet
import email.header


# Konfiguracja dostępu do skrzynki mailowej i tytułu maili zgłoszeniowych
file = 'dane_wrazliwe.txt'
f=  open(file, 'r')
f_content = f.readlines()
mail_server = f_content[0]
mail_username = f_content[1]
mail_password = f_content[2]
header = f_content[3]
f.close()

# Połączenie z serwerem IMAP
mail = imaplib.IMAP4_SSL(mail_server)
mail.login(mail_username, mail_password)
mail.select('inbox')

# Wyszukiwanie wiadomości
result, data = mail.search(None, 'ALL')
mail_ids = data[0].split()
num_mails = len(mail_ids)
print(f"Liczba maili w skrzynce to {num_mails}.")
# print(mail_ids)

# Tworzenie Excelka (później trzeba to zautoamtyzować i z chmurą połączyć)
wb = Workbook()
ws = wb.active
pola_arkusza = [
    "Imię", "Nazwisko", "PESEL", "Data urodzenia", "Adres", "Kod pocztowy",
    "Miasto", "Imię matki/opiekuna", "Nazwisko matki/opiekuna", 
    "Imię ojca/opiekuna", "Nazwisko ojca/opiekuna", "Numer kontaktowy",
    "Email kontaktowy", "Dieta wegetariańska", "Rozmiar koszulki",
    "Informacje dodatkowe", "Był/a już na obozie?", "Zapoznano z regulaminem"
] #potencjalnie za dużo
ws.append(pola_arkusza)

# Parsowania treści maila
def parse_email(msg):
    parsed_email = {}
    for key in ['From', 'Date']:
        value = msg[key]
        parsed_email[key] = value

    subject = msg.get('Subject')
    if subject:
        decoded_subject = email.header.decode_header(subject)
        subject_parts = []
        for part, encoding in decoded_subject: #dekodowanie
            if isinstance(part, bytes):
                try:
                    if encoding:
                        subject_parts.append(part.decode(encoding))
                    else:
                        subject_parts.append(part.decode('utf-8', errors='replace'))
                except LookupError:
                    subject_parts.append(part.decode('utf-8', errors='replace'))
            else:
                subject_parts.append(part)
        subject = ''.join(subject_parts)
        parsed_email['Subject'] = subject

    body = ""
    for part in msg.walk():
        if part.get_content_type() == "text/plain" or part.get_content_type() == "text/html":
            payload = part.get_payload(decode=True)
            if payload:
                encoding = chardet.detect(payload)['encoding']
                if encoding:
                    body += payload.decode(encoding, errors='replace')
    parsed_email['Body'] = body.strip()

    return parsed_email



for mail_id in mail_ids[(num_mails-10):num_mails]:
    result, data = mail.fetch(mail_id, "(RFC822)")
    raw_email = data[0][1]
    msg = email.message_from_bytes(raw_email)
    parsed_email = parse_email(msg)
    subject = parsed_email['Subject']
    author = parsed_email['From']
    date = parsed_email['Date']
    content = parsed_email['Body']
    print("Subject:", subject)
    # print("From:", author)
    # print("Date:", date)
    print("\n \n")
    if header in subject:
        print("Body:", content)
        for line in content.split('\n'):
            print(line)
        # row = [
        #     content.get("Imię", ""),
        #     content.get("Nazwisko", ""),
        #     content.get("PESEL", ""),
        #     content.get("Data urodzenia", ""),
        #     content.get("Adres zamieszkania (ulica, nr budynku)", ""),
        #     content.get("Kod pocztowy", ""),
        #     content.get("Miasto", ""),
        #     content.get("Imię i nazwisko matki/opiekuna", "").split()[0],
        #     content.get("Imię i nazwisko matki/opiekuna", "").split()[1],
        #     content.get("Imię i nazwisko ojca/opiekuna", "").split()[0],
        #     content.get("Imię i nazwisko ojca/opiekuna", "").split()[1],
        #     content.get("Numer kontaktowy do rodzica/opiekuna", ""),
        #     content.get("Email kontaktowy do rodzica/opiekuna", ""),
        #     content.get("Dieta wegetariańska", ""),
        #     content.get("Prosimy o wybranie rozmiaru koszulki:", ""),
        #     content.get("Informacje dodatkowe", ""),
        #     content.get("Czy dziecko było już na innym obozie min. 7-dniowym?", ""),
        #     content.get("Czy zapoznałeś się z informatorem obozowym i akceptujesz regulamin obozu?", "")
        # ]
        # ws.append(row)

# Zapisanie arkusza
timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
wb.save(f"form_data_{timestamp}.xlsx")

# Zamknięcie połączenia z serwerem
mail.close()
mail.logout()
