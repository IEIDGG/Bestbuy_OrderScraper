import imaplib
import email
from bs4 import BeautifulSoup
import re
import ssl


def connect_to_mail(EMAIL, PASSWORD, imap_server, imap_port, use_ssl=True):
    if use_ssl:
        mail = imaplib.IMAP4_SSL(imap_server, imap_port)
    else:
        context = ssl.create_default_context()
        context.check_hostname = False
        context.verify_mode = ssl.CERT_NONE
        mail = imaplib.IMAP4(imap_server, imap_port)
        mail.starttls(ssl_context=context)

    mail.login(EMAIL, PASSWORD)
    return mail


def get_email_content(mail, folder):
    mail.select(folder)
    _, message_numbers = mail.search(None,
                                     'FROM "BestBuyInfo@emailinfo.bestbuy.com" SUBJECT "Enjoy 1 month free of Game Pass Ultimate with your Best Buy purchase."')
    return message_numbers[0].split()


def extract_xbox_code(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    code_element = soup.find('strong', string=re.compile(r'Code:'))
    if code_element:
        code_match = re.search(r'Code:\s*([A-Z0-9-]+)', code_element.text)
        if code_match:
            return code_match.group(1)
    return None


def process_email(mail, num):
    _, msg_data = mail.fetch(num, "(RFC822)")
    email_body = msg_data[0][1]
    email_message = email.message_from_bytes(email_body)

    for part in email_message.walk():
        if part.get_content_type() == "text/html":
            html_body = part.get_payload(decode=True).decode()
            xbox_code = extract_xbox_code(html_body)
            if xbox_code:
                print(f"{xbox_code}")
                return 1
    return 0


def process_emails(mail, all_mail_folder):
    emails = get_email_content(mail, all_mail_folder)
    print(f"Found {len(emails)} matching emails")

    total_codes = 0
    for num in emails:
        total_codes += process_email(mail, num)

    mail.logout()
    print("Closed mail connection")
    print(f"\nTotal Xbox codes found: {total_codes}")


def main_proton(EMAIL, PASSWORD):
    mail = connect_to_mail(EMAIL, PASSWORD, "127.0.0.1", 1143, use_ssl=False)
    print("Connected to Proton Mail server...")
    process_emails(mail, '"All Mail"')


def main_google(EMAIL, PASSWORD):
    mail = connect_to_mail(EMAIL, PASSWORD, "imap.gmail.com", 993)
    print("Connected to Gmail server...")
    process_emails(mail, "INBOX")


def read_credentials(file_path):
    with open(file_path, 'r') as file:
        lines = file.readlines()
        user = lines[0].split('=')[1].strip()
        password = lines[1].split('=')[1].strip()
        proton = lines[2].split('=')[1].strip().lower() == 'true'
    return user, password, proton


if __name__ == "__main__":
    user, password, is_proton = read_credentials('credentials.txt')
    main_function = main_proton if is_proton else main_google
    main_function(user, password)
    input("Press ENTER to exit")
