import imaplib
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

def get_email_content(mail, folder, search_criteria):
    try:
        mail.select(folder)
        _, message_numbers = mail.search(None, search_criteria)
        return message_numbers[0].split()
    except imaplib.IMAP4.error as e:
        print(f"Error accessing folder '{folder}': {str(e)}")
        return []