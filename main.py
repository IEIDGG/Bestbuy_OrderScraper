from email_connector import connect_to_mail
from utils import read_credentials, print_summary
from config import GMAIL_IMAP_SERVER, GMAIL_IMAP_PORT, PROTON_IMAP_SERVER, PROTON_IMAP_PORT, ICLOUD_IMAP_SERVER, ICLOUD_IMAP_PORT
from email_handlers import process_confirmation_emails, process_cancellation_emails, process_shipped_emails
from file_handlers import save_and_display_orders

def main_proton(EMAIL, PASSWORD):
    mail = connect_to_mail(EMAIL, PASSWORD, PROTON_IMAP_SERVER, PROTON_IMAP_PORT, use_ssl=False)
    print("Connecting to mail server...")

    _, folders = mail.list()
    folder_dict = {'confirmation': None, 'cancelled': None, 'shipped': None}

    print("Searching for Best Buy folders...")
    for folder in folders:
        folder_name = folder.decode().split('"')[-2]
        if 'Bestbuy-Confirmation' in folder_name:
            folder_dict['confirmation'] = folder_name
        elif 'Bestbuy-Cancelled' in folder_name:
            folder_dict['cancelled'] = folder_name
        elif 'Bestbuy-Shipped' in folder_name:
            folder_dict['shipped'] = folder_name

    orders, confirmation_count = process_confirmation_emails(mail, folder_dict['confirmation']) if folder_dict[
        'confirmation'] else ([], 0)
    cancellation_count = process_cancellation_emails(mail, folder_dict['cancelled'], orders) if folder_dict[
        'cancelled'] else 0
    shipped_count, _ = process_shipped_emails(mail, folder_dict['shipped'], orders) if folder_dict['shipped'] else (
    0, 0)

    mail.logout()
    print("\nClosing mail connection")

    summary = save_and_display_orders(orders)
    print_summary(summary, confirmation_count, cancellation_count)


def main_google(EMAIL, PASSWORD):
    mail = connect_to_mail(EMAIL, PASSWORD, GMAIL_IMAP_SERVER, GMAIL_IMAP_PORT, use_ssl=True)
    print("Connecting to Gmail server...")

    orders, confirmation_count = process_confirmation_emails(mail, "INBOX")
    cancellation_count = process_cancellation_emails(mail, "INBOX", orders)
    shipped_count, _ = process_shipped_emails(mail, "INBOX", orders)

    mail.logout()
    print("\nClosing mail connection")

    summary = save_and_display_orders(orders)
    print_summary(summary, confirmation_count, cancellation_count)

def main_icloud(EMAIL, PASSWORD):
    mail = connect_to_mail(EMAIL, PASSWORD, ICLOUD_IMAP_SERVER, ICLOUD_IMAP_PORT, use_ssl=True)
    protocol = 'RFC5322'
    print("Connecting to iCloud server...")

    orders, confirmation_count = process_confirmation_emails(mail, "INBOX", protocol)
    cancellation_count = process_cancellation_emails(mail, "INBOX", orders, protocol)
    shipped_count, _ = process_shipped_emails(mail, "INBOX", orders, protocol)

    mail.logout()
    print("\nClosing mail connection")

    summary = save_and_display_orders(orders)
    print_summary(summary, confirmation_count, cancellation_count)

if __name__ == "__main__":
    user, password, email_type = read_credentials('credentials.txt')

    if email_type == 'proton':
        print("Using Proton method")
        main_function = main_proton
    elif email_type == 'icloud':
        print("Using iCloud method")
        main_function = main_icloud
    else:
        print("Using Gmail method")
        main_function = main_google

    main_function(user, password)
    input("ENTER to exit")