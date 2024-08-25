import imaplib
import email
from email.header import decode_header
import csv
from datetime import datetime
import ssl
from bs4 import BeautifulSoup


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


def process_email(mail, num, email_type):
    _, msg_data = mail.fetch(num, "(RFC822)")
    email_body = msg_data[0][1]
    email_message = email.message_from_bytes(email_body)

    date_tuple = email.utils.parsedate_tz(email_message['Date'])
    email_date = datetime.fromtimestamp(email.utils.mktime_tz(date_tuple)).strftime("%Y-%m-%d") if date_tuple else "Unknown"

    order_number = None
    tracking_numbers = []

    for part in email_message.walk():
        if part.get_content_type() == "text/html":
            html_body = part.get_payload(decode=True).decode()
            soup = BeautifulSoup(html_body, 'html.parser')

            if email_type == 'confirmation':
                order_span = soup.find('span', string=lambda text: text and 'BBY01-' in text)
                if order_span:
                    order_number = order_span.text.strip()
            elif email_type in ['cancelled', 'shipped']:
                order_span = soup.find('span', style='font: bold 23px Arial; color: #1d252c;')
                if order_span:
                    order_number = order_span.text.strip().replace("Order #", "")

            if email_type == 'shipped':
                tracking_spans = soup.find_all('span', style='font: bold 14px Arial')
                for span in tracking_spans:
                    if 'Tracking #:' in span.text:
                        tracking_link = span.find('a')
                        if tracking_link:
                            tracking_numbers.append(tracking_link.text.strip())

            break

    return email_date, order_number, tracking_numbers

def main_proton(EMAIL, PASSWORD):
    mail = connect_to_mail(EMAIL, PASSWORD, "127.0.0.1", 1143, use_ssl=False)

    print("Connecting to mail server...")
    _, folders = mail.list()
    folder_dict = {
        'confirmation': None,
        'cancelled': None,
        'shipped': None
    }

    print("Searching for Best Buy folders...")
    for folder in folders:
        folder_name = folder.decode().split('"')[-2]
        if 'Bestbuy-Confirmation' in folder_name:
            folder_dict['confirmation'] = folder_name
        elif 'Bestbuy-Cancelled' in folder_name:
            folder_dict['cancelled'] = folder_name
        elif 'Bestbuy-Shipped' in folder_name:
            folder_dict['shipped'] = folder_name

    for folder_type, folder_name in folder_dict.items():
        print(f"Found {folder_type} folder: {folder_name}")

    orders = []
    confirmation_count = 0
    cancellation_count = 0
    shipped_count = 0
    tracking_numbers_count = 0

    # Process confirmation emails
    if folder_dict['confirmation']:
        print(f"\nProcessing confirmation emails in folder: {folder_dict['confirmation']}")
        confirmation_emails = get_email_content(mail, folder_dict['confirmation'], 'FROM "BestBuyInfo@emailinfo.bestbuy.com" SUBJECT "Thanks for your order"')
        print(f"Found {len(confirmation_emails)} confirmation emails")
        for num in confirmation_emails:
            email_date, order_number, _ = process_email(mail, num, 'confirmation')
            if order_number:
                orders.append([email_date, order_number, "", []])
                confirmation_count += 1
                print(f"Processed confirmation: Order {order_number} from {email_date}")

    # Process cancelled emails
    if folder_dict['cancelled']:
        print(f"\nProcessing cancellation emails in folder: {folder_dict['cancelled']}")
        cancelled_emails = get_email_content(mail, folder_dict['cancelled'], 'OR SUBJECT "Your Best Buy order has been canceled" SUBJECT "Your order has been canceled"')
        print(f"Found {len(cancelled_emails)} cancellation emails")
        for num in cancelled_emails:
            _, cancelled_order_number, _ = process_email(mail, num, 'cancelled')
            if cancelled_order_number:
                for order in orders:
                    if order[1] == cancelled_order_number:
                        order[2] = "Cancelled"
                        cancellation_count += 1
                        print(f"Processed cancellation: Order {cancelled_order_number}")
                        break

    # Process shipped emails
    if folder_dict['shipped']:
        print(f"\nProcessing shipped emails in folder: {folder_dict['shipped']}")
        shipped_emails = get_email_content(mail, folder_dict['shipped'], 'FROM "BestBuyInfo@emailinfo.bestbuy.com"')
        print(f"Found {len(shipped_emails)} shipped emails")
        for num in shipped_emails:
            _, shipped_order_number, tracking_numbers = process_email(mail, num, 'shipped')
            if shipped_order_number:
                for order in orders:
                    if order[1] == shipped_order_number:
                        order[2] = "Shipped"
                        order[3] = tracking_numbers
                        shipped_count += 1
                        tracking_numbers_count += len(tracking_numbers)
                        print(f"Processed shipped: Order {shipped_order_number} with {len(tracking_numbers)} tracking number(s)")
                        break

    mail.logout()
    print("\nClosing mail connection")

    with open('bestbuy_orders.csv', 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['Order Date', 'Order Number', 'Status'] + [f'Tracking Number {i+1}' for i in range(max(len(order[3]) for order in orders) if orders else 0)])
        for order in orders:
            writer.writerow(order[:3] + order[3])

    print(f"\nOrder information has been saved to 'bestbuy_orders.csv'")
    print(f"Total orders found: {len(orders)}")
    print(f"Confirmation emails processed: {confirmation_count}")
    print(f"Cancellation emails processed: {cancellation_count}")
    print("Successful orders: " + str(confirmation_count - cancellation_count))
    print(f"Shipped emails processed: {shipped_count}")
    print(f"Total tracking numbers found: {tracking_numbers_count}")


def main_google(EMAIL, PASSWORD):
    mail = connect_to_mail(EMAIL, PASSWORD, "imap.gmail.com", 993, use_ssl=True)

    print("Connecting to Gmail server...")

    orders = []
    confirmation_count = 0
    cancellation_count = 0
    shipped_count = 0
    tracking_numbers_count = 0

    # Process confirmation emails
    print("\nProcessing confirmation emails")
    confirmation_emails = get_email_content(mail, "INBOX", '(FROM "BestBuyInfo@emailinfo.bestbuy.com") (SUBJECT "Thanks for your order")')
    print(f"Found {len(confirmation_emails)} confirmation emails")
    for num in confirmation_emails:
        email_date, order_number, _ = process_email(mail, num, 'confirmation')
        if order_number:
            orders.append([email_date, order_number, "", []])
            confirmation_count += 1
            print(f"Processed confirmation: Order {order_number} from {email_date}")

    # Process cancelled emails
    print("\nProcessing cancellation emails")
    cancelled_emails = get_email_content(mail, "INBOX", '(FROM "BestBuyInfo@emailinfo.bestbuy.com") (OR (SUBJECT "Your Best Buy order has been canceled.") (SUBJECT "Your order has been canceled."))')
    print(f"Found {len(cancelled_emails)} cancellation emails")
    for num in cancelled_emails:
        _, cancelled_order_number, _ = process_email(mail, num, 'cancelled')
        if cancelled_order_number:
            for order in orders:
                if order[1] == cancelled_order_number:
                    order[2] = "Cancelled"
                    cancellation_count += 1
                    print(f"Processed cancellation: Order {cancelled_order_number}")
                    break

    # Process shipped emails
    print("\nProcessing shipped emails")
    shipped_emails = get_email_content(mail, "INBOX", '(FROM "BestBuyInfo@emailinfo.bestbuy.com") (SUBJECT "Your order will be shipped soon!")')
    print(f"Found {len(shipped_emails)} shipped emails")
    for num in shipped_emails:
        _, shipped_order_number, tracking_numbers = process_email(mail, num, 'shipped')
        if shipped_order_number:
            for order in orders:
                if order[1] == shipped_order_number:
                    order[2] = "Shipped"
                    order[3] = tracking_numbers
                    shipped_count += 1
                    tracking_numbers_count += len(tracking_numbers)
                    print(f"Processed shipped: Order {shipped_order_number} with {len(tracking_numbers)} tracking number(s)")
                    break

    mail.logout()
    print("\nClosing mail connection")

    with open('bestbuy_orders.csv', 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['Order Date', 'Order Number', 'Status'] + [f'Tracking Number {i+1}' for i in range(max(len(order[3]) for order in orders) if orders else 0)])
        for order in orders:
            writer.writerow(order[:3] + order[3])

    print(f"\nOrder information has been saved to 'bestbuy_orders.csv'")
    print(f"Total orders found: {len(orders)}")
    print(f"Confirmation emails processed: {confirmation_count}")
    print(f"Cancellation emails processed: {cancellation_count}")
    print("Successful orders: " + str(confirmation_count - cancellation_count))
    print(f"Shipped emails processed: {shipped_count}")
    print(f"Total tracking numbers found: {tracking_numbers_count}")

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
