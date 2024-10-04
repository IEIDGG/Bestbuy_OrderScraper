import imaplib
import email
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


def parse_product_details(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    product_sections = soup.find_all('td', style=lambda value: value and 'width:60%;max-width:359px;' in value)
    products = []

    for section in product_sections:
        title_tag = section.find('a', style='text-decoration: none;')
        if title_tag:
            title = title_tag.text.strip()
            qty_tag = section.find_next('td', string='Qty:')
            qty = qty_tag.find_next_sibling('td').text.strip() if qty_tag else "N/A"
            price_tag = section.find('span', string=lambda text: text and '$' in text, style=lambda
                value: value and 'font-weight: 700;font-size: 14px;line-height: 18px;' in value)
            price = price_tag.text.strip() if price_tag else "N/A"

            if price != "N/A":
                products.append({
                    'title': title,
                    'quantity': qty,
                    'price': price
                })

    total_td = soup.find('td', align='right', style=lambda value: value and 'padding-top:12px; padding-left:0;padding-right:0; padding-bottom:0; color:#000000;' in value)
    total_price = total_td.text.strip() if total_td else "N/A"

    return products, total_price


def process_email(mail, num, email_type):
    _, msg_data = mail.fetch(num, "(RFC822)")
    email_body = msg_data[0][1]
    email_message = email.message_from_bytes(email_body)

    email_address = email_message['To']

    date_tuple = email.utils.parsedate_tz(email_message['Date'])
    email_date = datetime.fromtimestamp(email.utils.mktime_tz(date_tuple)).strftime(
        "%Y-%m-%d") if date_tuple else "Unknown"

    order_number = None
    tracking_numbers = []
    products = []
    total_price = "N/A"

    for part in email_message.walk():
        if part.get_content_type() == "text/html":
            html_body = part.get_payload(decode=True).decode()
            soup = BeautifulSoup(html_body, 'html.parser')

            if email_type == 'confirmation':
                order_span = soup.find('span', string=lambda text: text and 'BBY01-' in text)
                if order_span:
                    order_number = order_span.text.strip()
                products, total_price = parse_product_details(html_body)
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

    return email_date, order_number, tracking_numbers, products, total_price, email_address


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
        confirmation_emails = get_email_content(mail, folder_dict['confirmation'],
                                                'FROM "BestBuyInfo@emailinfo.bestbuy.com" SUBJECT "Thanks for your order"')
        print(f"Found {len(confirmation_emails)} confirmation emails")
        for num in confirmation_emails:
            email_date, order_number, _, products, total_price, email_address = process_email(mail, num, 'confirmation')
            if order_number:
                orders.append({
                    'date': email_date,
                    'number': order_number,
                    'status': "",
                    'tracking': [],
                    'products': products,
                    'total_price': total_price,
                    'email_address': email_address
                })
                confirmation_count += 1
                print(f"Processed confirmation: Order {order_number} from {email_date} with {len(products)} products")

    # Process cancelled emails
    if folder_dict['cancelled']:
        print(f"\nProcessing cancellation emails in folder: {folder_dict['cancelled']}")
        cancelled_emails = get_email_content(mail, folder_dict['cancelled'],
                                             'OR SUBJECT "Your Best Buy order has been canceled" SUBJECT "Your order has been canceled"')
        print(f"Found {len(cancelled_emails)} cancellation emails")
        for num in cancelled_emails:
            _, cancelled_order_number, _, _, _, _ = process_email(mail, num, 'cancelled')
            if cancelled_order_number:
                for order in orders:
                    if order['number'] == cancelled_order_number:
                        order['status'] = "Cancelled"
                        cancellation_count += 1
                        print(f"Processed cancellation: Order {cancelled_order_number}")
                        break

    # Process shipped emails
    if folder_dict['shipped']:
        print(f"\nProcessing shipped emails in folder: {folder_dict['shipped']}")
        shipped_emails = get_email_content(mail, folder_dict['shipped'], 'FROM "BestBuyInfo@emailinfo.bestbuy.com"')
        print(f"Found {len(shipped_emails)} shipped emails")
        for num in shipped_emails:
            _, shipped_order_number, tracking_numbers, _, _, _ = process_email(mail, num, 'shipped')
            if shipped_order_number:
                for order in orders:
                    if order['number'] == shipped_order_number:
                        order['status'] = "Shipped"
                        order['tracking'].extend(tracking_numbers)  # Use extend instead of assign
                        shipped_count += 1
                        tracking_numbers_count += len(tracking_numbers)
                        print(
                            f"Processed shipped: Order {shipped_order_number} with {len(tracking_numbers)} tracking number(s)")
                        break

    mail.logout()
    print("\nClosing mail connection")

    with open('bestbuy_orders.csv', 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)

        # Write the header
        max_tracking = max(len(order['tracking']) for order in orders) if orders else 0
        header = ['Product', 'Price', 'Quantity', 'Total Price', 'Order Date', 'Order Number', 'Status']
        header.extend([f'Tracking Number {i + 1}' for i in range(max_tracking)])
        header.append('Email Address')
        writer.writerow(header)

        # Write the data
        for order in orders:
            if order['products']:
                for product in order['products']:
                    row = [
                        product['title'],
                        product['price'],
                        product['quantity'],
                        order['total_price'],
                        order['date'],
                        order['number'],
                        order['status']
                    ]
                    row.extend(order['tracking'] + [''] * (max_tracking - len(order['tracking'])))
                    row.append(order['email_address'])
                    writer.writerow(row)
            else:
                row = ['N/A', 'N/A', 'N/A', order['total_price'], order['date'], order['number'], order['status']]
                row.extend(order['tracking'] + [''] * (max_tracking - len(order['tracking'])))
                row.append(order['email_address'])
                writer.writerow(row)

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
    confirmation_emails = get_email_content(mail, "INBOX",
                                            '(FROM "BestBuyInfo@emailinfo.bestbuy.com") (SUBJECT "Thanks for your order")')
    print(f"Found {len(confirmation_emails)} confirmation emails")
    for num in confirmation_emails:
        email_date, order_number, _, products, total_price, email_address = process_email(mail, num, 'confirmation')
        if order_number:
            orders.append({
                'date': email_date,
                'number': order_number,
                'status': "",
                'tracking': [],
                'products': products,
                'total_price': total_price,
                'email_address': email_address
            })
            confirmation_count += 1
            print(f"Processed confirmation: Order {order_number} from {email_date} with {len(products)} products")

    # Process cancelled emails
    print("\nProcessing cancellation emails")
    cancelled_emails = get_email_content(mail, "INBOX",
                                         '(FROM "BestBuyInfo@emailinfo.bestbuy.com") (OR (SUBJECT "Your Best Buy order has been canceled.") (SUBJECT "Your order has been canceled."))')
    print(f"Found {len(cancelled_emails)} cancellation emails")
    for num in cancelled_emails:
        _, cancelled_order_number, _, _, _ = process_email(mail, num, 'cancelled')
        if cancelled_order_number:
            for order in orders:
                if order['number'] == cancelled_order_number:
                    order['status'] = "Cancelled"
                    cancellation_count += 1
                    print(f"Processed cancellation: Order {cancelled_order_number}")
                    break

    # Process shipped emails
    print("\nProcessing shipped emails")
    shipped_emails = get_email_content(mail, "INBOX",
                                       '(FROM "BestBuyInfo@emailinfo.bestbuy.com") (SUBJECT "Your order will be shipped soon!")')
    print(f"Found {len(shipped_emails)} shipped emails")
    for num in shipped_emails:
        _, shipped_order_number, tracking_numbers, _, _, _ = process_email(mail, num, 'shipped')
        if shipped_order_number:
            for order in orders:
                if order['number'] == shipped_order_number:
                    order['status'] = "Shipped"
                    order['tracking'] = tracking_numbers
                    print(
                        f"Processed shipped: Order {shipped_order_number} with {len(tracking_numbers)} tracking number(s)")
                    break

    mail.logout()
    print("\nClosing mail connection")

    with open('bestbuy_orders.csv', 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)

        # Write the header
        max_tracking = max(len(order['tracking']) for order in orders) if orders else 0
        header = ['Product', 'Price', 'Quantity', 'Total Price', 'Order Date', 'Order Number', 'Status']
        header.extend([f'Tracking Number {i + 1}' for i in range(max_tracking)])
        header.append('Email Address')
        writer.writerow(header)

        # Write the data
        for order in orders:
            if order['products']:
                for product in order['products']:
                    row = [
                        product['title'],
                        product['price'],
                        product['quantity'],
                        order['total_price'],
                        order['date'],
                        order['number'],
                        order['status']
                    ]
                    row.extend(order['tracking'] + [''] * (max_tracking - len(order['tracking'])))
                    row.append(order['email_address'])
                    writer.writerow(row)
            else:
                row = ['N/A', 'N/A', 'N/A', order['total_price'], order['date'], order['number'], order['status']]
                row.extend(order['tracking'] + [''] * (max_tracking - len(order['tracking'])))
                row.append(order['email_address'])
                writer.writerow(row)

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
    input("ENTER to exit")
