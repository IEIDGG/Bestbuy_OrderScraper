import imaplib
import email
from email.header import decode_header
import csv
import re
from datetime import datetime
import ssl
from bs4 import BeautifulSoup


def main_proton(EMAIL, PASSWORD, cancellation):
    imap_server = "127.0.0.1"
    imap_port = 1143  # or whatever port you've set in ProtonMail Bridge

    context = ssl.create_default_context()
    context.check_hostname = False
    context.verify_mode = ssl.CERT_NONE

    mail = imaplib.IMAP4(imap_server, imap_port)

    mail.starttls(ssl_context=context)

    mail.login(EMAIL, PASSWORD)

    _, folders = mail.list()
    bestbuy_confirmation_folder = None
    bestbuy_cancelled_folder = None
    for folder in folders:
        if b'Bestbuy-Confirmation' in folder:
            bestbuy_confirmation_folder = folder.decode().split('"')[-2]
        if b'Bestbuy-Cancelled' in folder:
            bestbuy_cancelled_folder = folder.decode().split('"')[-2]

    if not bestbuy_confirmation_folder:
        print("Bestbuy-Confirmation folder not found. Available folders:")
        for folder in folders:
            print(folder.decode())
        mail.logout()
        return

    print(f"Bestbuy-Confirmation folder found: {bestbuy_confirmation_folder}")

    # List to store order information
    orders = []

    # Process Bestbuy-Confirmation folder
    mail.select(bestbuy_confirmation_folder)

    try:
        # Search for emails from BestBuy with the specific subject
        _, message_numbers = mail.search(None,
                                         'FROM "BestBuyInfo@emailinfo.bestbuy.com" SUBJECT "Thanks for your order"')
        print(f"Number of matching emails in Confirmation folder: {len(message_numbers[0].split())}")

        # Process each email
        for num in message_numbers[0].split():
            try:
                _, msg_data = mail.fetch(num, "(RFC822)")
                email_body = msg_data[0][1]
                email_message = email.message_from_bytes(email_body)

                # Get the email date
                date_tuple = email.utils.parsedate_tz(email_message['Date'])
                if date_tuple:
                    local_date = datetime.fromtimestamp(email.utils.mktime_tz(date_tuple))
                    email_date = local_date.strftime("%Y-%m-%d")
                else:
                    email_date = "Unknown"

                # Extract order number from email body
                order_number = None
                for part in email_message.walk():
                    if part.get_content_type() == "text/html":
                        html_body = part.get_payload(decode=True).decode()
                        soup = BeautifulSoup(html_body, 'html.parser')
                        order_span = soup.find('span', string=re.compile(r'BBY01-\d+'))
                        if order_span:
                            order_number = order_span.text.strip()
                        break

                if order_number:
                    orders.append([email_date, order_number, ""])

                print(f"Processed email {num}: Date={email_date}, Order Number={order_number}")
            except Exception as e:
                print(f"Error processing email {num}: {str(e)}")
    except Exception as e:
        print(f"Error searching emails: {str(e)}")

    # Process Bestbuy-Cancelled folder if enabled
    if cancellation and bestbuy_cancelled_folder:
        mail.select(bestbuy_cancelled_folder)

        try:
            _, cancelled_messages = mail.search(None,
                                                'SUBJECT "Your Best Buy order has been canceled"')
            print(f"Number of cancelled order emails: {len(cancelled_messages[0].split())}")

            for num in cancelled_messages[0].split():
                try:
                    _, msg_data = mail.fetch(num, "(RFC822)")
                    email_body = msg_data[0][1]
                    email_message = email.message_from_bytes(email_body)

                    # Extract order number from cancelled email
                    cancelled_order_number = None
                    for part in email_message.walk():
                        if part.get_content_type() == "text/html":
                            html_body = part.get_payload(decode=True).decode()
                            soup = BeautifulSoup(html_body, 'html.parser')
                            order_span = soup.find('span', style='font: bold 23px Arial; color: #1d252c;')
                            if order_span:
                                cancelled_order_number = order_span.text.strip().replace("Order #", "")
                            break

                    if cancelled_order_number:
                        # Mark the order as cancelled in the orders list
                        for order in orders:
                            if order[1] == cancelled_order_number:
                                order[2] = "Cancelled"
                                break

                    print(f"Processed cancelled order: {cancelled_order_number}")
                except Exception as e:
                    print(f"Error processing cancelled email {num}: {str(e)}")
        except Exception as e:
            print(f"Error searching cancelled emails: {str(e)}")

    # Close the connection
    mail.close()
    mail.logout()

    # Write orders to CSV file
    with open('bestbuy_orders.csv', 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['Order Date', 'Order Number', 'Status'])
        writer.writerows(orders)

    print(f"Order information has been saved to 'bestbuy_orders.csv'")
    print(f"Total orders found: {len(orders)}")


def main_google(EMAIL, PASSWORD, cancellation):
    # Connect to Gmail IMAP server
    imap_server = "imap.gmail.com"
    imap_port = 993

    mail = imaplib.IMAP4_SSL(imap_server, imap_port)

    mail.login(EMAIL, PASSWORD)

    mail.select("inbox")

    orders = []

    try:
        _, message_numbers = mail.search(None,
                                         'FROM "BestBuyInfo@emailinfo.bestbuy.com" SUBJECT "Thanks for your order"')
        print(f"Number of matching emails in Confirmation folder: {len(message_numbers[0].split())}")

        # Process each email
        for num in message_numbers[0].split():
            try:
                _, msg_data = mail.fetch(num, "(RFC822)")
                email_body = msg_data[0][1]
                email_message = email.message_from_bytes(email_body)

                # Get the email date
                date_tuple = email.utils.parsedate_tz(email_message['Date'])
                if date_tuple:
                    local_date = datetime.fromtimestamp(email.utils.mktime_tz(date_tuple))
                    email_date = local_date.strftime("%Y-%m-%d")
                else:
                    email_date = "Unknown"

                # Extract order number from email body
                order_number = None
                for part in email_message.walk():
                    if part.get_content_type() == "text/html":
                        html_body = part.get_payload(decode=True).decode()
                        soup = BeautifulSoup(html_body, 'html.parser')
                        order_span = soup.find('span', string=re.compile(r'BBY01-\d+'))
                        if order_span:
                            order_number = order_span.text.strip()
                        break

                if order_number:
                    orders.append([email_date, order_number, ""])

                print(f"Processed email {num}: Date={email_date}, Order Number={order_number}")
            except Exception as e:
                print(f"Error processing email {num}: {str(e)}")
    except Exception as e:
        print(f"Error searching emails: {str(e)}")

    if cancellation:
        try:
            _, cancelled_messages = mail.search(None,
                                                'FROM "BestBuyInfo@emailinfo.bestbuy.com" SUBJECT "Your Best Buy order has been canceled"')
            print(f"Number of cancelled order emails: {len(cancelled_messages[0].split())}")

            for num in cancelled_messages[0].split():
                try:
                    _, msg_data = mail.fetch(num, "(RFC822)")
                    email_body = msg_data[0][1]
                    email_message = email.message_from_bytes(email_body)

                    # Extract order number from cancelled email
                    cancelled_order_number = None
                    for part in email_message.walk():
                        if part.get_content_type() == "text/html":
                            html_body = part.get_payload(decode=True).decode()
                            soup = BeautifulSoup(html_body, 'html.parser')
                            order_span = soup.find('span', style='font: bold 23px Arial; color: #1d252c;')
                            if order_span:
                                cancelled_order_number = order_span.text.strip().replace("Order #", "")
                            break

                    if cancelled_order_number:
                        # Mark the order as cancelled in the orders list
                        for order in orders:
                            if order[1] == cancelled_order_number:
                                order[2] = "Cancelled"
                                break

                    print(f"Processed cancelled order: {cancelled_order_number}")
                except Exception as e:
                    print(f"Error processing cancelled email {num}: {str(e)}")
        except Exception as e:
            print(f"Error searching cancelled emails: {str(e)}")

    mail.close()
    mail.logout()

    with open('bestbuy_orders.csv', 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['Order Date', 'Order Number', 'Status'])
        writer.writerows(orders)

    print(f"Order information has been saved to 'bestbuy_orders.csv'")
    print(f"Total orders found: {len(orders)}")


def read_credentials(file_path):
    try:
        with open(file_path, 'r') as file:
            lines = file.readlines()
            if len(lines) >= 2:
                user = lines[0].split('=')[1].strip()
                password = lines[1].split('=')[1].strip()
                proton = lines[2].split('=')[1].strip()
                if proton == "TRUE" or proton == "True":
                    proton = True
                if proton == "FALSE" or proton == "False":
                    proton = False
                cancellation = lines[3].split('=')[1].strip()
                return user, password, proton, cancellation
            else:
                raise ValueError("The file should contain at least two lines (email and password).")
    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found.")
        raise
    except Exception as e:
        print(f"Error in read_credentials: {e}")
        raise


if __name__ == "__main__":
    file_path = 'credentials.txt'
    user_check, password_check, proton_check, cancellation_check = read_credentials(file_path)
    if proton_check is True:
        main_proton(user_check, password_check, cancellation_check)
    if proton_check is False:
        main_google(user_check, password_check, cancellation_check)
