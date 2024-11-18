import email
from datetime import datetime
from bs4 import BeautifulSoup
from order_parser import parse_product_details


def process_email(mail, num, email_type, protocol='RFC822'):
    try:
        if protocol == 'RFC5322':
            _, msg_data = mail.fetch(num, '(BODY[])')
        else:
            _, msg_data = mail.fetch(num, f'({protocol})')

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

                    order_tds = soup.find_all('td', style='padding-bottom:12px;')
                    for td in order_tds:
                        if 'Order number:' in td.text:
                            order_span = td.find('span', style=lambda
                                value: value and 'font-weight: 700' in value and 'font-size: 14px' in value)
                            if order_span:
                                order_number = order_span.text.strip()

                    if email_type == 'shipped':
                        tracking_spans = soup.find_all('span', style='font: bold 14px Arial')
                        for span in tracking_spans:
                            if 'Tracking #:' in span.text:
                                tracking_link = span.find('a')
                                if tracking_link:
                                    tracking_numbers.append(tracking_link.text.strip())

                        tracking_tds = soup.find_all('td', style='padding-bottom:12px;')
                        for td in tracking_tds:
                            if 'Tracking Number:' in td.text:
                                tracking_span = td.find('span', style=lambda
                                    value: value and 'font-weight: 700' in value and 'font-size: 14px' in value)
                                if tracking_span:
                                    tracking_numbers.append(tracking_span.text.strip())

        return email_date, order_number, tracking_numbers, products, total_price, email_address

    except Exception as e:
        print(f"Error processing email: {str(e)}")
        return None, None, [], [], "N/A", None
