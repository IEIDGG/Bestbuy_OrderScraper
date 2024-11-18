from email_connector import get_email_content
from email_processor import process_email


def process_confirmation_emails(mail, folder_name, protocol='RFC822'):
    orders = []
    confirmation_count = 0
    print("Using protocol: ", protocol)
    print(f"\nProcessing confirmation emails in folder: {folder_name}")

    def get_confirmation_emails(mail, folder_name):
        search_criteria = (
            '(OR '
            '(FROM "BestBuyInfo@emailinfo.bestbuy.com") '
            '(FROM "BestBuyInfo")'
            ') '
            'SUBJECT "Thanks for your order"'
        )
        return get_email_content(mail, folder_name, search_criteria)

    confirmation_emails = get_confirmation_emails(mail, folder_name)
    print(f"Found {len(confirmation_emails)} confirmation emails")

    for num in confirmation_emails:
        email_date, order_number, _, products, total_price, email_address = process_email(mail, num, 'confirmation', protocol)
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
    return orders, confirmation_count


def process_cancellation_emails(mail, folder_name, orders, protocol='RFC822'):
    cancellation_count = 0
    print(f"\nProcessing cancellation emails in folder: {folder_name}")
    cancelled_emails = get_email_content(mail, folder_name,
                                         'OR SUBJECT "Your Best Buy order has been canceled" SUBJECT "Your order has been canceled"')
    print(f"Found {len(cancelled_emails)} cancellation emails")

    for num in cancelled_emails:
        _, cancelled_order_number, _, _, _, _ = process_email(mail, num, 'cancelled', protocol)
        if cancelled_order_number:
            for order in orders:
                if order['number'] == cancelled_order_number:
                    order['status'] = "Cancelled"
                    cancellation_count += 1
                    print(f"Processed cancellation: Order {cancelled_order_number}")
                    break
    return cancellation_count


def process_shipped_emails(mail, folder_name, orders, protocol='RFC822'):
    shipped_count = 0
    tracking_numbers_count = 0
    print(f"\nProcessing shipped emails in folder: {folder_name}")

    def get_shipped_emails(mail, folder_name):
        search_criteria = (
            '(OR '
            '(SUBJECT "Your order will be shipped soon!") '
            '(SUBJECT "We have your tracking number.")'
            ')'
        )
        return get_email_content(mail, folder_name, search_criteria)

    shipped_emails = get_shipped_emails(mail, folder_name)
    print(f"Found {len(shipped_emails)} shipped emails")

    for num in shipped_emails:
        _, shipped_order_number, tracking_numbers, _, _, _ = process_email(mail, num, 'shipped', protocol)
        if shipped_order_number:
            for order in orders:
                if order['number'] == shipped_order_number:
                    order['status'] = "Shipped"
                    order['tracking'] = tracking_numbers
                    shipped_count += 1
                    tracking_numbers_count += len(tracking_numbers)
                    print(
                        f"Processed shipped: Order {shipped_order_number} with {len(tracking_numbers)} tracking number(s)")
                    break
    return shipped_count, tracking_numbers_count