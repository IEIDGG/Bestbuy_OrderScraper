def read_credentials(filename):
    with open(filename, 'r') as f:
        lines = f.readlines()
        email = lines[0].split('=')[1].strip()
        password = lines[1].split('=')[1].strip()
        email_type = lines[2].split('=')[1].strip().lower() if len(lines) > 2 else 'gmail'

        if email_type not in ['gmail', 'proton', 'icloud']:
            email_type = 'gmail'

        return email, password, email_type

def print_summary(summary, confirmation_count, cancellation_count):
    unique_orders, total_orders, shipped_count, tracking_numbers_count = summary

    print(f"\nOrder information has been saved to the database")
    print(f"Total unique orders found: {unique_orders}")
    print(f"Total orders found: {total_orders}")
    print(f"Confirmation emails processed: {confirmation_count}")
    print(f"Cancellation emails processed: {cancellation_count}")
    print(f"Successful orders: {confirmation_count - cancellation_count}")
    print(f"Shipped orders: {shipped_count}")
    print(f"Total tracking numbers found: {tracking_numbers_count}")