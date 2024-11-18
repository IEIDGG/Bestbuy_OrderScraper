import csv
from database import save_orders_to_db, create_successful_orders_table, get_successful_orders, get_order_summary


def save_to_csv(orders):
    with open('bestbuy_orders.csv', 'w', newline='') as csvfile:
        fieldnames = ['order_number', 'order_date', 'total_price', 'status', 'email_address', 'products',
                      'tracking_numbers']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for order in orders:
            products_str = '; '.join(
                [f"{p['title']} (Qty: {p['quantity']}, Price: {p['price']})" for p in order['products']])
            writer.writerow({
                'order_number': order['number'],
                'order_date': order['date'],
                'total_price': order['total_price'],
                'status': order['status'],
                'email_address': order['email_address'],
                'products': products_str,
                'tracking_numbers': ', '.join(order['tracking'])
            })


def save_and_display_orders(orders):
    save_to_csv(orders)
    print("Orders saved to CSV file")

    conn = save_orders_to_db(orders)
    print("Orders saved to SQLite database")

    create_successful_orders_table(conn)
    successful_orders = get_successful_orders(conn)

    print("\nSuccessful Orders:")
    for order in successful_orders:
        print(f"Order Number: {order[0]}")
        print(f"Order Date: {order[1]}")
        print(f"Total Price: {order[2]}")
        print(f"Status: {order[3]}")
        print(f"Products: {order[4]}")
        print(f"Quantities: {order[5]}")
        print(f"Tracking Numbers: {order[6] if order[6] else 'N/A'}")
        print("---")

    summary = get_order_summary(conn)
    conn.close()
    return summary