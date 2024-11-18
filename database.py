import sqlite3
import os

DB_FILE = 'bestbuy_orders.sqlite3'

def create_connection():
    if not os.path.exists(DB_FILE):
        conn = sqlite3.connect(DB_FILE)
        create_tables(conn)
    else:
        conn = sqlite3.connect(DB_FILE)
        create_tables(conn)

    return conn


def create_tables(conn):
    cursor = conn.cursor()

    cursor.executescript('''
    CREATE TABLE IF NOT EXISTS orders (
        order_number TEXT PRIMARY KEY,
        order_date TEXT,
        total_price TEXT,
        status TEXT,
        email_address TEXT
    );

    CREATE TABLE IF NOT EXISTS products (
        id INTEGER PRIMARY KEY,
        order_id TEXT,
        title TEXT,
        price TEXT,
        quantity TEXT,
        FOREIGN KEY (order_id) REFERENCES orders (order_number)
    );

    CREATE TABLE IF NOT EXISTS tracking_numbers (
        id INTEGER PRIMARY KEY,
        order_id TEXT,
        tracking_number TEXT,
        FOREIGN KEY (order_id) REFERENCES orders (order_number)
    );

    CREATE TABLE IF NOT EXISTS successful_orders (
        order_number TEXT PRIMARY KEY,
        order_date TEXT,
        total_price TEXT,
        status TEXT,
        title TEXT,
        quantity TEXT,
        tracking_number TEXT
    );
    ''')

    conn.commit()


def insert_order(conn, order):
    cursor = conn.cursor()

    cursor.execute('SELECT * FROM orders WHERE order_number = ?', (order['number'],))
    existing_order = cursor.fetchone()

    if existing_order:
        cursor.execute('''
        UPDATE orders 
        SET order_date = ?, total_price = ?, status = ?, email_address = ?
        WHERE order_number = ?
        ''', (order['date'], order['total_price'], order['status'], order['email_address'], order['number']))
    else:
        cursor.execute('''
        INSERT INTO orders (order_number, order_date, total_price, status, email_address)
        VALUES (?, ?, ?, ?, ?)
        ''', (order['number'], order['date'], order['total_price'], order['status'], order['email_address']))

    order_id = order['number']

    cursor.execute('DELETE FROM products WHERE order_id = ?', (order_id,))
    cursor.execute('DELETE FROM tracking_numbers WHERE order_id = ?', (order_id,))

    for product in order['products']:
        cursor.execute('''
        INSERT INTO products (order_id, title, price, quantity)
        VALUES (?, ?, ?, ?)
        ''', (order_id, product['title'], product['price'], product['quantity']))

    for tracking_number in order['tracking']:
        cursor.execute('''
        INSERT INTO tracking_numbers (order_id, tracking_number)
        VALUES (?, ?)
        ''', (order_id, tracking_number))

    conn.commit()


def save_orders_to_db(orders):
    conn = create_connection()

    for order in orders:
        insert_order(conn, order)

    create_successful_orders_table(conn)
    conn.commit()
    return conn


def get_order_summary(conn):
    cursor = conn.cursor()
    cursor.execute('''
    SELECT COUNT(DISTINCT order_number) as unique_orders,
           COUNT(*) as total_orders,
           SUM(CASE WHEN status = 'Shipped' THEN 1 ELSE 0 END) as shipped_count,
           (SELECT COUNT(*) FROM tracking_numbers) as tracking_numbers_count
    FROM orders
    ''')
    return cursor.fetchone()


def create_successful_orders_table(conn):
    cursor = conn.cursor()

    cursor.execute('DROP TABLE IF EXISTS successful_orders')

    cursor.execute('''
    CREATE TABLE successful_orders (
        order_number TEXT PRIMARY KEY,
        order_date TEXT,
        total_price TEXT,
        status TEXT,
        title TEXT,
        quantity TEXT,
        tracking_number TEXT
    )
    ''')

    cursor.execute('''
    INSERT OR REPLACE INTO successful_orders (order_number, order_date, total_price, status, title, quantity, tracking_number)
    SELECT 
        o.order_number,
        o.order_date,
        o.total_price,
        o.status,
        GROUP_CONCAT(p.title, '; ') as title,
        GROUP_CONCAT(p.quantity, '; ') as quantity,
        GROUP_CONCAT(t.tracking_number, '; ') as tracking_number
    FROM 
        orders o
    LEFT JOIN 
        products p ON o.order_number = p.order_id
    LEFT JOIN 
        tracking_numbers t ON o.order_number = t.order_id
    WHERE 
        o.status != 'Cancelled'
    GROUP BY 
        o.order_number
    ''')

    conn.commit()


def get_successful_orders(conn):
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM successful_orders')
    return cursor.fetchall()