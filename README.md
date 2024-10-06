# Bestbuy_OrderScraper
Requires the latest version of python along with pip: https://www.python.org/downloads/

Requires BeautifulSoup: `pip install beautifulsoup4`

Edit credentials in `credentials.txt`. True for Proton Mail (Requires Proton Bridge) and False for Gmail.

iCloud you only need to change line `mail = connect_to_mail(EMAIL, PASSWORD, "imap.gmail.com", 993, use_ssl=True)` to `mail = connect_to_mail(EMAIL, PASSWORD, "imap.mail.me.com", 993, use_ssl=True)` and it should work.

To run the code with no editor, you can just open the file with Python and it should run. In order to catch errors, you will need an IDE.

The script will create a csv file with all the order details, and as a warning, it may miss some tracking numbers that may come with new BestBuy email style changes.
