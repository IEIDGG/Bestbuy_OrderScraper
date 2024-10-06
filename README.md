# Bestbuy_OrderScraper
Requires the latest version of python along with pip: https://www.python.org/downloads/

Requires BeautifulSoup: `pip install beautifulsoup4`

Edit credentials in `credentials.txt`. True for Proton Mail (Requires Proton Bridge) and False for Gmail.

iCloud you only need to change line `mail = connect_to_mail(EMAIL, PASSWORD, "imap.gmail.com", 993, use_ssl=True)` to `mail = connect_to_mail(EMAIL, PASSWORD, "imap.mail.me.com", 993, use_ssl=True)` and it shoudl work.
