from bs4 import BeautifulSoup

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