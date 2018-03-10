from urllib.request import urlopen as ureg
from bs4 import BeautifulSoup as Soup

url = 'https://www.settrade.com/C04_02_stock_historical_p1.jsp?txtSymbol=LIT&ssoPageId=9&selectPage=2'

uClient = ureg(url)
page_html = uClient.read()
uClient.close()

page_soup = Soup(page_html, "html.parser")
price = page_soup.findAll("div", {"class": "col-xs-6 col-xs-offset-6 colorRed"})
date = page_soup.findAll("div", {"class": "stt-remark"})

text_price = price[0].text
date_text = date[0].text


# Print
print("Current price : {} ".format(text_price))
print("Current date : {} ".format(date_text))
