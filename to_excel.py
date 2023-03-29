import xmltodict
from openpyxl import Workbook

# create new Excel file and prepare headers of columns
wb = Workbook()
ws = wb.active
ws.append(['ID', 'Name', 'Price', 'Old price', 'Stock', 'Description', 'IMG', 'Params[beta]'])

# get data from XML file
with open("rozetkaxml.xml", "r", encoding='UTF-8') as f:
    xml_string = f.read()

xml_dict = xmltodict.parse(xml_string)

# pass through the elements and write data to an Excel table
for offer in xml_dict['yml_catalog']['shop']['offers']['offer']:
    # get info about offer by elements
    vendor = offer['vendorCode']
    name = offer['name']
    price = offer['price']
    if 'price_old' not in offer:
        price_old = offer['price']
    else:
        price_old = offer['price_old']
    description = offer['description']
    photo_url = offer['picture']
    stock = offer['stock_quantity']

    # add data to Excel
    ws.append([vendor, name, price, price_old, stock, description, photo_url])

# save Excel file
wb.save('offers.xlsx')
