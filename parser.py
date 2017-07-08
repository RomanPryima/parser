# -*- coding: utf-8 -*-

from bs4 import BeautifulSoup
from datetime import datetime
from time import time
import re
import requests
import xlsxwriter

# secret
url = 'https://bombayshop.com.ua/admin/'
username = raw_input('input Login:')
password = raw_input('input password:')

start_execution = time()

# starting session  !Necessary
current_session = requests.Session()
current_session.mount(url, requests.adapters.HTTPAdapter(max_retries=5))


def getting_general_table_page_url(
        site_url, site_username, site_password, session):
    headers = {
        'User-Agent': 'Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9a3pre)'
    }
    main_page_response = session.post(site_url, {
        'username': site_username,
        'password': site_password,
        'headers': headers
    })

    main_page = BeautifulSoup(main_page_response.text.encode(
        'utf-8'), "html.parser")

    return main_page.find('li', id='sale').find(
        'a', text=u'Замовлення').get('href')


def getting_general_table_page(general_table_page_url):
    general_table_page_response = current_session.post(general_table_page_url)

    general_table_page = BeautifulSoup(general_table_page_response.text.encode(
        'utf-8'), "html.parser")
    return general_table_page


def getting_all_general_pages_urls(general_table_page):
    all_page_urls = []
    pages_urls = general_table_page.find('div', {'class': 'pagination'})
    common_part_of_url = '='.join(re.split(r'=', (pages_urls.find('a').get(
        'href')))[:-1])
    page_numbers = []
    for url_item in pages_urls.findAll('a'):
        page_numbers.append(int(re.split(r'=', (url_item.get('href')))[-1]))
    for page_number in xrange(1, max(page_numbers) + 1):
        all_page_urls.append(common_part_of_url + '=' + str(page_number))
    return all_page_urls


def getting_id_link_dictionary(all_page_urls):
    id_link = {}
    for page_url in all_page_urls:
        general_table_page = getting_general_table_page(page_url)
        table_body = general_table_page.find('table', {'class': 'list'}).find(
            'tbody')
        for row in table_body.findAll('tr'):
            if row.find('input').get('value') != '':
                id_link[int(row.find('input').get('value'))] = str(row.find(
                    'a').get('href'))
    return id_link


def create_summary_dictionary(order_url):
    full_table_page = current_session.post(order_url).text
    table = BeautifulSoup(
        full_table_page.encode('utf-8'), 'html.parser').findAll(
        'table', {'class': 'form'})
    summary_dictionary = {}
    summary_dictionary['buyer'] = table[0].find(
        'td', text=u'Покупець').next_sibling.next_sibling.string
    summary_dictionary['email'] = table[0].find(
        'td', text=u'E-mail:').next_sibling.next_sibling.string
    summary_dictionary['phone'] = table[0].find(
        'td', text=u'Телефон').next_sibling.next_sibling.string
    summary_dictionary['city'] = table[1].find(
        'td', text=u'Місто:').next_sibling.next_sibling.string
    summary_dictionary['order_date'] = table[0].find(
        'td', text=u'Дата замовлення:').next_sibling.next_sibling.string
    summary_dictionary['sum'] = table[0].find(
        'td', text=u'Усього:').next_sibling.next_sibling.string
    summary_dictionary['summary_order_goods'] = filling_order_table(
        full_table_page)
    return summary_dictionary


def filling_order_table(_full_table_page):
    product_table_list = BeautifulSoup(
        _full_table_page.encode('utf-8'), 'html.parser').find(
        id='tab-product').find('tbody').findAll('tr')
    orders = []
    for row in product_table_list:
        order = {'good': row.find('td').find('a').string,
                 'manufacturer': row.findAll('td')[1].string,
                 'quantity': row.findAll('td')[2].string,
                 'price': row.findAll('td')[3].string}
        try:
            order['size'] = row.find('td').find('small').string
        except Exception:
            order['size'] = None
        orders.append(order)
    return orders


def creating_final_dictionary(id_links):
    final_dictionary = {}
    progress = float(0)
    for key, value in id_links.items():
        print ((progress / len(id_links) * 100), key)
        final_dictionary[key] = create_summary_dictionary(
             value)
        progress += 1
    return final_dictionary


def filling_xlsx(final_dictionary):
    workbook = xlsxwriter.Workbook(
        'bombay {}.xlsx'.format(datetime.now().strftime("%Y-%m-%d")))
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, 'ID')
    worksheet.write(0, 1, u"Прізвище та ім'я")
    worksheet.write(0, 2, u'Електронна адреса')
    worksheet.write(0, 3, u'Номер телефону')
    worksheet.write(0, 4, u'Адреса доставки')
    worksheet.write(0, 5, u'Дата замовлення')
    worksheet.write(0, 6, u'Сума')
    worksheet.write(0, 7, u'Замовлення')

    row = 1
    for key, value in final_dictionary.items():
        worksheet.write(row, 0, key)
        worksheet.write(row, 1, value.get('buyer'))
        worksheet.write(row, 2, value.get('email'))
        worksheet.write(row, 3, value.get('phone'))
        worksheet.write(row, 4, value.get('city'))
        worksheet.write(row, 5, value.get('order_date'))
        worksheet.write(row, 6, value.get('sum'))
        column = 7
        goods = value['summary_order_goods']
        for good in goods:
            worksheet.write(row, column, u'Товар:')
            worksheet.write(row, column + 1, good.get('good'))
            worksheet.write(row, column + 2, good.get('size'))
            worksheet.write(row, column + 3, u'Кількість:')
            worksheet.write(row, column + 4, good.get('quantity'))
            worksheet.write(row, column + 5, u'Ціна:')
            worksheet.write(row, column + 6, good.get('price'))
            column += 7
        row += 1
    workbook.close()

start_url = getting_general_table_page_url(url, username, password,
                                           current_session)
first_general_page = getting_general_table_page(start_url)

all_general_urls = getting_all_general_pages_urls(
    first_general_page)

all_id_links = getting_id_link_dictionary(all_general_urls)

filling_xlsx(creating_final_dictionary(all_id_links))
#
# workbook = xlsxwriter.Workbook(
#         'test{}.xlsx'.format(datetime.now().strftime("%Y-%m-%d")))
# worksheet = workbook.add_worksheet()
# row = 0
# for key in all_general_urls:
#     worksheet.write(row, 0, key)
#     #worksheet.write(row, 1, value)
#     row += 1
# workbook.close()

print ('Execution finished in {} sec.'.format(time() - start_execution))

