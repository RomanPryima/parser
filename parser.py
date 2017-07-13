# -*- coding: utf-8 -*-

from bs4 import BeautifulSoup
from datetime import datetime
from time import time
from urlparse import urlparse, parse_qs
import re
import requests
import xlsxwriter

# secret
url = 'https://bombayshop.com.ua/admin/'
username = raw_input('input Login:')
password = raw_input('input password:')
headers = {'User-Agent':
               'Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9a3pre)'}
authentication = dict(
    username=username, password=password, headers=headers)

start_execution = time()


class Session(requests.Session):

    def login(self, site_url, authentication_data):
        self.logined_url = self.post(site_url, authentication_data).url
        self.base_url = site_url

    def get_token(self):
        self.token = '&token=' + parse_qs(urlparse(
            self.logined_url).query).get('token')[0]

    def get_top_number_of_general_page(self):
        route = 'index.php?route=sale/order'
        print (self.base_url + route + self.token)
        general_table_page_response = self.post(
            self.base_url + route + self.token)
        general_table_page = BeautifulSoup(
            general_table_page_response.text.encode('utf-8'), "html.parser")
        pages_urls = general_table_page.find('div', {'class': 'pagination'})
        page_numbers = []
        for url_item in pages_urls.findAll('a'):
            page_numbers.append(int(re.split(r'=', (
                url_item.get('href')))[-1]))
        self.top_page_number = max(page_numbers)

    def get_id_list(self):
        route = 'index.php?route=sale/order'
        page = '&page='
        id_list = []
        for page_number in xrange(self.top_page_number):
            table_page = BeautifulSoup(self.post(
                self.base_url + route + self.token + page + str(
                    page_number+1)).text.encode('utf-8'), "html.parser")
            table_body = table_page.find('table', {'class': 'list'}).find(
                'tbody')
            for row in table_body.findAll('tr'):
                if row.find('input').get('value') != '':
                    id_list.append(int(row.find('input').get('value')))
        return id_list


def create_summary_dictionary(session, order_id):
    route = 'index.php?route=sale/order/info'
    order = '&order_id='
    order_url = session.base_url + route + session.token + order + str(
        order_id)
    full_table_page_response = session.post(order_url)
    full_table_page = full_table_page_response.text
    table = BeautifulSoup(
        full_table_page.encode('utf-8'), 'html.parser').findAll(
        'table', {'class': 'form'})
    summary_dictionary = dict()
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


def filling_order_table(full_table_page):
    product_table_list = BeautifulSoup(
        full_table_page.encode('utf-8'), 'html.parser').find(
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


def creating_final_dictionary(session, id_links):
    final_dictionary = {}
    progress = float(0)
    used_id_links = []
    for order_id in id_links:
        print ((progress / len(id_links) * 100), order_id)
        final_dictionary[order_id] = create_summary_dictionary(session,
                                                               order_id)
        used_id_links.append(order_id)
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


# starting session  !Necessary
def start_session():
    session = Session()
    session.mount(url, requests.adapters.HTTPAdapter(max_retries=5))
    session.login(url, authentication)
    session.get_token()
    print (session.logined_url)
    print session.token
    session.get_top_number_of_general_page()
    print session.top_page_number
    return session


current_session = start_session()
id_list = current_session.get_id_list()

full_dictionary = creating_final_dictionary(current_session, id_list)

filling_xlsx(full_dictionary)

print ('Execution finished in {} sec.'.format(time() - start_execution))

