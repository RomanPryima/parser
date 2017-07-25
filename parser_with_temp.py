# -*- coding: utf-8 -*-
"""
General description:
The module is purposed for grabbing certain data from tables on different pages
in the admin part of aimed site.
First page with general table can be requested with route:"route=sale/order"
There are two types of tables needed to be parsed:
1. General tables wich contain general data of an orders. Only order id's to
be grabbed from these tables. Route for these pages "route=sale/order",
"&page=" and number of the page.
2.Full tables page wich contain detailed data of an order can be reached with
route "index.php?route=sale/order", "&order_id=" and numerical order id grabbed
from the general tables. All necessary data is received from full tables.

Module iterates a list with numbers of general tables pages,
takes all order id's than iterates a list with order id's calling pages with
full tables and grabbing data from the tables into one full data dictionary.

When dictionary is completed module creates an .xlsx file where inserts data
from the dictionary in certain order.
"""

from bs4 import BeautifulSoup
from datetime import datetime
from os import remove
from time import time
from urlparse import urlparse, parse_qs
import ast
import re
import requests
import sys
import xlsxwriter


class Session(requests.Session):
    """Class inherited from requests.Session object and has additional methods
    """

    def login(self, site_url, authentication_data):
        """Sends a request for getting response with a token from the server
        creates session properties containing base url and url with token for
        future usage"""
        try:
            self.base_url = site_url
            self.logined_url = self.post(
                self.base_url, authentication_data).url
        except requests.adapters.ConnectionError:
            sys.exit(
                'Connection issues. Check connection and try again')

    def get_token(self):
        """parses token from the response and creates session token property"""
        try:
            self.token = '&token=' + parse_qs(urlparse(
                self.logined_url).query).get('token')[0]
        except TypeError:
            sys.exit(
                'Login issues. Please enter properly username and password.')

    def get_top_number_of_general_page(self):
        """ sends request for the first page containing general table and grabs
        the top number of pages containing general tables. Applies
        top_page_number to the session"""
        route = 'index.php?route=sale/order'
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
        """Iterates a list up to top_page_number, calling pages with general
        tables and grabs all order id's into a list with all orders id's
        :return a list with all orders numbers.
        """
        route = 'index.php?route=sale/order'
        page = '&page='
        _id_list = []
        for page_number in range(self.top_page_number):
            table_page = BeautifulSoup(self.post(
                self.base_url + route + self.token + page + str(
                    page_number + 1)).text.encode('utf-8'), "html.parser")
            table_body = table_page.find('table', {'class': 'list'}).find(
                'tbody')
            for row in table_body.findAll('tr'):
                if row.find('input').get('value') != '':
                    _id_list.append(int(row.find('input').get('value')))
        return _id_list


def create_summary_dictionary(session, order_id):
    """
    Requests a page with full table, using session and order id.
    Calls filling_order_table, wich returns detailed data of ordered goods.
    Grabs necessary data into the summary dictionary.
    :param session: Session object
    :param order_id: int - number of desired order.
    :return: summarized dictionary with full data about the order
    """
    route = 'index.php?route=sale/order/info'
    order = '&order_id='
    order_url = session.base_url + route + session.token + order + str(
        order_id)
    with session.post(order_url) as full_table_page_response:
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
    indexed_dictionary = {}
    indexed_dictionary[order_id] = summary_dictionary
    with open('temp.txt', 'a') as temp:
        temp.write(str(indexed_dictionary) + '\n')


def filling_order_table(full_table_page):
    """
    Returns a list of dictionaries with all ordered goods and their properties
    in certain order.
    :param full_table_page:
    :return: list of dictionaries with all ordered goods and their properties.
    """
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
        except AttributeError:
            order['size'] = None
        orders.append(order)
    return orders


def creating_final_dictionary(session, _id_list):
    """
     Using session, iterates a list of id's, calling create_summary_dictionary.
    Joins results into one big final dictionary.
    :param session: current session
    :param _id_list: a list of all id's
    """
    progress = float(0)
    for order_id in _id_list:
        print ("Processed {} %. Processing order No. {}".format(
            round((progress / len(_id_list) * 100), 2), order_id))
        create_summary_dictionary(session, order_id)
        progress += 1


def filling_xlsx():
    """
    Creates an .xlxs file and fills it with data from temp file.
    :return: filled .xlsx file.
    """
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
    with open('temp.txt', 'r') as temp:
        lines = [line.rstrip('\n') for line in temp]
        for line in lines:
            final_dictionary = ast.literal_eval(line)
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
def start_session(url, authentication):
    """
    1.creates a new Session inherited by requests.Session object,
    2.mounts adapters to session (I'm not sure if ti is necessary)
    3. calls own method "login" and sends there site's url and authentication
    data.
    4. calls own method "get_token", wich returns a token string for adding it
    to next requests.
    5. calls own method "get_top_number_of_general_page, wich returns the
    highest number for using in route for navigation in general tables pages.
    :return: session with additional properties.
    """
    session = Session()
    session.mount(url, requests.adapters.HTTPAdapter(max_retries=5))
    session.login(url, authentication)
    session.get_token()
    session.get_top_number_of_general_page()
    return session


def run_parser(username, password):
    print ('Start parsing')
    url = 'https://bombayshop.com.ua/admin/'
    headers = {'User-Agent':
                   'Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9a3pre)'}
    authentication = dict(
        username=username, password=password, headers=headers)

    start_execution = time()
    try:
        remove('temp.txt')
    except OSError:
        pass
    current_session = start_session(url, authentication)
    id_list = sorted(current_session.get_id_list(), reverse=True)
    creating_final_dictionary(current_session, id_list)
    filling_xlsx()
    remove('temp.txt')
    print ('Execution finished in {} sec.'.format(time() - start_execution))
