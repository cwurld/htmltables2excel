import unittest
import os
import json
import datetime
import traceback
from operator import methodcaller
import re
import six

from bs4 import NavigableString, BeautifulSoup

import xlsxwriter




def load_html():
    with open('/home/chuck/sqdb/htmltables2excel/htmltables2excel/data_for_tests/test_data2.html', 'rb') as fp:
        html = fp.read()
    return html


def parse_table(table):
    rows = table.find_all('tr')
    parsed_rows = []
    for row in rows:
        parsed_row = []
        cells = row.find_all(['th', 'td'])
        for cell in cells:
            parsed_row.append(cell)
        parsed_rows.append(parsed_row)
    return parsed_rows


def parse(html):
    soup = BeautifulSoup(html, 'html.parser')
    tables = soup.find_all('table')

    results = []
    for table in tables:
        results.append(parse_table(table))

    return results

def class_to_format(cell, allowedclasses):

    if 'class' not in cell.attrs:
        return None
    for i in cell.attrs['class']:
        if i in allowedclasses[cell.name]:
            return allowedclasses[cell.name][i]
    return None


def write_excel(parsed_table):
    workbook = xlsxwriter.Workbook('test_excel.xlsx')
    allowedclasses = {
        'th': {
            'red': workbook.add_format({'font_color': 'red', 'bold':True})
        },
        'td': {
            'red': workbook.add_format({'font_color': 'red'})
        }
    }

    worksheet = workbook.add_worksheet(name='Sheet1')
    worksheet.merge_range(0, 0, 0, 10, 'The Page Title')

    rownum = 2
    for row in parsed_table:
        for col, cell in enumerate(row):
            worksheet.write(rownum, col, cell.string, class_to_format(cell, allowedclasses))
        rownum += 1

    workbook.close()


def main():
    html = load_html()
    parsed_table = parse(html)
    write_excel(parsed_table[0])


main()
