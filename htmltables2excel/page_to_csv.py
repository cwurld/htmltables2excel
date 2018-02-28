import unittest
import os
import csv

from bs4 import BeautifulSoup


def remove_dollar_sign(s):
    if s and s[0] == '$':
        return s[1:]
    else:
        return s


def clean_cell(cell):
    if cell.string:
        s = str(cell.string).strip()
        if s is None or s == u'\n' or len(s) == 0:
            return ''
        else:
            return remove_dollar_sign(s)
    else:
        s = remove_dollar_sign(''.join([x for x in cell.stripped_strings]))
        return s


def parse_table(table):
    rows = []
    if table.caption:
        title = table.caption.string
        rows.append([title])

    header = [x.string for x in table.thead.tr.find_all('th')]
    rows.append(header)

    table_rows = table.tbody.find_all('tr')
    if table.tfoot:
        table_rows += table.tfoot.find_all('tr')

    for row in table_rows:
        csv_row = []
        for cell in row.find_all('td'):
            colspan = int(cell.attrs.get('colspan', '0'))
            for i in range(max(0, colspan - 1)):
                csv_row.append(u'')

            s = clean_cell(cell)
            csv_row.append(s)
        rows.append(csv_row)
    rows.append([])
    return rows


def parse_row(row, cell_type):
    result = []
    for cell in row.find_all(cell_type):
        colspan = int(cell.attrs.get('colspan', '0'))
        for i in range(max(0, colspan - 1)):
            result.append(u'')

        s = clean_cell(cell)
        result.append(s)
    return result


def parse_tables(html, excluded_tables, parse_table_func):
    soup = BeautifulSoup(html, 'html.parser')
    tables = soup.find_all('table')

    parsed_tables = []
    for table in tables:
        table_id = table.attrs.get(u'id')
        if table_id not in excluded_tables:
            parsed_tables.append(parse_table_func(table))
    return parsed_tables


def page_to_csv(file_full_path, html,  extra_headers=None, excluded_tables=None):
    """
    Page can contain one or more tables. The tables need to be well structured (eg. thead, tbody. tfoot)

    :param file_full_path:
    :param html:
    :param extra_headers: a list of lists of header text
    :param excluded_tables: a list of table ids to be excluded.
    :return:
    """
    csv_rows = extra_headers or []
    parsed_tables = parse_tables(html, excluded_tables or [], parse_table)

    for table in parsed_tables:
        csv_rows += table

    fp = open(file_full_path, 'wb')
    writer = csv.writer(fp)
    for row in csv_rows:
        writer.writerow(row)
    fp.close()
    return csv_rows


class TestPageToCSV(unittest.TestCase):
    def setUp(self):
        fp = open(os.path.join('data_for_tests/table_to_csv_test_data.html'), 'rb')
        self.html = fp.read()
        fp.close()

    def test_1(self):
        path = 'page_to_csv.csv'
        rows = page_to_csv(path, self.html)
        self.assertTrue(os.path.exists(path))
        self.assertEqual(len(rows), 291)
        self.assertEqual(rows[-2], ['', '', 'Total', '28,852.00'])
        self.assertEqual(rows[180], [u'', u'', u'', u'', 'Total', '1,282.00', '3.50', '57,190.99', u'', 'NB = 0.27%'])
        self.assertEqual(rows[179][9], '1121927')
