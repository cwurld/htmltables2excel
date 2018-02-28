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
from xlsxwriter.utility import xl_rowcol_to_cell, xl_col_to_name

from page_to_csv import parse_tables


def style_to_dict(style):
    """Parses an HTML tag style attribute.
    :param style:
    """
    if isinstance(style, dict):
        return style

    d = {}
    styles = style.split(';')
    for s in styles:
        # noinspection PyBroadException
        try:
            key, value = s.split(':')
            d[key.strip()] = value.strip()
        except:
            pass
    return d


def clean_cell(cell):
    if cell.string:
        s = six.text_type(cell.string).strip()
        if s is None or s == u'\n' or len(s) == 0:
            return u''
        else:
            return s
    else:
        s = u''.join([x for x in cell.stripped_strings])
        return s


def locate_cells(formula, current_row, current_col):
    """
    Converts all cell location codes to an excel locations relative to the current cell. Cell location codes are
    in the following format.

    Cols: colmddd - the string col then either an m or a p then a 3 digit offset. m means the offset should be
        subtracted from the current col. p means the offset should be added to the current col

    Rows: rowpddd - similar to cols

    :param formula: the formula string
    :param current_row: zero based value of the current row
    :param current_col: zero based value of the current col
    :return: the formula with the row and column references in excel format (eg. B10).
    """

    def locate_col(match):
        sign = match.group('sign')
        offset = match.group('offset')
        try:
            offset = int(offset)
        except ValueError:
            return ''

        if sign == 'm':
            return xl_col_to_name(current_col - offset)
        elif sign == 'p':
            return xl_col_to_name(current_col + offset)
        return ''

    def locate_row(match):
        sign = match.group('sign')
        offset = match.group('offset')
        try:
            offset = int(offset)
        except ValueError:
            return ''

        if sign == 'm':
            return str(current_row - offset + 1)
        elif sign == 'p':
            return str(current_row + offset + 1)
        return ''

    col_re = re.compile('col(?P<sign>[m,p])(?P<offset>\d\d\d)')
    new_formula = col_re.sub(locate_col, formula)
    row_re = re.compile('row(?P<sign>[m,p])(?P<offset>\d\d\d)')
    new_formula = row_re.sub(locate_row, new_formula)
    return new_formula


# noinspection PyUnusedLocal
def make_formula(formula_str, row, col, first_data_row=None):
    # noinspection SpellCheckingInspection
    """
        A cell will be written as a formula if the HTML tag has the attribute "data-excel" set.

        Note that this function is called when the spreadsheet is being created. The cell it applies to knows where it
        is and what the first data row is.

        Allowed formula strings:

            "SUM ROW A-C": sum the current row from A-C

            "SUM ROW A,C": sum cells A and C in the current row

            "SUM COL": sums current col from first_row to row - 1

            "FORMULA RAW IF(F13 > 0, (F13-E13)/F13, '')": uses formula as is

            "FORMULA RELATIVE IF(colm001rowp000 > 0, (colm001rowp0-colm002rowp000)/colm001rowp001, '')": creates the
                formula relative to the current location. colm002 means two cols to the left of the current cell.
                rowp000 means the current row plus 0 (e.g. the current row)

        :param formula_str: the value of the "data-excel" tag containing params for generating the formula
        :param row: cell row
        :param col: cell column
        :param first_data_row: for column formulas
        :return: a string
        """
    parts = formula_str.split(' ')
    func = parts[0]
    args = parts[-1]

    formula = ''
    if func == 'SUM':
        func_modifier = parts[1]
        if func_modifier == 'ROW':
            if '-' in args:
                cols = args.split('-')
                formula = '=SUM({}{}:{}{})'.format(cols[0], row + 1, cols[1], row + 1)
            elif ',' in args:
                cols = map(methodcaller('strip'), args.split(','))

                # Put the row number after each col letter and then add them together
                cols = '+'.join(map(lambda x: x + str(row + 1), cols))
                formula = '=SUM({})'.format(cols)
        elif func_modifier == 'COL':
            formula = '=SUM({}:{})'.format(xl_rowcol_to_cell(first_data_row, col), xl_rowcol_to_cell(row - 1, col))
    elif func == 'FORMULA':
        func_modifier = parts[1]
        formula_str = ' '.join(parts[2:])
        if func_modifier == 'RAW':
            formula = '=' + formula_str
        elif func_modifier == 'RELATIVE':
            formula = '=' + locate_cells(formula_str, row, col)
    return formula


def configure_worksheet(worksheet, table, first_data_row):
    """
    Currently just for freezing. Set the attribute "data-excel" of HTML table tag to:

        FREEZE <row>,<col>

    :param first_data_row:
    :param worksheet:
    :param table: a beautiful soup parsed table tag
    :return:
    """
    data_excel = table.attrs.get('data-excel')
    if data_excel:
        func, args = data_excel.split(' ')
        if func == 'FREEZE':
            args = map(int, args.split(','))
            worksheet.freeze_panes(*args)
    else:
        # Freeze headers by default
        worksheet.freeze_panes(first_data_row, 0)


def parse_row(row):
    """

    :param row: a beautiful soup table row.
    :return: a list of parsed cells, each item is a dict:
            {'value': value, 'attrs': cell.attrs, 'tag': cell.name, 'is_money': False}
    """
    result = []
    for cell in row.children:
        if not isinstance(cell, NavigableString):
            if 'style' in cell.attrs:
                cell.attrs['style'] = style_to_dict(cell.attrs['style'])

            if 'class' in cell.attrs:
                cell.attrs['class'] = filter(lambda x: x != '', cell.attrs['class'])

            s = six.text_type(clean_cell(cell))
            if s and s[0] == u'$':
                value = s[1:].replace(u',', u'')
                is_money = True
                try:
                    value = float(value)
                except ValueError:
                    is_money = False
                contents = {'value': value,
                            'attrs': cell.attrs, 'tag': cell.name, 'is_money': is_money}

            elif s and s[-1] == u'%':
                try:
                    number = float(s[0: -1]) / 100.0
                except ValueError:
                    number = None

                if number:
                    contents = {'value': number, 'attrs': cell.attrs, 'tag': cell.name, 'is_percent': True}
                else:
                    contents = contents = {'value': s, 'attrs': cell.attrs, 'tag': cell.name, 'is_money': False}
            else:
                if s.isnumeric():
                    value = int(s)
                else:
                    # http://stackoverflow.com/questions/736043/checking-if-a-string-can-be-converted-to-float-in-python
                    try:
                        value = float(s.replace(',', ''))
                    except ValueError:
                        value = s
                contents = {'value': value, 'attrs': cell.attrs, 'tag': cell.name, 'is_money': False}
            result.append(contents)
    return result


def parse_table(table):
    if table.name == u'[document]':
        table = table.table

    data = {'table': table}  # we may need some attributes or classes
    if table.caption:
        title = table.caption.string
        data['caption'] = six.text_type(title)

    data['headers'] = []
    if table.thead:
        for row in table.thead.find_all('tr'):
            data['headers'].append(parse_row(row))

    data['rows'] = []
    for row in table.tbody.find_all('tr'):
        data['rows'].append(parse_row(row))

    data['footers'] = []
    if table.tfoot:
        for row in table.tfoot.find_all('tr'):
            data['footers'].append(parse_row(row))

    return data


def parse_tables_from_table_list(table_list):
    """
    :param table_list: a list of table html
    :return:
    """
    parsed_tables = []
    for table in table_list:
        soup = BeautifulSoup(table, 'html.parser')
        parsed_tables.append(parse_table(soup))
    return parsed_tables


def full_page_to_excel(file_full_path, html, **kwargs):
    """Converts a full HTML page to excel.
    :param kwargs:
    :param html:
    :param file_full_path:
    """
    excluded_tables = kwargs.pop('excluded_tables', [])
    tables = parse_tables(html, excluded_tables, parse_table)
    PageToExcel(file_full_path, tables, **kwargs)


class PageToExcel(object):
    def __init__(self, file_full_path, tables, work_sheet_names=None, extra_headers=None, col_widths=None,
                 custom_formats=None, show_table_captions=None, external_workbook=None, include_formulas=True):
        """
        Writes tables to excel. NOTE: there can be more than one table. Each table is a separate worksheet.

        You can write a cell as a formula by setting the HTML tag attribute "data-excel". For details see the
        function make_formula().

        To apply one of the formats in self.format, put the name of the format in the cell CSS class.

        :param file_full_path:
        :param tables: a list of html for each table.
        :param work_sheet_names: either None or a list with length = number of tables
        :param extra_headers: None or headers for each worksheet. E.g.
            [[worksheet 1 header1, worksheet 1 header2, ...], [worksheet 2 header1, worksheet 2 header2, ...]].
            Each header is a merged row, the width of the worksheet.
        :param col_widths: None or columns widths for each page. E.g.
            [[('B:B', 50), ('D:F', 20)], [('A:B', 25), ('D:F', 20)]]
        :param custom_formats: None or a dict of format:
            {'class name 1': {format params}, 'class name 2': {format params}}.
            Put the class in the HTML tag.
        :param show_table_captions: a list of booleans, one for each table. If none, then all captions are shown
        :param external_workbook: for adding to an existing workbook
        :param include_formulas: when false, cell values are not replaced by formulas.
        :return: None
        """
        self.file_full_path = file_full_path
        workbook = external_workbook or xlsxwriter.Workbook(file_full_path)
        self.workbook = workbook
        self.include_formulas = include_formulas

        self.formats = {
            'money': workbook.add_format({'num_format': '$#,##0.00', 'align': 'right'}),
            'dollars': workbook.add_format({'num_format': '$#,##0', 'align': 'right'}),
            'hours': workbook.add_format({'num_format': '#,##0.0', 'align': 'right'}),
            'percent': workbook.add_format({'num_format': '0.00%', 'align': 'right'}),
            'integer': workbook.add_format({'num_format': '#,##0', 'align': 'right'}),

            'header': workbook.add_format(
                {'bold': True, 'bg_color': '#CCCCCC', 'bottom_color': 'black', 'bottom': 1}),

            'centered_header': workbook.add_format(
                {'bold': True, 'bg_color': '#CCCCCC', 'bottom_color': 'black', 'bottom': 1, 'align': 'center_across'}),

            'right_header': workbook.add_format(
                {'bold': True, 'bg_color': '#CCCCCC', 'bottom_color': 'black', 'bottom': 1, 'align': 'right'}),

            'upper_header': workbook.add_format({'bold': True, 'bg_color': '#CCCCCC'}),

            'bold': workbook.add_format({'bold': True}),
            'underline': workbook.add_format({'underline': 1}),
            'title': workbook.add_format({'bold': True, 'font_size': 13}),
            'url': workbook.add_format({'font_color': 'blue', 'underline': 1}),
            'right_align': workbook.add_format({'align': 'right'}),
            'row_date': workbook.add_format({'num_format': 'D-MMM'}),

            # HTML-like formatting
            'th': workbook.add_format({'bold': True}),
            'td': None
        }

        if custom_formats:
            for format_name, format_dict in six.iteritems(custom_formats):
                self.formats[format_name] = workbook.add_format(format_dict)

        cw = []
        eh = []
        for i, table in enumerate(tables):
            if col_widths:
                cw = col_widths[i]

            if extra_headers:
                eh = extra_headers[i]

            if work_sheet_names:
                name = work_sheet_names[i]
            else:
                name = 'sheet_{}'.format(i + 1)

            if show_table_captions:
                show_table_caption = show_table_captions[i]
            else:
                show_table_caption = True

            self.write_page(name, table, eh, cw, show_table_caption)

        if not external_workbook:
            self.workbook.close()

    def get_fmt(self, cell, default=None):
        if 'class' in cell['attrs']:
            for the_class in cell['attrs']['class']:
                if the_class in self.formats:
                    return self.formats[the_class]

        if cell.get('is_money'):
            return self.formats['money']
        elif cell.get('is_percent'):
            return self.formats['percent']
        else:
            return default

    def write_cell(self, worksheet, row, col, cell, cell_format=None, first_data_row=None):
        colspan = int(cell['attrs'].get('colspan', u'1'))

        # Write formula if there is one
        formula_str = cell['attrs'].get('data-excel')
        if formula_str and self.include_formulas:
            value = make_formula(formula_str, row, col, first_data_row=first_data_row)
        else:
            value = cell['value']

        if colspan > 1:
            the_format = self.get_fmt(cell, default=cell_format)
            worksheet.merge_range(row, col, row, col + colspan - 1, value, the_format)
            next_col = col + colspan
        else:
            worksheet.write(row, col, value, self.get_fmt(cell, default=cell_format))
            next_col = col + 1
        return next_col

    def write_page(self, name, data, extra_headers, col_widths, show_table_caption):
        """
        :param name:
        :param extra_headers:
        :param show_table_caption:
        :param data:
        :param col_widths: a list, each element ['B:F', 12]. Widths are in chars of default font size. Set to 0 to hide.
        :return:
        """
        worksheet = self.workbook.add_worksheet(name)

        # Count columns
        n_cols = 2
        for row in data['headers']:
            n_cols = max(n_cols, len(row))

        # Set column widths
        if col_widths:
            for cols, width in col_widths:
                if width == 0:
                    worksheet.set_column(cols, None, None, {'hidden': 1})
                else:
                    worksheet.set_column(cols, width)

        # Write page headers -------------------------------------------------------------------------------------------
        headers = []
        if extra_headers:
            headers += extra_headers

        if data.get('caption') and show_table_caption:
            headers.append(data['caption'])

        row = 0
        if headers:
            for i, h in enumerate(headers):
                if i == 0:
                    cell_format = self.formats['title']
                else:
                    cell_format = self.formats['bold']
                worksheet.merge_range(row, 0, row, n_cols - 1, h, cell_format)
                row += 1
            worksheet.write(row, 0, '')
            row += 1

        # Write data headers ---------------------------------------------------------------------------------------
        for table_row in data['headers']:
            col = 0
            for cell in table_row:
                col = self.write_cell(worksheet, row, col, cell, self.formats['header'])
            row += 1

        # Freeze
        configure_worksheet(worksheet, data['table'], row)

        first_data_row = row
        # Write data ---------------------------------------------------------------------------------------
        for table_row in data['rows']:
            col = 0
            for cell in table_row:
                col = self.write_cell(worksheet, row, col, cell, first_data_row=first_data_row)
            row += 1

        # Write data footers ---------------------------------------------------------------------------------------
        for table_row in data['footers']:
            col = 0
            for cell in table_row:
                col = self.write_cell(worksheet, row, col, cell, first_data_row=first_data_row)
            row += 1


# ------------------------------------------------------------------------------------------------------------------
class TestPageExcel(unittest.TestCase):
    def test_make_formula(self):
        formula = make_formula('SUM ROW A-C', 1, 2)
        self.assertEqual(formula, '=SUM(A2:C2)')

        formula = make_formula('SUM ROW A,C', 1, 2)
        self.assertEqual(formula, '=SUM(A2+C2)')

        formula = make_formula('SUM COL', 3, 2, first_data_row=1)
        self.assertEqual(formula, '=SUM(C2:C3)')

    def test_to_excel(self):
        fp = open(os.path.join(settings.SITE_PATH, 'utils/table_to_csv_test_data.html'), 'rb')
        self.html = fp.read()
        fp.close()

        full_page_to_excel(
            'test_page_to_excel.xlsx',
            self.html,
            work_sheet_names=['Labor', 'Revenue'],
            extra_headers=[['Header 1', 'Subheader 1'], ['Header 2']],
            col_widths=[[('A:A', 20)], [['B:B', 30]]]
        )

    def test_formulas(self):
        fp = open(os.path.join(settings.SITE_PATH, 'utils/simple_table_for_testing.html'), 'rb')
        self.html = fp.read()
        fp.close()

        full_page_to_excel(
            'test_page_to_excel.xlsx',
            self.html,
            work_sheet_names=['S1'],
        )

    def test_locate_cell(self):
        current_row = 13
        current_column = 5
        # noinspection SpellCheckingInspection
        x = locate_cells('colm001rowp000 + colm002rowp001', current_row, current_column)
        self.assertEqual(x, 'E14 + D15')
