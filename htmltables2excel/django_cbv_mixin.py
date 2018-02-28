import datetime
import json
import os
import traceback

from django.conf import settings
import django.core.mail
from django.http import JsonResponse

from convert_tables import parse_tables_from_table_list, PageToExcel


# noinspection PyMethodMayBeStatic
class PageToExcelViewMixin(object):
    """
    A Django view mixin to add page to excel to a page.

    Add a var: excel_base_name or override get_excel_file_name()

    Add this button to you page:
        <button onclick="page_to_excel();" class="btn btn-default btn-xs">Download Report</button>

    Add this to the js section of your page: {% include 'page_to_excel_js_include.html' %}
    """

    def get_excel_file_name(self, request):
        assert hasattr(self, 'excel_base_name'), 'PageToExcelViewMixin Config Error: Add var excel_base_name to class'
        base_name = request.POST.get('base_name', self.excel_base_name)
        csv_name = '%s_%s.xlsx' % (base_name, datetime.datetime.now().strftime('%Y_%m_%d_%H_%M'))
        url = settings.MEDIA_URL + csv_name
        full_path = os.path.join(settings.MEDIA_ROOT, csv_name)
        return full_path, url

    def get_to_excel_params(self):
        """
        Makes the params dict that controls how excel file is created. You probably will want to over-ride this.
        See the class PageToExcel declaration for details.
        """
        kwargs = {}
        return kwargs

    def get_to_excel_excludes(self):
        """
        A list of table IDs to exclude. You probably will want to over-ride this.
        """
        return []

    def post(self, request, *args, **kwargs):
        if request.is_ajax() and 'to_excel' in request.POST:
            # noinspection PyBroadException
            try:
                file_full_path, file_url = self.get_excel_file_name(request)
                to_excel_kwargs = self.get_to_excel_params()
                html_tables = json.loads(request.POST['tables'])
                parsed_tables = parse_tables_from_table_list(html_tables)
                PageToExcel(file_full_path, parsed_tables, **to_excel_kwargs)
                return JsonResponse({'success': True, 'file_url': file_url})
            except:
                txt = '\n'.join(['Got error while rendering page to excel: ' + request.path,
                                 'User: ' + request.user.email,
                                 traceback.format_exc()])
                django.core.mail.mail_admins('HTS: Got error while rendering page to excel', txt)
                return JsonResponse({'success': False})

        else:
            return super(PageToExcelViewMixin, self).post(request, *args, **kwargs)

    def get_context_data(self, **kwargs):
        kwargs['to_excel_excludes'] = self.get_to_excel_excludes()
        kwargs['to_excel'] = True
        kwargs = super(PageToExcelViewMixin, self).get_context_data(**kwargs)
        return kwargs


