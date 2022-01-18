from collections import OrderedDict

import xlrd
import xlsxwriter
from django.conf import settings
from django.db.models import Field
from .Sheet import Sheet

class DeclarativeCombinedSheetsMetaclass(type):
    """
    Metaclass that converts `.Field` objects defined on a class to the
    dictionary `.Sheet.explicit`, taking into account parent class
    `base_fields` as well.
    """

    def __new__(mcs, name, bases, attrs):
        # extract declared columns
        sheets, remainder,explicit = {}, {},{}
        for attr_name, attr in attrs.items():
            if isinstance(attr, Sheet):
                sheets[attr_name]=attr
            elif isinstance(attr, Field):
                attr.name = attr_name
                verbose_name  = attr.verbose_name
                if verbose_name in (None,""):
                    attr.verbose_name = attr_name
                explicit[attr_name]=attr
            else:
                remainder[attr_name] = attr

        attrs = remainder


        # If this class is subclassing other tables, add their fields as
        # well. Note that we loop over the bases in *reverse* - this is
        # necessary to preserve the correct order of columns.
        parent_columns = []
        for base in reversed(bases):
            if hasattr(base, "base_fields"):
                parent_columns = list(base.base_fields.items()) + parent_columns

        # Start with the parent columns
        base_fields = OrderedDict(parent_columns)



        attrs["explicit"]=explicit
        attrs["sheets"] =sheets
        return super().__new__(mcs, name, bases, attrs)

class CombinedSheet(metaclass=DeclarativeCombinedSheetsMetaclass):

    def generate_xls(self):
        workbook = xlsxwriter.Workbook(settings.BASE_DIR + '/data_validate.xls')
        worksheet = workbook.add_worksheet()
        header_format = workbook.add_format({
                'border': 1,
                'bg_color': '#C6EFCE',
                'bold': True,
                'text_wrap': True,
                'valign': 'vcenter',
                'indent': 1,
            })
        worksheet.set_row(0, height=30)
        col_offset = 0

        for sheet_name,sheet in self.sheets.items():
            col_offset = sheet.generate_xls(worksheet,close=False,col_offset=col_offset,header_format=header_format)
        workbook.close()

    def load(self,file_name=None, file_content=None):
        content =[]
        wb = xlrd.open_workbook(file_name)
        sh = wb.sheets()[0]
        for sheet_name, sheet in self.sheets.items():
            c = sheet.convert_json(sh)
            for i,row in enumerate(c):
                if len(content)==i:
                    content.append({})
                content[i].update(row)
        for row in content:
            row_objs = {}
            for sheet_name, sheet in self.sheets.items():
                row_objs[getattr(sheet._meta,"object_name",sheet_name)] = sheet.row_processor(row,row_objs)
        #sh.close()