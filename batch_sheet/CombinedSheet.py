from collections import OrderedDict

import xlrd
import xlsxwriter
from django.conf import settings
from django.core.exceptions import ValidationError
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

    def __init__(self,*args,**kwargs):
        super(CombinedSheet, self).__init__()
        self.data = []
        self.cleaned_data = []
        self.errors = {}
        self.instances = []
        self._valid = None
        self.data_sheet_name = 'BatchSheetDetails'

    def generate_xls(self,file_path=settings.BASE_DIR + '/data_validate.xls'):
        workbook = xlsxwriter.Workbook(file_path)
        worksheet = workbook.add_worksheet()
        data_worksheet = workbook.add_worksheet(name=self.data_sheet_name)
        header_format = workbook.add_format({
                'border': 1,
                'bold': True,
                'text_wrap': True,
                'valign': 'vcenter',
                'indent': 1,
            })
        required_format = workbook.add_format({
            'border': 1,
            'color': '#ff0000',
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'indent': 1,
        })
        worksheet.set_row(0, height=30)
        col_offset = 0

        for sheet_name,sheet in self.sheets.items():
            col_offset = sheet.generate_xls(worksheet=worksheet,data_worksheet=data_worksheet,close=False,col_offset=col_offset,header_format=header_format,required_format=required_format)
        workbook.close()

    def open(self, file_name=None, file_content=None):
        content = []
        wb = xlrd.open_workbook(file_name)
        self.xls_sheet = wb.sheets()[0]

    def is_valid(self):
        if self._valid is None:
            for sheet_name, sheet in self.sheets.items():
                sheet.sheet =self.xls_sheet
                if not sheet.is_valid():
                    for k in sheet.errors:
                        if not k in self.errors:
                            self.errors[k]={}
                        self.errors[k].update(sheet.errors[k])
                for i,d in enumerate(sheet.data):
                    if len(self.data) == i:
                        self.data.append({})
                    self.data[i].update(d)
                for i, d in enumerate(sheet.cleaned_data):
                    if len(self.cleaned_data) == i:
                        self.cleaned_data.append({})
                    self.cleaned_data[i].update(d)

        self._valid = len(self.errors) == 0
        return self._valid

    def process(self):
        if self._valid:
            for row in self.cleaned_data:
                row_objs = {}
                for sheet_name, sheet in self.sheets.items():
                    obj_name = getattr(sheet._meta, "object_name", sheet_name)
                    row_objs[obj_name] = sheet.row_processor(row, row_objs)
                self.instances.append(row_objs)


    def load(self,file_name=None, file_content=None):
        self.open(file_name,file_content)
        if self.is_valid():
            self.process()
        else:
            raise ValidationError("Sheet is not valid")
        #sh.close()