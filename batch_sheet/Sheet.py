import datetime

import xlrd
import xlsxwriter
from django.conf import settings
from django.core.exceptions import ImproperlyConfigured


class Sheet:
    columns = None
    exclude = None
    model = None
    verboses_names = {}
    names = {}
    rows_count = 10
    instances = []

    def __init__(self, *args, **kwargs):
        cls = self.Meta
        self.model = cls.Model
        if "columns" not in cls.__dict__ and "exclude" not in cls.__dict__:
            raise ImproperlyConfigured(
                "Calling Sheet without defining 'fields' or "
                "'exclude' explicitly is prohibited."
            )

        if "columns" in cls.__dict__ and "exclude" in cls.__dict__:
            raise ImproperlyConfigured(
                "Cannot call both columns and exclude. Please specify only one"
            )

        if "exclude" in cls.__dict__: self.exclude = cls.exclude
        if "columns" in cls.__dict__: self.columns = cls.columns
        if "rows_count" in cls.__dict__: self.rows_count = cls.rows_count

        self._set_active_columns()
        self._get_labels()

    def _set_active_columns(self):
        self.columns = []
        # for item in vars(self):
        #     self.columns.append(item)
        if self.exclude:
            self.columns.extend([f for f in self.model._meta.fields if
                                 f.name not in self.exclude and getattr(self, f.name, None) is None])
        else:
            self.columns.extend(
                [f for f in self.model._meta.fields if f.name in self.columns and getattr(self, f.name, None) is None])
        print(44, self.columns)

    def _get_labels(self):
        for field in self.columns:
            self.verboses_names[str(field.verbose_name)] = field
            self.names[str(field.name)] = field

    def sheet_data_validation(self, field):
        options = {'validate': 'any', 'criteria': 'not equal to', 'value': None}
        if field.get_internal_type() == "BooleanField":
            options = {'validate': 'list', 'source': ["", "Yes", "No"]}
        elif field.get_internal_type() == "CharField":
            if field.choices:
                options = {'validate': 'list', 'source': [c[0] for c in field.choices]}
        elif field.get_internal_type() == "DateField":
            options = {'validate': 'date',
                       'criteria': '>',
                       'value': datetime.date(1900, 1, 1),
                       'input_title': 'Date format: YYYY-MM-DD',
                       'input_message': 'Greater than 1900-01-01',
                       'error_title': 'Date is not valid!',
                       'error_message': 'It should be greater than 1900-01-01',
                       'error_type': 'information'
                       }
        elif field.get_internal_type() == "DateTimeField":
            options = {'validate': 'date',
                       'criteria': '>',
                       'value': datetime.date(1900, 1, 1),
                       'input_title': 'Date format: YYYY-MM-DD HH:MM',
                       'input_message': 'Greater than 1900-01-01 00:00:00',
                       'error_title': 'Date is not valid!',
                       'error_message': 'It should be greater than 1900-01-01 00:00:00',
                       'error_type': 'information'
                       }
        return options

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
        i = 0
        for name, field in self.verboses_names.items():
            worksheet.set_column(0, i, width=20)
            worksheet.write(0, i, name, header_format)
            options = self.sheet_data_validation(field)
            worksheet.data_validation(1, i, self.rows_count, i, options)
            i += 1
        workbook.close()

    def pre_load(self):
        pass

    def valid(self):
        return True

    def process(self):
        for row in self.content:
            final_row = {k: v for k, v in row.items() if k in self.names}
            obj = self.model(**final_row)
            obj.save()
            self.instances.append(obj)

    def post_process(self):
        print("Called Class post_process")

    def row_processor(self, row):
        return row

    def convert_types(self, field, user_val):
        val = field.get_default()
        null = field.null
        if field.get_internal_type() == "BooleanField":
            if user_val == "Yes": val = True
            if user_val == "No": val = False
            return val
        elif field.get_internal_type() == "IntegerField":
            if user_val != "":
                return int(user_val)
        elif field.get_internal_type() == "DateField":
            if user_val != "":
                return datetime.datetime(*xlrd.xldate_as_tuple(user_val, 0))
        elif field.get_internal_type() == "DateTimeField":
            if user_val != "":
                return datetime.datetime(*xlrd.xldate_as_tuple(user_val, 0))
        else:
            return user_val if user_val != "" else val

    def convert_json(self, sheet):
        result = []
        if not result:
            keys = sheet.row_values(0)
            for row in range(1, sheet.nrows):
                temp_result = {}
                for col, key in enumerate(keys):
                    user_val = sheet.row_values(row)[col]
                    field = self.verboses_names[str(key).strip()]
                    temp_result[field.name] = self.convert_types(field, user_val)
                result.append(self.row_processor(temp_result))
        self.content = result

    def load(self, file_name=None, file_content=None):
        if file_name:
            wb = xlrd.open_workbook(file_name)
            sh = wb.sheets()[0]
            self.pre_load()
            self.convert_json(sh)
            print(self.content)
            res = self.valid()
            if res:
                self.process()
                self.post_process()

# def dump_old():
#     workbook = xlsxwriter.Workbook('data_validate.xlsx')
#     worksheet = workbook.add_worksheet()
#
#     header_format = workbook.add_format({
#         'border': 1,
#         'bg_color': '#C6EFCE',
#         'bold': True,
#         'text_wrap': True,
#         'valign': 'vcenter',
#         'indent': 1,
#     })
#
#     heading1 = 'Some examples of data validation in XlsxWriter'
#     heading2 = 'Enter values in this column'
#
#     worksheet.write('A1', heading1, header_format)
#     worksheet.write('B1', heading2, header_format)
#
#     # Example 1. Limiting input to an integer in a fixed range.
#     #
#     txt = 'Enter an integer between 1 and 10'
#
#     worksheet.write('A3', txt)
#     worksheet.data_validation('B3', {'validate': 'integer',
#                                      'criteria': 'between',
#                                      'minimum': 1,
#                                      'maximum': 10})
#
#     # Example 2. Limiting input to an integer outside a fixed range.
#     #
#     txt = 'Enter an integer that is not between 1 and 10 (using cell references)'
#
#     worksheet.write('A5', txt)
#     worksheet.data_validation('B5', {'validate': 'integer',
#                                      'criteria': 'not between',
#                                      'minimum': '=E3',
#                                      'maximum': '=F3'})
#
#     # Example 3. Limiting input to an integer greater than a fixed value.
#     #
#     txt = 'Enter an integer greater than 0'
#
#     worksheet.write('A7', txt)
#     worksheet.data_validation('B7', {'validate': 'integer',
#                                      'criteria': '>',
#                                      'value': 0})
#
#     # Example 4. Limiting input to an integer less than a fixed value.
#     #
#     txt = 'Enter an integer less than 10'
#
#     worksheet.write('A9', txt)
#     worksheet.data_validation('B9', {'validate': 'integer',
#                                      'criteria': '<',
#                                      'value': 10})
#
#     # Example 5. Limiting input to a decimal in a fixed range.
#     #
#     txt = 'Enter a decimal between 0.1 and 0.5'
#
#     worksheet.write('A11', txt)
#     worksheet.data_validation('B11', {'validate': 'decimal',
#                                       'criteria': 'between',
#                                       'minimum': 0.1,
#                                       'maximum': 0.5})
#
#     # Example 6. Limiting input to a value in a dropdown list.
#     #
#     txt = 'Select a value from a drop down list'
#
#     worksheet.write('A13', txt)
#     worksheet.data_validation('B13', {'validate': 'list',
#                                       'source': ['open', 'high', 'close']})
#
#     # Example 7. Limiting input to a value in a dropdown list.
#     #
#     txt = 'Select a value from a drop down list (using a cell range)'
#
#     worksheet.write('A15', txt)
#     worksheet.data_validation('B15', {'validate': 'list',
#                                       'source': '=$E$4:$G$4'})
#
#     # Example 8. Limiting input to a date in a fixed range.
#     #
#     txt = 'Enter a date between 1/1/2013 and 12/12/2013'
#
#     worksheet.write('A17', txt)
#     worksheet.data_validation('B17', {'validate': 'date',
#                                       'criteria': 'between',
#                                       'minimum': date(2013, 1, 1),
#                                       'maximum': date(2013, 12, 12)})
#
#     # Example 9. Limiting input to a time in a fixed range.
#     #
#     txt = 'Enter a time between 6:00 and 12:00'
#
#     worksheet.write('A19', txt)
#     worksheet.data_validation('B19', {'validate': 'time',
#                                       'criteria': 'between',
#                                       'minimum': time(6, 0),
#                                       'maximum': time(12, 0)})
#
#     # Example 10. Limiting input to a string greater than a fixed length.
#     #
#     txt = 'Enter a string longer than 3 characters'
#
#     worksheet.write('A21', txt)
#     worksheet.data_validation('B21', {'validate': 'length',
#                                       'criteria': '>',
#                                       'value': 3})
#
#     # Example 11. Limiting input based on a formula.
#     #
#     txt = 'Enter a value if the following is true "=AND(F5=50,G5=60)"'
#
#     worksheet.write('A23', txt)
#     worksheet.data_validation('B23', {'validate': 'custom',
#                                       'value': '=AND(F5=50,G5=60)'})
#
#     # Example 12. Displaying and modifying data validation messages.
#     #
#     txt = 'Displays a message when you select the cell'
#
#     worksheet.write('A25', txt)
#     worksheet.data_validation('B25', {'validate': 'integer',
#                                       'criteria': 'between',
#                                       'minimum': 1,
#                                       'maximum': 100,
#                                       'input_title': 'Enter an integer:',
#                                       'input_message': 'between 1 and 100'})
#
#     # Example 13. Displaying and modifying data validation messages.
#     #
#     txt = "Display a custom error message when integer isn't between 1 and 100"
#
#     worksheet.write('A27', txt)
#     worksheet.data_validation('B27', {'validate': 'integer',
#                                       'criteria': 'between',
#                                       'minimum': 1,
#                                       'maximum': 100,
#                                       'input_title': 'Enter an integer:',
#                                       'input_message': 'between 1 and 100',
#                                       'error_title': 'Input value is not valid!',
#                                       'error_message':
#                                           'It should be an integer between 1 and 100'})
#
#     # Example 14. Displaying and modifying data validation messages.
#     #
#     txt = "Display a custom info message when integer isn't between 1 and 100"
#
#     worksheet.write('A29', txt)
#     worksheet.data_validation('B29', {'validate': 'integer',
#                                       'criteria': 'between',
#                                       'minimum': 1,
#                                       'maximum': 100,
#                                       'input_title': 'Enter an integer:',
#                                       'input_message': 'between 1 and 100',
#                                       'error_title': 'Input value is not valid!',
#                                       'error_message': 'It should be an integer between 1 and 100',
#                                       'error_type': 'information'})
#
#     workbook.close()
