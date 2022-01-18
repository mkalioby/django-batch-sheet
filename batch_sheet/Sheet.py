import datetime
from collections import OrderedDict

import django.db.models
import xlrd
import xlsxwriter
from django.conf import settings
from django.core.exceptions import ImproperlyConfigured, ValidationError


class DeclarativeColumnsMetaclass(type):
    """
    Metaclass that converts `.Field` objects defined on a class to the
    dictionary `.Sheet.explicit`, taking into account parent class
    `base_fields` as well.
    """

    def __new__(mcs, name, bases, attrs):
        attrs["_meta"] = opts = SheetOptions(attrs.get("Meta", None), name)

        # extract declared columns
        cols, remainder,explicit = [], {},{}
        for attr_name, attr in attrs.items():
            if isinstance(attr, django.db.models.Field):
                attr.name = attr_name
                verbose_name  = attr.verbose_name
                if verbose_name in (None,""):
                    attr.verbose_name = attr_name
                explicit[attr_name]=attr
            else:
                remainder[attr_name] = attr
        attrs = remainder

        cols.sort(key=lambda x: x[1].creation_counter)

        # If this class is subclassing other tables, add their fields as
        # well. Note that we loop over the bases in *reverse* - this is
        # necessary to preserve the correct order of columns.
        parent_columns = []
        for base in reversed(bases):
            if hasattr(base, "base_fields"):
                parent_columns = list(base.base_fields.items()) + parent_columns

        # Start with the parent columns
        base_fields = OrderedDict(parent_columns)

        # Possibly add some generated columns based on a model
        if opts.model:
            extra = OrderedDict()

            # honor Table.Meta.fields, fallback to model._meta.fields
            if opts.columns is not None:
                # Each item in opts.fields is the name of a model field or a normal attribute on the model
                for field_name in opts.columns:
                    extra[field_name] = field_name
            else:
                for field in opts.model._meta.fields:
                    extra[field.name] = field

            # update base_columns with extra columns
            for key, column in extra.items():
                # skip current column because the parent was explicitly defined,
                # and the current column is not.
                if key in base_fields and base_fields[key]._explicit is True:
                    continue
                base_fields[key] = column

        # Explicit columns override both parent and generated columns
        base_fields.update(OrderedDict(explicit))


        attrs["base_fields"] = base_fields
        attrs["explicit"]=explicit
        return super().__new__(mcs, name, bases, attrs)


class SheetOptions:
    """
    Extracts and exposes options for a `.Sheet` from a `.Sheet.Meta`
    when the sheet is defined.
    Arguments:
        options (`.Sheet.Meta`): options for a sheet from `.Sheet.Meta`
    """

    def __init__(self, options, class_name):
        super().__init__()
        #self._check_types(options, class_name)

        SHEET_ATTRS = getattr(settings, "SHEET_ATTRS", {})

        self.attrs = getattr(options, "attrs", SHEET_ATTRS)
        self.rows_count = getattr(options, "rows_count", 10)
        self.columns = getattr(options, "columns", ())
        self.exclude = getattr(options, "exclude", ())
        self.model = getattr(options, "Model", None)
        self.raw_cols = getattr(options, "raw_cols", [])
        self.validation_exclude = getattr(options, "validation_exclude", [])
        if getattr(options, "object_name", None):
            self.object_name = getattr(options, "object_name", None)



class Sheet(metaclass=DeclarativeColumnsMetaclass):
    """Main Sheet Class"""
    columns = []
    exclude = []
    model = None
    verbose_names = {}
    names = {}
    rows_count = 10
    instances = []
    foreign_keys = {}
    errors = {}
    data = []
    cleaned_data = []

    def __init__(self, *args, **kwargs):
        super().__init__()

        self.model = self._meta.model
        self.exclude = self._meta.exclude
        self.attrs = self._meta.attrs
        self.selected_columns = self._meta.columns
        self.rows_count = self._meta.rows_count
        self.not_provided = []
        if len(self.columns) == 0 and len(self.exclude) == 0:
            raise ImproperlyConfigured(
                "Calling Sheet without defining 'fields' or "
                "'exclude' explicitly is prohibited."
            )

        if len(self.selected_columns) > 0 and len(self.exclude) < 0:
            raise ImproperlyConfigured(
                "Cannot call both columns and exclude. Please specify only one"
            )

        if self.model is None:
            raise ImproperlyConfigured("Model is required for now")
        self._set_active_columns()
        self._get_labels()

    def _set_active_columns(self):
        "Applied columns and exclude settings from the Meta and generate the final list of columns to handle"
        self.columns =[]
        if self.exclude:
            self.columns.extend([f for f in self.model._meta.fields if
                                 f.name not in self.explicit and not f.name in self.exclude])
        else:
            self.columns.extend(
                [f for f in self.model._meta.fields if f.name in self.columns and f.name not in self.explicit])

        for name,field in self.explicit.items():
            self.columns.append(field)


    def _get_labels(self):
        """Generate a dictionary for verbose_names and other one for field names"""
        self.verbose_names={}
        self.names={}
        for field in self.columns:
            self.verbose_names[str(field.verbose_name)] = field
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
        elif field.get_internal_type() == "ForeignKey":
            if field.name not in self._meta.raw_cols:
                l=[]
                if field.null or field.blank:
                    l.append('---')
                l.extend([str(o) for o in field.related_model.objects.all()])
                options = {'validate': 'list', 'source': l}

        return options

    def generate_xls(self,worksheet=None,close=True,col_offset=0,**kwargs):

        if not worksheet:
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
        else:
            header_format = kwargs.get("header_format")

        i = col_offset

        for name, field in self.verbose_names.items():
            if field.get_internal_type() == "NOT_PROVIDED":
                pass
            worksheet.set_column(0, i, width=20)
            print(name, type(name))
            worksheet.write(0, i, name, header_format)
            options = self.sheet_data_validation(field)
            worksheet.data_validation(1, i, self.rows_count, i, options)
            i += 1
        if close:
            workbook.close()
        return col_offset

    def pre_load(self):
        pass


    def is_valid(self):
        if len(self.cleaned_data)==0:
            self.convert_json(self.sheet)
        return len(self.errors) == 0

    def row_preprocessor(self,row):
        return row

    def row_processor(self, row,row_objs={}):
        final_row = {k: v for k, v in row.items() if k in self.names}
        obj = self.model(**final_row)
        x = self.save(obj,row_objs)
        return x

    def process(self,):
        self.convert_json(self.sheet)
        for row in self.cleaned_data:
            x = self.row_processor(row)
            if x:
                self.instances.append(x)

    def post_process(self):
        pass

    def clean(self,row_number, row):
        final_row = {k: v for k, v in row.items() if k in self.names}
        obj = self.model(**final_row)
        try:
            obj.clean_fields(exclude=self._meta.validation_exclude)
        except ValidationError as exp:
            self.errors[row_number] = exp.message_dict

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
        elif field.get_internal_type() == "FloatField":
            if user_val != "":
                return float(user_val)
        elif field.get_internal_type() == "DecimalField":
            from decimal import Decimal
            if user_val != "":
                return Decimal(user_val)
        elif field.get_internal_type() in ["DateField","DateTimeField"]:
            if user_val != "":
                try:
                    return datetime.datetime(*xlrd.xldate_as_tuple(user_val, 0))
                except:
                    raise ValidationError("%s is not a valid %s value"%(user_val,field.get_internal_type()))
        elif field.get_internal_type() == "ForeignKey":
            if not field.name in self.foreign_keys:
                self.foreign_keys[field.name] = {str(o):o for o in field.related_model.objects.all()}
            if user_val != "":
                return self.foreign_keys[field.name].get(user_val,val)
        return user_val if user_val != "" else val

    def convert_json(self, sheet):
        result = []
        self.cleaned_data= []
        self.data= []
        self.errors ={}
        if not result:
            keys = sheet.row_values(0)
            for row in range(1, sheet.nrows):
                temp_result = {}
                row_data = {}
                for col, key in enumerate(keys):
                    user_val = sheet.row_values(row)[col]
                    field = self.verbose_names.get(str(key).strip())
                    if field is None: continue
                    field_name = field.name
                    row_data [field_name]=user_val
                    try:
                        temp_result[field_name] = self.convert_types(field, user_val)
                    except ValidationError as exp:
                        #temp_result[field_name] = None
                        if not row in self.errors: self.errors[row] = {}
                        self.errors[row][field_name] = exp
                for c in self.verbose_names:
                    if not c in temp_result:
                        temp_result[c]=None
                final_row = self.row_preprocessor(temp_result)
                self.cleaned_data.append(final_row)
                self.data.append(row_data)
                self.clean(row,final_row)

    def open(self, file_name=None, file_content=None):
        if file_name:
            wb = xlrd.open_workbook(file_name)
            self.sheet = wb.sheets()[0]

    def load(self, file_name=None, file_content=None):
            self.open(file_name,file_content)
            self.pre_load()
            if self.is_valid():
                self.process()
                self.post_process()
            else:
                raise ValidationError(self.errors)



