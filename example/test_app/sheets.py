from .models import Patient,Test,RequestedTest
from django.db import models

from batch_sheet.Sheet import Sheet
from batch_sheet.CombinedSheet import CombinedSheet


class PatientSheet(Sheet):
    def save(self, obj, row_objs):
        patient = Patient.objects.filter(MRN = obj.MRN)
        if patient.exists():
            p = patient[0]
        else:
            p = obj
            p.save()
        return p
    class Meta:
        exclude=('id','date_admitted','lastUpdate')
        Model = Patient
        obj_name = "patient"


class RequestedTestSheet(Sheet):
    def save(self,obj,row_objs):
        patient = row_objs.get('patient')
        if patient is None:
            return None
        else:
            obj.patient = patient
            obj.save()
            return obj

    class Meta:
        exclude =('id','sample_date','lastUpdate','patient')
        validation_exclude = ('patient',)
        title_header = True
        Model = RequestedTest


class TestSheet(CombinedSheet):
    patient = PatientSheet()
    sample = RequestedTestSheet()
