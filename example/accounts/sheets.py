from .models import Patient,Test,Sample

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


class SampleSheet(Sheet):
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
        Model = Sample


class NGSSheet(CombinedSheet):
    patient = PatientSheet()
    sample = SampleSheet()
