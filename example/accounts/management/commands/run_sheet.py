from django.conf import settings

from accounts.models import Student,Courses
from django.core.management.base import BaseCommand
from django.db import models

from batch_sheet.Sheet import Sheet


class StudentSheet(Sheet):
    courses = models.ForeignKey(Courses,on_delete=models.PROTECT)

    def save(self,obj):

        if Student.objects.filter(student_id=obj.student_id).exists():
            s = Student.objects.get(student_id=obj.student_id)
            s.courses.append(obj.courses.id)
        else:
            s = obj
            s.courses=[obj.courses.id]
        s.save()
        return s


    class Meta:
        Model = Student
        exclude = ('id',)


class Command(BaseCommand):
    help = "Test management command"

    requires_system_checks = False

    def handle(self, *args, **options):
        # from accounts.sheets import UserSheet
        # u= UserSheet()
        # u.load(file_name="data_validate.xls")
        std = StudentSheet()
        #std.generate_xls()
        std.load(file_name=settings.BASE_DIR + "/data_validate.xls")
