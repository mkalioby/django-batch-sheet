from django.conf import settings

from accounts.models import Student
from django.core.management.base import BaseCommand
from django.db import models

from batch_sheet.Sheet import Sheet


class StudentSheet(Sheet):

    class Meta:
        Model = Student
        raw_cols = ['gender']
        exclude = ('id',)


class Command(BaseCommand):
    help = "Test management command"

    requires_system_checks = False

    def handle(self, *args, **options):
        # from accounts.sheets import UserSheet
        # u= UserSheet()
        # u.load(file_name="data_validate.xls")
        std = StudentSheet()
        std.generate_xls()
        #std.load(file_name=settings.BASE_DIR + "/data_validate.xls")
