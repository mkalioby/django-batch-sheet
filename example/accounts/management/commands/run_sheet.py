from accounts.models import Student
from django.core.management.base import BaseCommand
from django.db import models

from batch_sheet.Sheet import Sheet


class StudentSheet(Sheet):
    # courses = models.CharField(choices=[("C", "C"), ("D", "D")])

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
        # std.generate_xls()
        std.load(file_name="data_validate.xls")
