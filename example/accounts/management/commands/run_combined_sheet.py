from django.conf import settings
from django.core.management import BaseCommand

from accounts.sheets import NGSSheet


class Command(BaseCommand):
    help = "Test management command"

    requires_system_checks = False

    def handle(self, *args, **options):
        sheet = NGSSheet()
        #sheet.generate_xls()
        sheet.open(file_name=settings.BASE_DIR + "/data_validate.xls")
        if sheet.is_valid():
             sheet.process()
             print(sheet.instances)
        else:
             print(sheet.errors)