from django.conf import settings
from django.core.management import BaseCommand

from accounts.sheets import NGSSheet


class Command(BaseCommand):
    help = "Test management command"

    requires_system_checks = False

    def handle(self, *args, **options):
        # from accounts.sheets import UserSheet
        # u= UserSheet()
        # u.load(file_name="data_validate.xls")
        sheet = NGSSheet()
        #sheet.generate_xls()
        sheet.load(file_name=settings.BASE_DIR + "/data_validate.xls")