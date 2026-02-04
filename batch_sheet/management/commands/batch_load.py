import importlib

from django.core.management import BaseCommand


class Command(BaseCommand):
    help = "Generate XLS sheet for a Sheet Class"

    def add_arguments(self, parser):
         parser.add_argument('--xls', nargs='?', type=str,default=None)
         parser.add_argument('--sheet', nargs='?', type=str,default=None)

    def handle(self, *args, **options):
        if options["xls"] is None:
            print("--xls path of input file is required")
        if options["sheet"] is None:
            print("Class to import is required --sheet")

        module = importlib.import_module(".".join(options["sheet"].split(".")[:-1]))
        sheet = getattr(module, options["sheet"].split(".")[-1])
        s = sheet()
        s.open(file_name = options["xls"])
        if s.is_valid():
            s.process()
        else:
            print("File isn't valid")
            print(s.errors)
            exit(-1)