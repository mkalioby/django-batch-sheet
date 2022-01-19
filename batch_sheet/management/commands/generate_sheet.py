import importlib

from django.core.management import BaseCommand


class Command(BaseCommand):
    help = "Generate XLS sheet for a Sheet Class"

    def add_arguments(self, parser):
         parser.add_argument('--xls', nargs='?', type=str,default=None)
         parser.add_argument('--sheet', nargs='?', type=str,default=None)

    def handle(self, *args, **options):
        if options["xls"] is None:
            print("--xls path of output file is required")
            exit(-2)
        if options["sheet"] is None:
            print("Class to export is required --sheet")
            exit(-2)
        module = importlib.import_module(".".join(options["sheet"].split(".")[:-1]))
        s = getattr(module,options["sheet"].split(".")[-1])()
        s.generate_xls(file_path = options["xls"])
        print("Wrote Excel to '%s'"%options["xls"])