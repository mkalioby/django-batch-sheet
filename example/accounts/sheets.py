from django.contrib.auth.models import User

from batch_sheet.Sheet import Sheet






class UserSheet(Sheet):
    def post_process(self):
        print("Called UserSheet post_process")
        super().post_process()
        for obj in self.instances:
            obj.set_password('test1234')
            obj.save()
    class Meta:
        Model = User
        exclude = ('last_login', 'password',"date_joined","id")

def test_sheet():
   u =  UserSheet()
   print (u.verboses_names)
   u.generate_xls()