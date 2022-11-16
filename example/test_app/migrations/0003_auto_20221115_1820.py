# Generated by Django 4.1.3 on 2022-11-15 18:20

from django.db import migrations

def add_dummy_tests(apps, schema_editor):
    Test = apps.get_model("test_app", "Test")
    for i in range(4,300):
        test_obj = Test(name=f'Test {i}')
        test_obj.save()



class Migration(migrations.Migration):

    dependencies = [
        ('test_app', '0002_inital_data'),
    ]

    operations = [
        migrations.RunPython(add_dummy_tests, migrations.RunPython.noop)
    ]