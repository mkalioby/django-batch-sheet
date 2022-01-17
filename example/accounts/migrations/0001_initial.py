# Generated by Django 2.2.26 on 2022-01-17 08:09

from django.db import migrations, models


class Migration(migrations.Migration):

    initial = True

    dependencies = [
    ]

    operations = [
        migrations.CreateModel(
            name='Student',
            fields=[
                ('id', models.AutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('name', models.CharField(max_length=50)),
                ('age', models.IntegerField()),
                ('courses', models.CharField(choices=[('A', 'A'), ('B', 'B')], max_length=1)),
                ('date_of_birth', models.DateTimeField(blank=True, null=True)),
                ('graduate', models.DateField(blank=True, null=True)),
            ],
        ),
    ]
