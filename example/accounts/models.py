from django.db import models
from jsonfield.fields import JSONField



class  Student(models.Model):
    name = models.CharField(max_length=50)
    age = models.IntegerField()
    courses = models.CharField(choices=[("A", "A"), ("B", "B")])

    #courses = JSONField()