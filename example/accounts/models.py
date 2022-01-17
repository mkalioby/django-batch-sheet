from django.db import models
from jsonfield.fields import JSONField



class  Student(models.Model):
    name = models.CharField(max_length=50)
    age = models.IntegerField()
    courses = models.CharField(max_length=1,choices=[("A", "A"), ("B", "B")])
    date_of_birth = models.DateTimeField(null=True,blank=True)
    graduate = models.DateField(null=True,blank=True)
    grade = models.FloatField(null=True, blank=True)

    #courses = JSONField()