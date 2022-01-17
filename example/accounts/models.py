from django.db import models
from jsonfield.fields import JSONField

class Courses(models.Model):
    name = models.CharField(max_length=50)

    def __str__(self):
        return self.name

class Gender (models.Model):
    name = models.CharField(max_length=50)

    def __str__(self):
        return self.name

class  Student(models.Model):
    name = models.CharField(max_length=50)
    age = models.IntegerField()
    gender= models.ForeignKey(Gender,on_delete=models.PROTECT,null=True)
    courses = JSONField()
    date_of_birth = models.DateTimeField(null=True,blank=True)
    graduate = models.DateField(null=True,blank=True)
    grade = models.FloatField(null=True, blank=True)

    #courses = JSONField()