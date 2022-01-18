from django.db import models
from jsonfield.fields import JSONField

class Test(models.Model):
    name = models.CharField(max_length=50)

    def __str__(self):
        return self.name


class Gender (models.Model):
    name = models.CharField(max_length=50)

    def __str__(self):
        return self.name


class Patient(models.Model):
    name = models.CharField(max_length=50,verbose_name="Patient Name")
    MRN = models.CharField(max_length=50)
    gender= models.ForeignKey(Gender,on_delete=models.PROTECT,null=True, verbose_name="Gender")
    date_of_birth = models.DateField(null=True,blank=True,verbose_name="Birth Date")
    date_admitted = models.DateField(auto_now_add=True,auto_created=True)
    lastUpdate = models.DateTimeField(auto_now=True)


class Sample(models.Model):
    patient = models.ForeignKey(Patient,on_delete=models.CASCADE)
    test = models.ForeignKey(Test,on_delete=models.PROTECT,verbose_name="Test")
    sample_date = models.DateField(auto_now_add=True,auto_created=True)
    sample_type = models.CharField(max_length=50, verbose_name="Sample Type",choices=(('Blood','Blood'),('DNA','DNA')))
    lastUpdate = models.DateTimeField(auto_now_add=True)