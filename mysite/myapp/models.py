from django.db import models
from django.core.validators import RegexValidator

alphanumeric = RegexValidator(r'^[0-9a-zA-Z]*$', 'Only alphanumeric characters are allowed.')

# Create your models here.


class Data(models.Model):
    Date_of_entry = models.CharField(max_length=1000, null=True, blank=True, validators=[alphanumeric])
    Country = models.CharField(max_length=2000, null=True, blank=True, validators=[alphanumeric])
    Vendor = models.CharField(max_length=1000, null=True, blank=True, validators=[alphanumeric])
    Part_number = models.CharField(max_length=1000, null=True, blank=True, validators=[alphanumeric])
    License_name = models.CharField(max_length=1000, null=True, blank=True, validators=[alphanumeric])
    Metric = models.CharField(max_length=1000, null=True, blank=True, validators=[alphanumeric])
    Currency = models.CharField(max_length=500, null=True, blank=True, validators=[alphanumeric])
    Process_number = models.CharField(max_length=1000, null=True, blank=True, validators=[alphanumeric])
    GLobal_price_listprice_USD = models.FloatField(max_length=2000, null=True, validators=[alphanumeric])
    Discount_from_pricelist = models.CharField(max_length=2000, null=True, validators=[alphanumeric])
    Final_price_USD = models.FloatField(max_length=2000, null=True, validators=[alphanumeric])
    Finalprice_of_Localcurrency = models.FloatField(max_length=2000, null=True, blank=True, validators=[alphanumeric])
    Awarded = models.CharField(max_length=100, null=True, blank=True, validators=[alphanumeric])
    Quantity = models.CharField(max_length=500, null=True, blank=True, validators=[alphanumeric])
    Volume_of_Deal = models.CharField(max_length=2000, null=True, blank=True, validators=[alphanumeric])

