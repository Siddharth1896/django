from django.contrib import admin
from import_export.admin import ImportExportActionModelAdmin

# Register your models here.

from .models import Data
#admin.site.register(Data)

@admin.register(Data)
class ViewAdmin(ImportExportActionModelAdmin):
    list_display = ('License_name','Part_number','Date_of_entry','Country','Vendor','Metric','Currency','Process_number','GLobal_price_listprice_USD','Discount_from_pricelist','Final_price_USD','Finalprice_of_Localcurrency','Awarded','Quantity','Volume_of_Deal')

