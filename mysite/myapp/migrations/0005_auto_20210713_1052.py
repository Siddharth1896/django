# Generated by Django 3.2.5 on 2021-07-13 10:52

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('myapp', '0004_alter_data_finalprice_of_localcurrency'),
    ]

    operations = [
        migrations.AlterField(
            model_name='data',
            name='Discount_from_pricelist',
            field=models.CharField(max_length=200, null=True),
        ),
        migrations.AlterField(
            model_name='data',
            name='pub_date',
            field=models.DateTimeField(blank=True, null=True, verbose_name='Date of Entry'),
        ),
    ]
