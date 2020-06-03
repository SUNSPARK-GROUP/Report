# This is an auto-generated Django model module.
# You'll have to do the following manually to clean this up:
#   * Rearrange models' order
#   * Make sure each model has one field with primary_key=True
#   * Make sure each ForeignKey has `on_delete` set to the desired behavior.
#   * Remove `managed = False` lines if you wish to allow Django to create, modify, and delete the table
# Feel free to rename the models, but don't rename db_table values or field names.
from django.db import models


class ViewMs211Daytotal(models.Model):
    sdate = models.CharField(max_length=12)
    sa_no = models.CharField(max_length=5)
    sname = models.CharField(max_length=100, blank=True, null=True)
    total = models.FloatField(blank=True, null=True)
    postok = models.CharField(max_length=8, blank=True, null=True)
    udate = models.CharField(max_length=14, blank=True, null=True)
    weather = models.CharField(max_length=10, blank=True, null=True)
    daily = models.CharField(max_length=200, blank=True, null=True)
    people = models.IntegerField(blank=True, null=True)
    salecount = models.IntegerField(blank=True, null=True)

    class Meta:
	    db_table = 'View_ms211daytotal'
	    managed = False
		
    '''def _str_(self):
	    return str(self.record_id)'''