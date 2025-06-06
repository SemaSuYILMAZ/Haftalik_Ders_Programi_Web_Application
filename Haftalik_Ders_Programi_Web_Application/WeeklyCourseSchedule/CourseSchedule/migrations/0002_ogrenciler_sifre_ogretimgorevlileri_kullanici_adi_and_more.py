# Generated by Django 5.0.14 on 2025-05-01 15:44

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('CourseSchedule', '0001_initial'),
    ]

    operations = [
        migrations.AddField(
            model_name='ogrenciler',
            name='sifre',
            field=models.CharField(default=1, max_length=11),
            preserve_default=False,
        ),
        migrations.AddField(
            model_name='ogretimgorevlileri',
            name='kullanici_adi',
            field=models.CharField(default=1, max_length=150, unique=True),
            preserve_default=False,
        ),
        migrations.AddField(
            model_name='ogretimgorevlileri',
            name='sifre',
            field=models.CharField(default=1, max_length=11),
            preserve_default=False,
        ),
        migrations.AlterField(
            model_name='dersler',
            name='statu',
            field=models.CharField(choices=[('LAB', 'Laboratuvar'), ('NORMAL', 'Normal')], max_length=10, null=True),
        ),
        migrations.DeleteModel(
            name='OgrenciDers',
        ),
    ]
