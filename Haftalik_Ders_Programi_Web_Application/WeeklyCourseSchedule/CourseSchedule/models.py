from django.db import models

class Fakulte(models.Model):
    fakulte_id = models.AutoField(primary_key=True)
    fakulte_adi = models.CharField(max_length=255, unique=True)

    def __str__(self):
        return self.fakulte_adi

    class Meta:
        verbose_name = "Fakülte"
        verbose_name_plural = "Fakülteler"


class Bolumler(models.Model):
    bolum_id = models.AutoField(primary_key=True)
    fakulte = models.ForeignKey(Fakulte, on_delete=models.CASCADE, db_column='fakulte_id')
    bolum_adi = models.CharField(max_length=255, unique=True)

    def __str__(self):
        return self.bolum_adi

    class Meta:
        verbose_name = "Bölüm"
        verbose_name_plural = "Bölümler"


class OgretimGorevlileri(models.Model):
    ogrv_id = models.AutoField(primary_key=True)
    fakulte = models.ForeignKey(Fakulte, on_delete=models.CASCADE, db_column='fakulte_id')
    ogretim_gorevlisi = models.CharField(max_length=255, unique=True)
    pazartesi = models.CharField(max_length=255, null=True, blank=True)
    sali = models.CharField(max_length=255, null=True, blank=True)
    carsamba = models.CharField(max_length=255, null=True, blank=True)
    persembe = models.CharField(max_length=255, null=True, blank=True)
    cuma = models.CharField(max_length=255, null=True, blank=True)
    kullanici_adi = models.CharField(max_length=150, unique=True)
    sifre = models.CharField(max_length=11)

    def __str__(self):
        return self.ogretim_gorevlisi

    class Meta:
        verbose_name = "Öğretim Görevlisi"
        verbose_name_plural = "Öğretim Görevlileri"


class Ogrenciler(models.Model):
    student_id = models.AutoField(primary_key=True)
    fakulte = models.ForeignKey(Fakulte, on_delete=models.CASCADE, db_column='fakulte_id')
    bolum = models.ForeignKey(Bolumler, on_delete=models.CASCADE, db_column='bolum_id', default=1)
    sinif = models.IntegerField()
    numara = models.CharField(max_length=50, unique=True)
    sifre = models.CharField(max_length=11)

    def __str__(self):
        return str(self.numara) if self.numara else f"Öğrenci ID: {self.student_id}"

    class Meta:
        verbose_name = "Öğrenci"
        verbose_name_plural = "Öğrenciler"


class Dersler(models.Model):
    course_id = models.AutoField(primary_key=True)
    fakulte = models.ForeignKey(Fakulte, on_delete=models.SET_NULL, null=True, db_column='fakulte_id')
    bolum = models.ForeignKey(Bolumler, on_delete=models.SET_NULL, null=True, db_column='bolum_id', default=1)
    sinif = models.IntegerField()
    ders_kodu = models.CharField(max_length=50)
    ders_adi = models.CharField(max_length=255)
    ogretim_uyesi = models.ForeignKey(OgretimGorevlileri, on_delete=models.SET_NULL, null=True, db_column='ogrv_id',
                                      default=1)
    haftalik_saat = models.IntegerField()
    online = models.CharField(max_length=255, null=True, blank=True)
    zorunlu_saat = models.CharField(max_length=255, null=True, blank=True)

    STATU_CHOICES = [
        ('LAB', 'Laboratuvar'),
        ('NORMAL', 'Normal'),
    ]
    statu = models.CharField(max_length=10, null=True, choices=STATU_CHOICES)

    def __str__(self):
        return f"{self.ders_kodu} - {self.ders_adi}"

    class Meta:
        verbose_name = "Ders"
        verbose_name_plural = "Dersler"


class Derslikler(models.Model):
    class_id = models.AutoField(primary_key=True)
    derslik_id = models.CharField(max_length=50, unique=True)
    kapasite = models.IntegerField()

    STATU_CHOICES = [
        ('LAB', 'Laboratuvar'),
        ('NORMAL', 'Normal'),
    ]
    statu = models.CharField(max_length=10, choices=STATU_CHOICES)

    def __str__(self):
        return self.derslik_id

    class Meta:
        verbose_name = "Derslik"
        verbose_name_plural = "Derslikler"


class OgrenciDers(models.Model):
    std_course_id = models.AutoField(primary_key=True)
    ogrenci = models.ForeignKey(Ogrenciler, on_delete=models.CASCADE, db_column='student_id')
    ders = models.ForeignKey(Dersler, on_delete=models.CASCADE, db_column='course_id')

    class Meta:
        unique_together = ('ogrenci', 'ders')
        verbose_name = "Öğrenci Ders"
        verbose_name_plural = "Öğrenci Ders"

    def __str__(self):
        return f"{self.ogrenci.numara} - {self.ders.ders_adi}"

