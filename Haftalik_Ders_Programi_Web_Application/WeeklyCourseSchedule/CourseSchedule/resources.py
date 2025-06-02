from import_export import resources, fields
from import_export.widgets import Widget
from import_export.widgets import ForeignKeyWidget
from .models import Dersler, Derslikler, Fakulte, Bolumler, OgretimGorevlileri, Ogrenciler, OgrenciDers

class OgretimGorevlileriResource(resources.ModelResource):
    fakulte = fields.Field(
        column_name='Fakülte',
        attribute='fakulte',
        widget=ForeignKeyWidget(Fakulte, 'fakulte_adi')
    )
    ogretim_gorevlisi = fields.Field(
        column_name='Öğretim Görevlisi',
        attribute='ogretim_gorevlisi'
    )
    pazartesi = fields.Field(
        column_name='Pazartesi',
        attribute='pazartesi'
    )
    sali = fields.Field(
        column_name='Salı',
        attribute='sali'
    )
    carsamba = fields.Field(
        column_name='Çarşamba',
        attribute='carsamba'
    )
    persembe = fields.Field(
        column_name='Perşembe',
        attribute='persembe'
    )
    cuma = fields.Field(
        column_name='Cuma',
        attribute='cuma'
    )
    kullanici_adi = fields.Field(
        column_name='Kullanıcı Adı',
        attribute='kullanici_adi'
    )
    sifre = fields.Field(
        column_name='Şifre',
        attribute='sifre'
    )
    class Meta:
        model = OgretimGorevlileri
        import_id_fields = ['ogretim_gorevlisi']
        fields = ('fakulte', 'ogretim_gorevlisi', 'pazartesi', 'sali', 'carsamba', 'persembe', 'cuma', 'kullanici_adi', 'sifre')


class DersliklerResource(resources.ModelResource):
    derslik_id = fields.Field(
        column_name='Derslik_ID',
        attribute='derslik_id'
    )
    kapasite = fields.Field(
        column_name='Kapasite',
        attribute='kapasite'
    )
    statu = fields.Field(
        column_name='Statü',
        attribute='statu'
    )

    class Meta:
        model = Derslikler
        import_id_fields = ['derslik_id']
        fields = ('derslik_id', 'kapasite', 'statu')
        skip_unchanged = True
        report_skipped = True


class DerslerResource(resources.ModelResource):
    fakulte = fields.Field(
        column_name='Fakülte',
        attribute='fakulte',
        widget=ForeignKeyWidget(Fakulte, 'fakulte_adi')
    )
    bolum = fields.Field(
        column_name='Bölüm',
        attribute='bolum',
        widget=ForeignKeyWidget(Bolumler, 'bolum_adi')
    )
    ogretim_uyesi = fields.Field(
        column_name='Öğretim Üyesi',
        attribute='ogretim_uyesi',
        widget=ForeignKeyWidget(OgretimGorevlileri, 'ogretim_gorevlisi')
    )
    ders_kodu = fields.Field(
        column_name='Ders Kodu',
        attribute='ders_kodu'
    )
    ders_adi = fields.Field(
        column_name='Ders Adı',
        attribute='ders_adi'
    )
    sinif = fields.Field(
        column_name='Sınıf',
        attribute='sinif'
    )
    haftalik_saat = fields.Field(
        column_name='Haftalık Saat',
        attribute='haftalik_saat'
    )
    online = fields.Field(
        column_name='Online',
        attribute='online'
    )
    zorunlu_saat = fields.Field(
        column_name='Zorunlu Saat',
        attribute='zorunlu_saat'
    )
    statu = fields.Field(
        column_name='Statü',
        attribute='statu'
    )

    class Meta:
        model = Dersler
        import_id_fields = ['ders_kodu', 'sinif', 'ogretim_uyesi', 'bolum']
        fields = (
            'fakulte', 'bolum', 'sinif', 'ders_kodu', 'ders_adi',
            'ogretim_uyesi', 'haftalik_saat', 'online', 'zorunlu_saat', 'statu'
        )


class OgrencilerResource(resources.ModelResource):
    fakulte = fields.Field(
        column_name='Fakülte',
        attribute='fakulte',
        widget=ForeignKeyWidget(Fakulte, 'fakulte_adi')
    )
    bolum = fields.Field(
        column_name='Bölüm',
        attribute='bolum',
        widget=ForeignKeyWidget(Bolumler, 'bolum_adi')
    )
    sinif = fields.Field(
        column_name='Sınıf',
        attribute='sinif'
    )
    numara = fields.Field(
        column_name='Numara',
        attribute='numara'
    )
    sifre = fields.Field(
        column_name='Şifre',
        attribute='sifre'
    )

    class Meta:
        model = Ogrenciler
        import_id_fields = ['numara']
        fields = ('fakulte', 'bolum', 'sinif', 'numara', 'sifre')


class CustomDersWidgetWithBolum(Widget):
    def clean(self, value, row=None, *args, **kwargs):
        # 'Ders Adı' ve 'Numara' veriler işlenmeli
        ders_adi = str(row.get('Ders Adı', '')).strip()
        numara = str(row.get('Numara', '')).strip()

        from .models import Ogrenciler
        try:
            ogrenci = Ogrenciler.objects.get(numara=numara)
        except Ogrenciler.DoesNotExist:
            raise Exception(f"Öğrenci bulunamadı: {numara}")

        bolum_id = ogrenci.bolum.bolum_id

        # Dersler tablosunda ders adı ve bölüm id'ye göre filtreleme işlemi
        ders_qs = Dersler.objects.filter(ders_adi=ders_adi, bolum_id=bolum_id)

        if ders_qs.exists():
            return ders_qs.first()
        else:
            raise Exception(f"Ders bulunamadı: {ders_adi} / Bölüm ID: {bolum_id}")


class OgrenciDersResource(resources.ModelResource):
    ogrenci = fields.Field(
        column_name='Numara',
        attribute='ogrenci',
        widget=ForeignKeyWidget(Ogrenciler, 'numara')
    )
    ders = fields.Field(
        column_name='Ders Adı',
        attribute='ders',
        widget=CustomDersWidgetWithBolum()
    )

    class Meta:
        model = OgrenciDers
        import_id_fields = ['ogrenci', 'ders']
        fields = ('ogrenci', 'ders')
