from django.contrib import admin
from import_export.admin import ImportExportModelAdmin
from .models import Fakulte, Bolumler, OgretimGorevlileri, Ogrenciler, Dersler, Derslikler, OgrenciDers

admin.site.site_header = "Ders Programı Yönetim Paneli"
admin.site.site_title = "Ders Programı Admin"
admin.site.index_title = "Yönetim Paneline Hoş Geldiniz"

admin.site.register(Fakulte)
admin.site.register(Bolumler)

from .resources import DerslerResource
@admin.register(Dersler)
class DerslerAdmin(ImportExportModelAdmin):
    resource_class = DerslerResource

from .resources import DersliklerResource
@admin.register(Derslikler)
class DersliklerAdmin(ImportExportModelAdmin):
    resource_class = DersliklerResource

from .resources import OgrencilerResource
@admin.register(Ogrenciler)
class OgrencilerAdmin(ImportExportModelAdmin):
    resource_class = OgrencilerResource
    list_display = ('numara', 'sinif', 'sifre')

from .resources import OgretimGorevlileriResource
@admin.register(OgretimGorevlileri)
class OgretimGorevlileriAdmin(ImportExportModelAdmin):
    resource_class = OgretimGorevlileriResource
    list_display = ('ogretim_gorevlisi', 'kullanici_adi', 'sifre')

from .resources import OgrenciDersResource
@admin.register(OgrenciDers)
class OgrenciDersAdmin(ImportExportModelAdmin):
    resource_class = OgrenciDersResource
