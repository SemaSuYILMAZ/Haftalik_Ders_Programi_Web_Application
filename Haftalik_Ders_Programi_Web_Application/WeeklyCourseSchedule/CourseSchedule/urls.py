from django.urls import path
from . import views  #views dosyasını projeye ekler
# url'lere karşılık gelen kaynaklara linklendirme işlemleri bu dosya ile gerçekleniyor.
#http://127.0.0.1:8000/

urlpatterns = [
    #path("", views.index),  #url in devamında bir ifade vs yoksa direkt "" ile işlevi yazabiliriz
    path("", views.index, name="index"),
    path("index", views.index),  #her iki işlev de index sayfasını çalıştıracaktır.
    path("ogrenci", views.ogrenci_sayfa, name="ogrenci_sayfa"),
    path('ogretim/', views.ogretim_anasayfa, name='ogretim_sayfa'),
    path('program-olustur/', views.program_olustur_view, name='program_olustur'),
    path("logout/", views.cikis_yap, name="logout"),
    path("ders-secimi-kaydet/", views.ders_secimi_kaydet, name="ders_secimi_kaydet"),
    path("ders-sil/", views.ders_sil, name="ders_sil"),
    path('manuel-atama/', views.manuel_atama, name='manuel_atama'),
]