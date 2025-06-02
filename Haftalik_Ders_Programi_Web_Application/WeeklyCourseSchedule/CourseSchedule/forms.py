from django import forms
from .models import Dersler, Derslikler

class ManuelAtamaForm(forms.Form):
    ders = forms.ModelChoiceField(queryset=Dersler.objects.all(), label="Ders")
    gun = forms.ChoiceField(
        choices=[('Pazartesi', 'Pazartesi'), ('Salı', 'Salı'), ('Çarşamba', 'Çarşamba'),
                 ('Perşembe', 'Perşembe'), ('Cuma', 'Cuma')],
        label="Gün"
    )
    saat = forms.CharField(max_length=20, label="Saat (örn. 10:00-11:00)")
    derslik = forms.ModelChoiceField(queryset=Derslikler.objects.all(), label="Derslik")