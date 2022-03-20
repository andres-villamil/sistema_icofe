from .models import RegistroAdministrativo, AreaTematicaRRAA, TemaRRAA, Entidades_oe, EstadoRegistroAd
import django_filters
from django import forms


class ItemsFilter(django_filters.FilterSet):

    tema = django_filters.ModelChoiceFilter(required=False, label="tema",  queryset=TemaRRAA.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))

    entidad_pri = django_filters.ModelChoiceFilter(required=False, label="entidad",  queryset=Entidades_oe.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))

    area_temat = django_filters.ModelChoiceFilter(required=False, label="área temática",  queryset=AreaTematicaRRAA.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))

    sist_estado = django_filters.ModelChoiceFilter(required=False, label="Fases",  queryset=EstadoRegistroAd.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))

    class Meta:
        model = RegistroAdministrativo
        fields = ['entidad_pri', 'area_temat', 'tema', 'sist_estado']



class NameFilterRRAA(django_filters.FilterSet):

    nombre_ra = django_filters.CharFilter(lookup_expr='icontains',
    widget=forms.TextInput(
        attrs={'placeholder': 'Buscar', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = RegistroAdministrativo
        fields = ['nombre_ra']