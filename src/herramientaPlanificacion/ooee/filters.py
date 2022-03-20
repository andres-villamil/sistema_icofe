from .models import OperacionEstadistica, Tema, AreaTematica, Entidades_oe, OoeeState
import django_filters
from django import forms


class TopicsFilter(django_filters.FilterSet):

    tema = django_filters.ModelChoiceFilter(required=False, label="tema",  queryset=Tema.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))

    entidad = django_filters.ModelChoiceFilter(required=False, label="entidad",  queryset=Entidades_oe.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))

    area_tematica = django_filters.ModelChoiceFilter(required=False, label="área temática",  queryset=AreaTematica.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))

    nombre_est = django_filters.ModelChoiceFilter(required=False, label="Fases",  queryset=OoeeState.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))

    class Meta:
        model = OperacionEstadistica
        fields = ['tema', 'entidad', 'area_tematica', 'nombre_est']


class NameFilterOOEE(django_filters.FilterSet):

    nombre_oe = django_filters.CharFilter(lookup_expr='icontains',
    widget=forms.TextInput(
        attrs={'placeholder': 'Buscar', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = OperacionEstadistica
        fields = ['nombre_oe']

