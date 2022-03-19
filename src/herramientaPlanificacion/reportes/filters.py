from demandas.models import demandaInfor, TemaPrincipal, AreaTematica, ddi_estado, InfSolRespRequerimiento
from entities.models import Entidades_oe
import django_filters
from django import forms


class TopicsFilterddi(django_filters.FilterSet):

    tema_prin = django_filters.ModelChoiceFilter(required=False, label="tema",  queryset=TemaPrincipal.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))

    pm_b_6 = django_filters.ModelChoiceFilter(required=False, label="",  queryset=InfSolRespRequerimiento.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))

    area_tem = django_filters.ModelChoiceFilter(required=False, label="área temática",  queryset=AreaTematica.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))

    nombre_est = django_filters.ModelChoiceFilter(required=False, label="Fases",  queryset=ddi_estado.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))

    class Meta:
        model = demandaInfor
        fields = ['tema_prin', 'pm_b_6', 'area_tem', 'nombre_est']


class NameFilterDDI(django_filters.FilterSet):

    pm_b_1 = django_filters.CharFilter(lookup_expr='icontains',
    widget=forms.TextInput(
        attrs={'placeholder': 'Buscar', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = demandaInfor
        fields = ['pm_b_1']

