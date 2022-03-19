from django import forms
from django.core import validators
from django.core.validators import validate_integer
from django.core.exceptions import ValidationError
from demandas.models import ddi_estado, AreaTematica, TemaPrincipal, TemaCompartido, demandaInfor, entidad_service, ooee_service, rraa_service, \
    ComiteEstSect, Identificacionddi, UtilizaInfEst, UsuariosPrincInfEst, Normas, InfSolRespRequerimiento, InfoProdTotalidad, TipoRequerimiento,\
        DesagregacionReq, DesagregacionGeoReq, DesagregacionGeoReqGeografica, DesagregacionGeoReqZona, PeriodicidadDifusion, PeriodicidadDifusionText, \
        UtilizaInfEstText, UsuariosPrincInfEstText, NormasText, InfSolRespRequerimientoText, TipoRequerimientoText, listaVariables, DesagregacionReqText, \
            DesagregacionGeoReqGeograficaText, SolucionReq, NovedadActualizacionddi, TipoNovedad, EstadoActualizacion, OpcAprovooee, \
                OpcAprovrraa, OpcAprovGenNueva, SolReqInformacion, DemandaInsaprio, PlanaccionSuplirDem, EsDemandadeInfor, Comentariosddi, \
                    EstadoCritica, Criticaddi, \
                        TipoEntidadddi, consumidores_info
from django.utils.translation import gettext as _
# url validators
from django.core.validators import URLValidator
from django.forms import modelformset_factory

# forms by Created ooee

FILE_TYPES = ['xls', 'xlsx']
FILE_TYPES_ANNEXES = ['xls', 'xlsx', 'pdf', 'doc', 'docx']


class CreateDDIForm(forms.ModelForm):

    nombre_est = forms.ModelChoiceField(label="Estado",  queryset=ddi_estado.objects.all(),
            widget=forms.Select(attrs={'class': 'form-control form-control-sm', 'onchange':'fondo(this)', 'title':'Estado del proceso de la demanda de información'}))

    area_tem = forms.ModelChoiceField(required=False, label="Area Temática",  queryset=AreaTematica.objects.all(),
            widget=forms.Select(attrs={'class': 'form-control form-control-sm', 'title': 'Módulo A: Área temática'}))

    tema_prin = forms.ModelChoiceField(required=False, label="Tema", queryset=TemaPrincipal.objects.all(),
        widget=forms.Select(attrs={'class': 'form-control form-control-sm', 'title': 'Módulo A: Tema'}))

    tema_comp = forms.ModelMultipleChoiceField(required=False,  widget=forms.SelectMultiple(
        attrs={"class":"selectpicker form-control", "data-live-search":"true", 'title': 'Módulo A: Temas compartidos', 'data-lang':'es_ES',
        'data-selected-text-format':'count', 'data-count-selected-text':"{0} items seleccionados", 'data-size':'5'}),
        queryset=TemaCompartido.objects.all())

    comite_est = forms.ModelMultipleChoiceField(required=False, 
        queryset=ComiteEstSect.objects.all(), widget=forms.CheckboxSelectMultiple(attrs={'title': 'Módulo A: Seleccione el comité estadístico sectorial al que pertenece la demanda de información estadística'}))

    quien_identddi = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Módulo A: ¿Quién identificó la demanda de información?'}),
        queryset=Identificacionddi.objects.all())
        
    entidad_qiddi = forms.ModelMultipleChoiceField(required=False, widget=forms.SelectMultiple(
        attrs={"class":"selectpicker form-control", "data-live-search":"true", 'title': 'Módulo A: Seleccione la entidad o las entidades', 'data-lang':'es_ES',
        'data-selected-text-format':'count', 'data-count-selected-text':"{0} items seleccionados", 'data-size':'5'}),
        queryset=entidad_service.objects.all())
    
    otra_entidad_qiddi = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otra(s)', 'class': 'form-control form-control-sm', 'title': 'Módulo A: escriba el nombre de la entidad'}))

    entidad_sol = forms.ModelChoiceField(required=False, label="Entidad",  queryset=entidad_service.objects.all(), 
        widget=forms.Select(attrs={'class': 'form-control form-control-sm', 'title': 'Módulo A: Nombre de la entidad solicitante'}))

    entidad_cons_sol = forms.ModelChoiceField(required=False, label="Entidad consumidora de información",  queryset=consumidores_info.objects.all(), 
        widget=forms.Select(attrs={'class': 'form-control form-control-sm', 'title': 'Módulo A: Nombre de la entidad consumidora solicitante'}))

    dependencia = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Dependencia', 'class': 'form-control form-control-sm', 'title': 'Módulo A: Dependencia'}))

    nombre_jef_dep = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Nombre jefe de la dependencia', 'class': 'form-control form-control-sm', 'title': 'Módulo A: Nombre del jefe de la dependencia'}))
    
    cargo_jef_dep = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Cargo', 'class': 'form-control form-control-sm', 'title': 'Módulo A: Cargo del jefe de la dependencia'}))
    
    correo_jef_dep = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Correo electrónico', 'class': 'form-control form-control-sm', 'title': 'Módulo A: Correo del jefe de la dependencia'}))
    
    telefono_jef_dep = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Teléfono', 'class': 'form-control form-control-sm', 'title': 'Módulo A: Teléfono del jefe de la dependencia'}))

    pers_req = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Persona que realiza el requerimiento', 'class': 'form-control form-control-sm', 'title': 'Módulo A: Nombre de la persona que realiza el requerimiento'}))

    cargo_pers_req =  forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Cargo', 'class': 'form-control form-control-sm', 'title': 'Módulo A: Cargo de la persona que realiza el requerimiento'}))

    correo_pers_req = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Correo electrónico', 'class': 'form-control form-control-sm', 'title': 'Módulo A: Correo electrónico de la persona que realiza el requerimiento'}))

    telefono_pers_req = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Teléfono', 'class': 'form-control form-control-sm', 'title': 'Módulo A: Teléfono de la persona que realiza el requerimiento'}))

    codigo_ddi = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Código', 'class': 'form-control form-control-sm', 'title': 'Código de la Demanda de Información'}))

    pm_b_1 = forms.CharField(required=False, widget=forms.Textarea(attrs={'class': 'form-control form-control-sm', 'title': 'Módulo B: Pregunta 1 ¿Cuál es el indicador o requerimiento de información estadística?'}))

    pm_b_2 = forms.CharField(required=False, widget=forms.Textarea(attrs={'class': 'form-control form-control-sm', 'title': 'Módulo B: Pregunta 2 Descripción general de la información que requiere'}))

    pm_b_3 = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Módulo B: Pregunta 3 ¿Para qué se utiliza la información estadística?'}), 
        queryset=UtilizaInfEst.objects.all())

    pm_b_4 = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Módulo B: Pregunta 4 ¿Cuáles podrían ser los usuarios principales de la información solicitada?'}),
        queryset=UsuariosPrincInfEst.objects.all())

    pm_b_5 = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Módulo B: Pregunta 5 La información solicitada responde a las siguientes normas'}),
        queryset=Normas.objects.all())

    pm_b_6 = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Módulo B: Pregunta 6 La información solicitada responde a los siguientes requerimientos'}),
        queryset=InfSolRespRequerimiento.objects.all())

    total_si_no_pmb7 = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Módulo B: Pregunta 7 ¿La información requerida está siendo producida en su totalidad?'}),
        queryset=InfoProdTotalidad.objects.all())

    entidad_pm_b7 = forms.ModelMultipleChoiceField(required=False, widget=forms.SelectMultiple(
        attrs={"class":"selectpicker form-control", "data-live-search":"true", 'title': 'Módulo B: Pregunta 7 Seleccione la entidad',  'data-lang':'es_ES',
        'data-selected-text-format':'count', 'data-count-selected-text':"{0} items seleccionados", 'data-size':'5'}),
        queryset=entidad_service.objects.all()) 

    otro_entidad_pm_b7 = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otra', 'class': 'form-control form-control-sm', 'title': 'Módulo B: Pregunta 7 Indique el nombre de la otra entidad'}))

    ooee_pm_b7 = forms.ModelMultipleChoiceField(required=False, widget=forms.SelectMultiple(
        attrs={"class":"selectpicker form-control", "data-live-search":"true", 'title': 'Módulo B: Pregunta 7 Seleccione la Operación estadística', 'data-lang':'es_ES',
        'data-selected-text-format':'count', 'data-count-selected-text':"{0} items seleccionados", 'data-size':'5'}),
        queryset=ooee_service.objects.all()) 
    
    otro_ooee_pm_b7 = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otra', 'class': 'form-control form-control-sm', 'title': 'Módulo B: Pregunta 7 Indique el nombre de la otra ooee'}))

    rraa_pm_b7 = forms.ModelMultipleChoiceField(required=False, widget=forms.SelectMultiple(
        attrs={"class":"selectpicker form-control", "data-live-search":"true",  'title': 'Módulo B: Pregunta 7 Seleccione el registro administrativo', 'data-lang':'es_ES',
        'data-selected-text-format':'count', 'data-count-selected-text':"{0} items seleccionados", 'data-size':'5'}),
        queryset=rraa_service.objects.all()) 

    otro_rraa_pm_b7 = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otra', 'class': 'form-control form-control-sm', 'title': 'Módulo B: Pregunta 7 Indique el nombre del otro rraa'}))

    pm_b_8 = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Módulo B: Pregunta 8 ¿Qué tipo de requerimiento es?'}),
        queryset=TipoRequerimiento.objects.all())

    pm_b_9_anexar = forms.FileField(
        label='Select a file',
        help_text='max. 42 megabytes',
        required=False,
        widget=forms.ClearableFileInput(attrs={"title": "Módulo B: Pregunta 9 ¿Qué variables necesita para suplir el requerimiento?"}))

    pm_b_10 = forms.ModelMultipleChoiceField(required=False, widget=forms.SelectMultiple(
        attrs={"class":"selectpicker form-control", "data-live-search":"true", 'title': 'Módulo B: Pregunta 10 ¿Cuál entidad considera que debe producir la información requerida?', 'data-lang':'es_ES',
        'data-selected-text-format':'count', 'data-count-selected-text':"{0} items seleccionados", 'data-size':'5'}),
        queryset=entidad_service.objects.all())

    pm_b_10_otro = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otra', 'class': 'form-control form-control-sm', 'title': 'Módulo B: Pregunta 10 Otra ¿Cuál entidad considera que debe producir la información requerida?'}))
    
    pm_b_11_1 = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Módulo B: Pregunta 11 ¿Cuál(es) de las siguientes opciones podrían dar solución al requerimiento de información?'}),
        queryset=SolucionReq.objects.all())

    otra_cual_11_1_d  = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otra (s) ¿Cuáles?', 'class': 'form-control form-control-sm mb-1', 
        'title': 'Módulo B: Pregunta 11 Otra (s) ¿Cuáles?'}))

    pm_b_12 = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Módulo B: Pregunta 12 Indique la desagregación requerida'}),
        queryset=DesagregacionReq.objects.all())

    pm_b_13 = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Módulo B: Pregunta 13 Indique la desagregación geográfica requerida'}),
        queryset=DesagregacionGeoReq.objects.all())

    pm_b_13_geo = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Módulo B: Pregunta 13 Indique la desagregación geográfica requerida'}),
        queryset=DesagregacionGeoReqGeografica.objects.all())

    pm_b_13_zona = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Módulo B: Pregunta 13 Indique la desagregación por zona requerida'}),
        queryset=DesagregacionGeoReqZona.objects.all())

    pm_b_14 = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Módulo B: Pregunta 14 Periodicidad de difusión requerida'}),
        queryset=PeriodicidadDifusion.objects.all())

    pm_c_1 = forms.CharField(required=False, widget=forms.Textarea(attrs={'class': 'form-control form-control-sm', 'title': 'Módulo C: Pregunta 1 Observaciones'}))

    pm_d_1_anexos = forms.FileField(
        label='Select a file',
        help_text='max. 42 megabytes',
        required=False,
        widget=forms.ClearableFileInput(attrs={"title": "Módulo D: Anexos"}))

    class Meta:
        model = demandaInfor
        fields = [
            'nombre_est',
            'area_tem',
            'tema_prin',
            'tema_comp',
            'comite_est',
            'quien_identddi',
            'entidad_qiddi',
            'entidad_sol',
            'otra_entidad_qiddi',
            'entidad_cons_sol',
            'dependencia',
            'nombre_jef_dep',
            'cargo_jef_dep',
            'correo_jef_dep',
            'telefono_jef_dep',
            'pers_req',
            'cargo_pers_req',
            'correo_pers_req',
            'telefono_pers_req',
            'pm_b_1',
            'pm_b_2',
            'pm_b_3',
            'pm_b_4',
            'pm_b_5',
            'pm_b_6',
            'total_si_no_pmb7',
            'entidad_pm_b7',
            'otro_entidad_pm_b7',
            'ooee_pm_b7',
            'otro_ooee_pm_b7',
            'rraa_pm_b7',
            'otro_rraa_pm_b7',
            'pm_b_8',
            'pm_b_9_anexar',
            'pm_b_10',
            'pm_b_10_otro',
            'pm_b_11_1',
            'otra_cual_11_1_d',
            'pm_b_12',
            'pm_b_13',
            'pm_b_13_geo',
            'pm_b_13_zona',
            'pm_b_14',
            'pm_c_1',
            'pm_d_1_anexos',
            'codigo_ddi',
        ]
    def clean(self):
        otra_entidad_qiddi = self.cleaned_data.get('otra_entidad_qiddi')
        dependencia = self.cleaned_data.get('dependencia')
        nombre_jef_dep = self.cleaned_data.get('nombre_jef_dep')
        cargo_jef_dep = self.cleaned_data.get('cargo_jef_dep')
        pers_req = self.cleaned_data.get('pers_req')
        cargo_pers_req = self.cleaned_data.get('cargo_pers_req')
        pm_b_1 = self.cleaned_data.get('pm_b_1')
        pm_b_2 = self.cleaned_data.get('pm_b_2')
        otro_entidad_pm_b7 = self.cleaned_data.get('otro_entidad_pm_b7')
        otro_ooee_pm_b7 = self.cleaned_data.get('otro_ooee_pm_b7')
        otro_rraa_pm_b7 = self.cleaned_data.get('otro_rraa_pm_b')
        pm_b_10_otro = self.cleaned_data.get('pm_b_10_otro')
        otra_cual_11_1_d  = self.cleaned_data.get('otra_cual_11_1_d')
        pm_c_1 = self.cleaned_data.get('pm_c_1')

        if otra_entidad_qiddi:
            if len(otra_entidad_qiddi) > 300: 
                self._errors['otra_entidad_qiddi'] = self.error_class([ 
                    'Máximo 300 caracteres requeridos'])

        if dependencia:
            if len(dependencia) > 150: 
                self._errors['dependencia'] = self.error_class([ 
                    'Máximo 150 caracteres requeridos'])

        if nombre_jef_dep:
            if len(nombre_jef_dep) > 150: 
                self._errors['nombre_jef_dep'] = self.error_class([ 
                    'Máximo 150 caracteres requeridos'])

        if cargo_jef_dep:
            if len(cargo_jef_dep) > 150: 
                self._errors['cargo_jef_dep'] = self.error_class([ 
                    'Máximo 150 caracteres requeridos'])

        if pers_req:
            if len(pers_req) > 150: 
                self._errors['pers_req'] = self.error_class([ 
                    'Máximo 150 caracteres requeridos'])

        if cargo_pers_req:
            if len(cargo_pers_req) > 150: 
                self._errors['cargo_pers_req'] = self.error_class([ 
                    'Máximo 150 caracteres requeridos'])

        if pm_b_1:
            if len(pm_b_1) > 2000: 
                self._errors['pm_b_1'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos']) 

        if pm_b_2:
            if len(pm_b_2) > 4000: 
                self._errors['pm_b_2'] = self.error_class([ 
                    'Máximo 4000 caracteres requeridos'])

        if otro_entidad_pm_b7:
            if len(otro_entidad_pm_b7) > 1000: 
                self._errors['otro_entidad_pm_b7'] = self.error_class([ 
                    'Máximo 1000 caracteres requeridos'])

        if otro_ooee_pm_b7:
            if len(otro_ooee_pm_b7) > 1000: 
                self._errors['otro_ooee_pm_b7'] = self.error_class([ 
                    'Máximo 1000 caracteres requeridos'])
        
        if otro_rraa_pm_b7:
            if len(otro_rraa_pm_b7) > 1000: 
                self._errors['otro_rraa_pm_b7'] = self.error_class([ 
                    'Máximo 1000 caracteres requeridos'])

        if pm_b_10_otro:
            if len(pm_b_10_otro) > 900: 
                self._errors['pm_b_10_otro'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])
        
        if otra_cual_11_1_d:
            if len(otra_cual_11_1_d) > 800: 
                self._errors['otra_cual_11_1_d'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        if pm_c_1:
            if len(pm_c_1) > 3000: 
                self._errors['pm_c_1'] = self.error_class([ 
                    'Máximo 3000 caracteres requeridos'])


        return(self.cleaned_data)

    def clean_telefono_jef_dep(self):

        telefono_jef_dep = self.cleaned_data['telefono_jef_dep']
        
        if telefono_jef_dep == '': 
            telefono_jef_dep = None 

        if telefono_jef_dep:
            if len(str(telefono_jef_dep)) < 7 or len(str(telefono_jef_dep)) > 10 or len(str(telefono_jef_dep)) == 8 or len(str(telefono_jef_dep)) == 9:
                raise forms.ValidationError(
                    _("el número telefónico  que ingreso '%s' debe contener 10 digitos para celular o 7 para teléfono fijo. " % telefono_jef_dep))
        return telefono_jef_dep

    def clean_correo_jef_dep(self):
        correo_jef_dep = self.cleaned_data['correo_jef_dep']
        if correo_jef_dep:
            if "@" not in correo_jef_dep:
                raise forms.ValidationError(
                    _("El correo electrónico que ingreso '%s' debe contener @." % correo_jef_dep))

            if len(correo_jef_dep) > 150:
                raise forms.ValidationError(
                    _("El correo electrónico que ingreso '%s' debe contener menos de 150 caracteres" % correo_jef_dep))

        return correo_jef_dep    


    def clean_telefono_pers_req(self):

        telefono_pers_req = self.cleaned_data['telefono_pers_req']

        if telefono_pers_req == '': 
            telefono_pers_req = None 
        
        if telefono_pers_req:
            if len(str(telefono_pers_req)) < 7 or len(str(telefono_pers_req)) > 10  or len(str(telefono_pers_req)) == 8 or len(str(telefono_pers_req)) == 9:
                raise forms.ValidationError(
                    _("el número telefónico  que ingreso '%s' debe contener 10 digitos para celular o 7 para teléfono fijo. " % telefono_pers_req))
        return telefono_pers_req 

    def clean_correo_pers_req(self):
        correo_pers_req = self.cleaned_data['correo_pers_req']
        if correo_pers_req:
            if "@" not in correo_pers_req:
                raise forms.ValidationError(
                    _("El correo electrónico que ingreso '%s' debe contener @." % correo_pers_req))

            if len(correo_pers_req) > 150:
                raise forms.ValidationError(
                    _("El correo electrónico que ingreso '%s' debe contener menos de 150 caracteres" % correo_pers_req))

        return correo_pers_req

    def clean_pm_b_9_anexar(self):
        pm_b_9_anexar = self.cleaned_data['pm_b_9_anexar']
        if pm_b_9_anexar:
            file_type_annexes = str(pm_b_9_anexar).split('.')[-1]
            file_type_annexes = file_type_annexes.lower()
            if file_type_annexes not in FILE_TYPES_ANNEXES:
                raise forms.ValidationError(
                    _("Archivo No valido '%s'  la extensión debe ser XLS, XLSX, PDF, DOC O DOCX " % pm_b_9_anexar))
        return pm_b_9_anexar 

    def clean_pm_d_1_anexos(self):
        pm_d_1_anexos = self.cleaned_data['pm_d_1_anexos']
        if pm_d_1_anexos:
            file_type_annexes = str(pm_d_1_anexos).split('.')[-1]
            file_type_annexes = file_type_annexes.lower()
            if file_type_annexes not in FILE_TYPES_ANNEXES:
                raise forms.ValidationError(
                    _("Archivo No valido '%s'  la extensión debe ser XLS, XLSX, PDF, DOC O DOCX " % pm_d_1_anexos))
        return pm_d_1_anexos     

        

##pregunta 3 modulo b
class CrearUtilizaInfEstTextForm(forms.ModelForm):

    otro_cual = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otro  ¿Cuál?', 'class': 'form-control form-control-sm mb-1', 'title': '¿Para qué se utiliza la información estadística?'}))

    class Meta:
        model = UtilizaInfEstText
        fields = ['otro_cual']

    def clean(self):

        otro_cual = self.cleaned_data.get('otro_cual')
        
        if otro_cual:
            if len(otro_cual) > 900: 
                self._errors['otro_cual'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        return(self.cleaned_data)

##pregunta 4 modulo b
class CrearUsPrincInfEstTextForm(forms.ModelForm):

    orginter_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Organismos internacionales', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: ¿Cuáles podrían ser los usuarios principales de la información solicitada?'}))

    ministerios_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Ministerios', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: ¿Cuáles podrían ser los usuarios principales de la información solicitada?'}))

    orgcontrol_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Organismos de Control', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: ¿Cuáles podrían ser los usuarios principales de la información solicitada?'}))

    oentidadesordenal_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otras entidades del orden Nacional', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: ¿Cuáles podrían ser los usuarios principales de la información solicitada?'}))

    entidadesordenterr_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Entidades de orden Territorial', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: ¿Cuáles podrían ser los usuarios principales de la información solicitada?'}))

    gremios_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Gremios', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: ¿Cuáles podrían ser los usuarios principales de la información solicitada?'}))

    entiprivadas_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Entidades privadas', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: ¿Cuáles podrían ser los usuarios principales de la información solicitada?'}))

    dependenentidad_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Dependencias de la misma entidad', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: ¿Cuáles podrían ser los usuarios principales de la información solicitada?'}))

    academia_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Academia', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: ¿Cuáles podrían ser los usuarios principales de la información solicitada?'}))

    otro_cual_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otra ¿cuál?', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: ¿Cuáles podrían ser los usuarios principales de la información solicitada?'}))

    class Meta:
        model = UsuariosPrincInfEstText
        fields = ['orginter_text', 'ministerios_text', 'orgcontrol_text', 'oentidadesordenal_text',
                    'entidadesordenterr_text', 'gremios_text', 'entiprivadas_text', 'dependenentidad_text',
                    'academia_text', 'otro_cual_text']

    def clean(self):

        orginter_text = self.cleaned_data.get('orginter_text') 
        ministerios_text = self.cleaned_data.get('ministerios_text')
        orgcontrol_text = self.cleaned_data.get('orgcontrol_text')
        oentidadesordenal_text = self.cleaned_data.get('oentidadesordenal_text')
        entidadesordenterr_text = self.cleaned_data.get('entidadesordenterr_text')
        gremios_text = self.cleaned_data.get('gremios_text')
        entiprivadas_text = self.cleaned_data.get('entiprivadas_text')
        dependenentidad_text = self.cleaned_data.get('dependenentidad_text')
        academia_text = self.cleaned_data.get('academia_text')
        otro_cual_text = self.cleaned_data.get('otro_cual_text')
        
        if orginter_text:
            if len(orginter_text) > 900: 
                self._errors['orginter_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if ministerios_text:
            if len(ministerios_text) > 900: 
                self._errors['ministerios_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if orgcontrol_text:
            if len(orgcontrol_text) > 900: 
                self._errors['orgcontrol_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if oentidadesordenal_text:
            if len(oentidadesordenal_text) > 900: 
                self._errors['oentidadesordenal_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if entidadesordenterr_text:
            if len(entidadesordenterr_text) > 900: 
                self._errors['entidadesordenterr_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if gremios_text:
            if len(gremios_text) > 900: 
                self._errors['gremios_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if entiprivadas_text:
            if len(entiprivadas_text) > 900: 
                self._errors['entiprivadas_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if dependenentidad_text:
            if len(dependenentidad_text) > 900: 
                self._errors['dependenentidad_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if academia_text:
            if len(academia_text) > 900: 
                self._errors['academia_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if otro_cual_text:
            if len(otro_cual_text) > 900: 
                self._errors['otro_cual_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        return(self.cleaned_data)


##pregunta 5 modulo b
class CrearNormasTextForm(forms.ModelForm):

    const_pol_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Constitución Política', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: La información solicitada responde a las siguientes normas'})) 
    ley_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Ley', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: La información solicitada responde a las siguientes normas'}))
    decreto_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Decreto (Nacional, departamental, etc.)', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: La información solicitada responde a las siguientes normas'}))
    otra_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otra (Resolución, ordenanza, acuerdo)', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: La información solicitada responde a las siguientes normas'}))

    class Meta:
        model = NormasText
        fields = ['const_pol_text', 'ley_text', 'decreto_text', 'otra_text']

    def clean(self):

        const_pol_text = self.cleaned_data.get('const_pol_text') 
        ley_text = self.cleaned_data.get('const_pol_text') 
        decreto_text = self.cleaned_data.get('const_pol_text') 
        otra_text = self.cleaned_data.get('const_pol_text') 
        
        if const_pol_text:
            if len(const_pol_text) > 900: 
                self._errors['const_pol_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if ley_text:
            if len(ley_text) > 900: 
                self._errors['ley_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if decreto_text:
            if len(decreto_text) > 900: 
                self._errors['decreto_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if otra_text:
            if len(otra_text) > 900: 
                self._errors['otra_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        return(self.cleaned_data)


##pregunta 6 modulo b
class CrearInfSolRespRequerimientoTextForm(forms.ModelForm):

    planalDes_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Plan Nacional de Desarrollo (Capítulo y línea)', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: La información solicitada responde a los siguientes requerimientos'})) 
    
    cuentasecomacroec_text =  forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Cuentas económicas y macroeconómicas', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: La información solicitada responde a los siguientes requerimientos'})) 

    plansecterrcom_text =  forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Plan Sectorial, Territorial o CONPES', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: La información solicitada responde a los siguientes requerimientos'})) 

    objdessost_text =  forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Objetivos de Desarrollo Sostenible (ODS)', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: La información solicitada responde a los siguientes requerimientos'})) 

    orgcooper_text =  forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Organización para la Cooperación y el Desarrollo Económicos (OCDE)', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: La información solicitada responde a los siguientes requerimientos'})) 

    otroscomprInt_text =  forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otros compromisos internacionales', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: La información solicitada responde a los siguientes requerimientos'})) 

    otros_text =  forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otro(s)', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: La información solicitada responde a los siguientes requerimientos'})) 


    class Meta:
        model = InfSolRespRequerimientoText
        fields = ['planalDes_text', 'cuentasecomacroec_text', 'plansecterrcom_text', 'objdessost_text',
                    'orgcooper_text', 'otroscomprInt_text', 'otros_text']

    def clean(self):

        planalDes_text = self.cleaned_data.get('planalDes_text') 
        cuentasecomacroec_text = self.cleaned_data.get('cuentasecomacroec_text') 
        plansecterrcom_text = self.cleaned_data.get('plansecterrcom_text') 
        objdessost_text = self.cleaned_data.get('objdessost_text') 
        orgcooper_text = self.cleaned_data.get('orgcooper_text') 
        otroscomprInt_text = self.cleaned_data.get('otroscomprInt_text') 
        otros_text = self.cleaned_data.get('otros_text') 
        
        if planalDes_text:
            if len(planalDes_text) > 900: 
                self._errors['planalDes_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if cuentasecomacroec_text:
            if len(cuentasecomacroec_text) > 900: 
                self._errors['cuentasecomacroec_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if plansecterrcom_text:
            if len(plansecterrcom_text) > 900: 
                self._errors['plansecterrcom_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if objdessost_text:
            if len(objdessost_text) > 900: 
                self._errors['objdessost_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if orgcooper_text:
            if len(orgcooper_text) > 900: 
                self._errors['orgcooper_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])
 
        if otroscomprInt_text:
            if len(otroscomprInt_text) > 900: 
                self._errors['otroscomprInt_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if otros_text:
            if len(otros_text) > 900: 
                self._errors['otros_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        return(self.cleaned_data)


#pregunta 8 modulo b
class TipoRequerimientoTextForm(forms.ModelForm):

    otros_c_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otro  ¿Cuál?', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: ¿Qué tipo de requerimiento es?'}))

    class Meta:
        model = TipoRequerimientoText
        fields = ['otros_c_text']

    def clean(self):

        otros_c_text = self.cleaned_data.get('otros_c_text')
        
        if otros_c_text:
            if len(otros_c_text) > 900: 
                self._errors['otros_c_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        return(self.cleaned_data)

#pregunta 9 modulo b
class listaVariablesForm(forms.ModelForm):

    lista_varia = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'title': 'Módulo B: ¿Qué variables necesita para suplir el requerimiento?',
            'placeholder': '¿Qué variables necesita para suplir el requerimiento?', 
            'class': 'form-control form-control-sm'}))

    class Meta:
        model = listaVariables
        fields = [
            'lista_varia'
        ]

    def clean(self):

        lista_varia = self.cleaned_data.get('lista_varia')
        
        if lista_varia:
            if len(lista_varia) > 900: 
                self._errors['lista_varia'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

         # return any errors if found 
        return(self.cleaned_data)

# modulo B pregunta 7  ***formsets
listaVariablesFormset = modelformset_factory(
    listaVariables,
    form=listaVariablesForm,
    fields=('lista_varia',),
    can_delete=True,
    extra=1,
)


### pregunta 11 

class CrearSolReqInformacionForm(forms.ModelForm):  ##testPreguntaOnceForm
    
    ooee_pmb11 = forms.ModelChoiceField(required=False, label="Operaciones estadísticas",  queryset=ooee_service.objects.all(),
            widget=forms.Select(attrs={'class': 'form-control form-control-sm',}))

    opc_apr_oe = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple,
        queryset=OpcAprovooee.objects.all())
    
    rraa_pmb11 = forms.ModelChoiceField(required=False, label="Registros administrativos",  queryset=rraa_service.objects.all(),
            widget=forms.Select(attrs={'class': 'form-control form-control-sm',}))

    opc_apr_ra = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple,
        queryset=OpcAprovrraa.objects.all())

    genera_nuev = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple,
        queryset=OpcAprovGenNueva.objects.all())

    inc_var_cual = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'a)2.1. Incluir variables/preguntas ¿Cuál(es)?', 'class': 'form-control form-control-sm mb-1', 
        'title': 'Módulo B: pregunta 11 a)2.1. Incluir variables/preguntas ¿Cuál(es)?'}))
    
    cam_preg_cual = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'a)2.2. Cambiar la formulación de alguna(s) pregunta(s) ¿Cuál(es)?', 'class': 'form-control form-control-sm mb-1', 
        'title': 'Módulo B: pregunta 11 a)2.2. Cambiar la formulación de alguna(s) pregunta(s) ¿Cuál(es)?'}))

    am_des_tem_cual =  forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'a)2.3. Ampliar la desagregación temática ¿Cuál(es)?', 'class': 'form-control form-control-sm mb-1', 
        'title': 'Módulo B: pregunta 11 a)2.3. Ampliar la desagregación temática ¿Cuál(es)?'}))

    am_des_geo_cual = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'a)2.4. Ampliar la desagregación geográfica ¿Cuál(es)?', 'class': 'form-control form-control-sm mb-1', 
        'title': 'Módulo B: pregunta 11 a)2.4. Ampliar la desagregación geográfica ¿Cuál(es)?'}))

    dif_resul_cual = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'a)2.5. Aumentar la frecuencia en la difusión de los resultados ¿Cuál(es)?', 'class': 'form-control form-control-sm mb-1', 
        'title': 'Módulo B: pregunta 11 a)2.5. Aumentar la frecuencia en la difusión de los resultados ¿Cuál(es)?'}))

    opc_aprov_cual = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'a)2.6. Otra ¿Cuál(es)?', 'class': 'form-control form-control-sm mb-1', 
        'title': 'Módulo B: a)2.6. Otra ¿Cuál(es)?'}))

    inc_varia_cual = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'b)2.1.  Inclusión de variables/preguntas', 'class': 'form-control form-control-sm mb-1', 
        'title': 'Módulo B: pregunta 11 Inclusión de variables/preguntas'}))
    
    camb_pregu_cual = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'b)2.2. Cambios en la formulación de alguna(s) pregunta(s)', 'class': 'form-control form-control-sm mb-1', 
        'title': 'Módulo B: Pregunta 11 b)2.2. Cambios en la formulación de alguna(s) pregunta(s)'}))

    otros_aprov_ra = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'b)2.3. Otro ', 'class': 'form-control form-control-sm mb-1', 
        'title': 'Módulo B: pregunta 11 b)2.3. Otro '}))

    nueva_oe = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'c)1. Operación estadística nueva', 'class': 'form-control form-control-sm mb-1', 
        'title': 'Módulo B: pregunta 11 c)1. Operación estadística nueva'}))
    
    indi_cual = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'c)2. Indicador ¿Cuál(es)?', 'class': 'form-control form-control-sm mb-1', 
        'title': 'Módulo B: pregunta 11 c)2. Indicador ¿Cuál(es)?'}))

    gen_nueva = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'c)3. Otra  ¿Cuál?', 'class': 'form-control form-control-sm mb-1', 
        'title': 'Módulo B: pregunta 11 c)3. Otra  ¿Cuál?'}))

    class Meta:
        model = SolReqInformacion
        fields = ['ooee_pmb11','opc_apr_oe','rraa_pmb11','opc_apr_ra','genera_nuev',
                    'inc_var_cual','cam_preg_cual','am_des_tem_cual', 'am_des_geo_cual','dif_resul_cual',
                    'opc_aprov_cual','inc_varia_cual','camb_pregu_cual','otros_aprov_ra','nueva_oe','indi_cual','gen_nueva']

    def clean(self):

        inc_var_cual = self.cleaned_data.get('inc_var_cual')
        cam_preg_cual = self.cleaned_data.get('cam_preg_cual')
        am_des_tem_cual = self.cleaned_data.get('am_des_tem_cual')
        am_des_geo_cual = self.cleaned_data.get('am_des_geo_cual')
        dif_resul_cual = self.cleaned_data.get('dif_resul_cual')
        opc_aprov_cual = self.cleaned_data.get('opc_aprov_cual')

        inc_varia_cual = self.cleaned_data.get('inc_varia_cual')
        camb_pregu_cual = self.cleaned_data.get('camb_pregu_cual')
        otros_aprov_ra = self.cleaned_data.get('otros_aprov_ra')

        nueva_oe = self.cleaned_data.get('nueva_oe')
        indi_cual = self.cleaned_data.get('indi_cual')
        gen_nueva = self.cleaned_data.get('gen_nueva')

        if inc_var_cual:
            if len(inc_var_cual) > 800: 
                self._errors['inc_var_cual'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        if cam_preg_cual:
            if len(cam_preg_cual) > 800: 
                self._errors['cam_preg_cual'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        if am_des_tem_cual:
            if len(am_des_tem_cual) > 800: 
                self._errors['am_des_tem_cual'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        if am_des_geo_cual:
            if len(am_des_geo_cual) > 800: 
                self._errors['am_des_geo_cual'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        if dif_resul_cual:
            if len(dif_resul_cual) > 800: 
                self._errors['dif_resul_cual'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        if opc_aprov_cual:
            if len(opc_aprov_cual) > 800: 
                self._errors['opc_aprov_cual'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])
        
        if inc_varia_cual:
            if len(inc_varia_cual) > 800: 
                self._errors['inc_varia_cual'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        if camb_pregu_cual:
            if len(camb_pregu_cual) > 800: 
                self._errors['camb_pregu_cual'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        if otros_aprov_ra:
            if len(otros_aprov_ra) > 800: 
                self._errors['otros_aprov_ra'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])
        
        if nueva_oe:
            if len(nueva_oe) > 800: 
                self._errors['nueva_oe'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        if indi_cual:
            if len(indi_cual) > 800: 
                self._errors['indi_cual'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        if gen_nueva:
            if len(gen_nueva) > 800: 
                self._errors['gen_nueva'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])
        
        return(self.cleaned_data)


###pregunta 12 modulo b
class DesagregacionReqTextForm(forms.ModelForm):

    otra_cual_text_a = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otro  ¿Cuál?', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: Indique la desagregación requerida'}))

    class Meta:
        model = DesagregacionReqText
        fields = ['otra_cual_text_a']

    def clean(self):

        otra_cual_text_a = self.cleaned_data.get('otra_cual_text_a')
        
        if otra_cual_text_a:
            if len(otra_cual_text_a) > 900: 
                self._errors['otra_cual_text_a'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        return(self.cleaned_data)

###pregunta 13 modulo b
class DesagregacionGeoReqGeograficaTextForm(forms.ModelForm):

    otra_cual_text_b = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otro  ¿Cuál?', 'class': 'form-control form-control-sm mb-1', 'title': 'Indique la desagregación geográfica requerida'}))

    class Meta:
        model = DesagregacionGeoReqGeograficaText
        fields = ['otra_cual_text_b']

    def clean(self):

        otra_cual_text_b = self.cleaned_data.get('otra_cual_text_b')
        
        if otra_cual_text_b:
            if len(otra_cual_text_b) > 900: 
                self._errors['otra_cual_text_b'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        return(self.cleaned_data)

###pregunta 14 modulo b
class PeriodicidadDifusionTextForm(forms.ModelForm):

    otra_cual_text_c = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otro  ¿Cuál?', 'class': 'form-control form-control-sm mb-1', 'title': 'Periodicidad de difusión requerida'}))

    class Meta:
        model = PeriodicidadDifusionText
        fields = ['otra_cual_text_c']

    def clean(self):

        otra_cual_text_c = self.cleaned_data.get('otra_cual_text_c')
        
        if otra_cual_text_c:
            if len(otra_cual_text_c) > 900: 
                self._errors['otra_cual_text_c'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        return(self.cleaned_data)

# formulario para actualización novedades

class NovedadDDIForm(forms.ModelForm):

    descrip_novedad = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Descripción de la novedad', 'class': 'form-control form-control-sm'}))
    
    novedad = forms.ModelChoiceField(required=False, label="tipo de novedad",  queryset=TipoNovedad.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))
    
    est_actualiz = forms.ModelChoiceField(required=False, label="estado de actualización",  queryset=EstadoActualizacion.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))

    class Meta:
        model = NovedadActualizacionddi
        fields = ['descrip_novedad', 'novedad', 'est_actualiz']


    def clean(self):
            
        descrip_novedad = self.cleaned_data.get('descrip_novedad')
    
        if descrip_novedad:
            if len(descrip_novedad) > 1000: 
                self._errors['descrip_novedad'] = self.error_class([ 
                    'Máximo 1000 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)


#### formulario para los comentarios  ##############
class ComentariosddiForm(forms.ModelForm):
    body = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Describa aquí las inconsistencias sobre la información diligenciada en el formulario', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = Comentariosddi
        fields = ['body']

    def clean(self):
            
        body = self.cleaned_data.get('body')
    
        if body:
            if len(body) > 2000:
                self._errors['body'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data) 

#### end formulario para los comentarios  ##############


class CriticaddiForm(forms.ModelForm):

    descrip_critica = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Observaciones de la crítica', 'class': 'form-control form-control-sm'}))
    estado_crit = forms.ModelChoiceField(label="tipo de novedad", empty_label=None, queryset=EstadoCritica.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))

    class Meta:
        model = Criticaddi
        fields = ['descrip_critica', 'estado_crit']

    def clean(self):
            
        descrip_critica = self.cleaned_data.get('descrip_critica')
    
        if descrip_critica:
            if len(descrip_critica) > 1000: 
                self._errors['descrip_critica'] = self.error_class([ 
                    'Máximo 1000 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)



## formulario editar ddi
class EditDDIForm(forms.ModelForm):

    nombre_est = forms.ModelChoiceField(label="Estado",  queryset=ddi_estado.objects.all(),
            widget=forms.Select(attrs={'class': 'form-control form-control-sm', 'onchange':'fondo(this)', 'title':'Estado del proceso de la demanda de información'}))

    area_tem = forms.ModelChoiceField(required=False, label="Area Temática",  queryset=AreaTematica.objects.all(),
            widget=forms.Select(attrs={'class': 'form-control form-control-sm', 'title': 'Módulo A: Área temática'}))

    tema_prin = forms.ModelChoiceField(required=False, label="Tema", queryset=TemaPrincipal.objects.all(),
        widget=forms.Select(attrs={'class': 'form-control form-control-sm', 'title': 'Módulo A: Tema'}))

    tema_comp = forms.ModelMultipleChoiceField(required=False,  widget=forms.SelectMultiple(
        attrs={"class":"selectpicker form-control", "data-live-search":"true", 'title': 'Módulo A: Temas compartidos', 'data-lang':'es_ES',
        'data-selected-text-format':'count', 'data-count-selected-text':"{0} items seleccionados", 'data-size':'5'}),
        queryset=TemaCompartido.objects.all())

    comite_est = forms.ModelMultipleChoiceField(required=False, 
        queryset=ComiteEstSect.objects.all(), widget=forms.CheckboxSelectMultiple(attrs={'title': 'Módulo A: Seleccione el comité estadístico sectorial al que pertenece la demanda de información estadística'}))

    quien_identddi = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Módulo A: ¿Quién identificó la demanda de información?'}),
        queryset=Identificacionddi.objects.all())
        
    entidad_qiddi = forms.ModelMultipleChoiceField(required=False, widget=forms.SelectMultiple(
        attrs={"class":"selectpicker form-control", "data-live-search":"true", 'title': 'Módulo A: Seleccione la entidad o las entidades', 'data-lang':'es_ES',
        'data-selected-text-format':'count', 'data-count-selected-text':"{0} items seleccionados", 'data-size':'5', 'data-width':"90%"}),
        queryset=entidad_service.objects.all())
    
    otra_entidad_qiddi = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otra(s)', 'class': 'form-control form-control-sm', 'title': 'Módulo A: escriba el nombre de la entidad'}))

    entidad_sol = forms.ModelChoiceField(required=False, label="Entidad",  queryset=entidad_service.objects.all(), 
        widget=forms.Select(attrs={'class': 'form-control form-control-sm', 'title': 'Módulo A: Nombre de la entidad solicitante'}))


    entidad_cons_sol = forms.ModelChoiceField(required=False, label="Entidad consumidora de información",  queryset=consumidores_info.objects.all(), 
        widget=forms.Select(attrs={'class': 'form-control form-control-sm', 'title': 'Módulo A: Nombre de la entidad consumidora solicitante'}))

    dependencia = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Dependencia', 'class': 'form-control form-control-sm', 'title': 'Módulo A: Dependencia'}))

    nombre_jef_dep = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Nombre jefe de la dependencia', 'class': 'form-control form-control-sm', 'title': 'Módulo A: Nombre del jefe de la dependencia'}))
    
    cargo_jef_dep = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Cargo', 'class': 'form-control form-control-sm', 'title': 'Módulo A: Cargo del jefe de la dependencia'}))
    
    correo_jef_dep = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Correo electrónico', 'class': 'form-control form-control-sm', 'title': 'Módulo A: Correo del jefe de la dependencia'}))
    
    telefono_jef_dep = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Teléfono', 'class': 'form-control form-control-sm', 'title': 'Módulo A: Teléfono del jefe de la dependencia'}))

    pers_req = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Persona que realiza el requerimiento', 'class': 'form-control form-control-sm', 'title': 'Módulo A: Nombre de la persona que realiza el requerimiento'}))

    cargo_pers_req =  forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Cargo', 'class': 'form-control form-control-sm', 'title': 'Módulo A: Cargo de la persona que realiza el requerimiento'}))

    correo_pers_req = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Correo electrónico', 'class': 'form-control form-control-sm', 'title': 'Módulo A: Correo electrónico de la persona que realiza el requerimiento'}))

    telefono_pers_req = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Teléfono', 'class': 'form-control form-control-sm', 'title': 'Módulo A: Teléfono de la persona que realiza el requerimiento'}))

    codigo_ddi = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Código', 'class': 'form-control form-control-sm', 'title': 'Código de la Demanda de Información'}))

    pm_b_1 = forms.CharField(required=False, widget=forms.Textarea(attrs={'class': 'form-control form-control-sm', 'title': 'Módulo B: Pregunta 1 ¿Cuál es el indicador o requerimiento de información estadística?'}))

    pm_b_2 = forms.CharField(required=False, widget=forms.Textarea(attrs={'class': 'form-control form-control-sm', 'title': 'Módulo B: Pregunta 2 Descripción general de la información que requiere'}))

    pm_b_3 = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Módulo B: Pregunta 3 ¿Para qué se utiliza la información estadística?'}), 
        queryset=UtilizaInfEst.objects.all())

    pm_b_4 = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Módulo B: Pregunta 4 ¿Cuáles podrían ser los usuarios principales de la información solicitada?'}),
        queryset=UsuariosPrincInfEst.objects.all())

    pm_b_5 = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Módulo B: Pregunta 5 La información solicitada responde a las siguientes normas'}),
        queryset=Normas.objects.all())

    pm_b_6 = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Módulo B: Pregunta 6 La información solicitada responde a los siguientes requerimientos'}),
        queryset=InfSolRespRequerimiento.objects.all())

    total_si_no_pmb7 = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Módulo B: Pregunta 7 ¿La información requerida está siendo producida en su totalidad?'}),
        queryset=InfoProdTotalidad.objects.all())

    entidad_pm_b7 = forms.ModelMultipleChoiceField(required=False, widget=forms.SelectMultiple(
        attrs={"class":"selectpicker form-control", "data-live-search":"true", 'title': 'Módulo B: Pregunta 7 Seleccione la entidad',  'data-lang':'es_ES',
        'data-selected-text-format':'count', 'data-count-selected-text':"{0} items seleccionados", 'data-size':'5'}),
        queryset=entidad_service.objects.all()) 

    otro_entidad_pm_b7 = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otra', 'class': 'form-control form-control-sm', 'title': 'Módulo B: Pregunta 7 Indique el nombre de la otra entidad'}))

    ooee_pm_b7 = forms.ModelMultipleChoiceField(required=False, widget=forms.SelectMultiple(
        attrs={"class":"selectpicker form-control", "data-live-search":"true", 'title': 'Módulo B: Pregunta 7 Seleccione la Operación estadística', 'data-lang':'es_ES',
        'data-selected-text-format':'count', 'data-count-selected-text':"{0} items seleccionados", 'data-size':'5'}),
        queryset=ooee_service.objects.all()) 
    
    otro_ooee_pm_b7 = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otra', 'class': 'form-control form-control-sm', 'title': 'Módulo B: Pregunta 7 Indique el nombre de la otra ooee'}))

    rraa_pm_b7 = forms.ModelMultipleChoiceField(required=False, widget=forms.SelectMultiple(
        attrs={"class":"selectpicker form-control", "data-live-search":"true",  'title': 'Módulo B: Pregunta 7 Seleccione el registro administrativo', 'data-lang':'es_ES',
        'data-selected-text-format':'count', 'data-count-selected-text':"{0} items seleccionados", 'data-size':'5'}),
        queryset=rraa_service.objects.all()) 

    otro_rraa_pm_b7 = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otra', 'class': 'form-control form-control-sm', 'title': 'Módulo B: Pregunta 7 Indique el nombre del otro rraa'}))

    pm_b_8 = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Módulo B: Pregunta 8 ¿Qué tipo de requerimiento es?'}),
        queryset=TipoRequerimiento.objects.all())

    pm_b_9_anexar = forms.FileField(
        label='Select a file',
        help_text='max. 42 megabytes',
        required=False,
        widget=forms.ClearableFileInput(attrs={"title": "Módulo B: Pregunta 9 ¿Qué variables necesita para suplir el requerimiento?"}))

    pm_b_10 = forms.ModelMultipleChoiceField(required=False, widget=forms.SelectMultiple(
        attrs={"class":"selectpicker form-control", "data-live-search":"true", 'title': 'Módulo B: Pregunta 10 ¿Cuál entidad considera que debe producir la información requerida?', 'data-lang':'es_ES',
        'data-selected-text-format':'count', 'data-count-selected-text':"{0} items seleccionados", 'data-size':'5'}),
        queryset=entidad_service.objects.all()) 
    
    pm_b_10_otro = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otra', 'class': 'form-control form-control-sm', 'title': 'Módulo B: Pregunta 10 Otra ¿Cuál entidad considera que debe producir la información requerida?'}))
    
    pm_b_11_1 = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Módulo B: Pregunta 11 ¿Cuál(es) de las siguientes opciones podrían dar solución al requerimiento de información?'}),
        queryset=SolucionReq.objects.all())

    otra_cual_11_1_d  = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otra (s) ¿Cuáles?', 'class': 'form-control form-control-sm mb-1', 
        'title': 'Módulo B: Pregunta 11 Otra (s) ¿Cuáles?'}))

    pm_b_12 = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Módulo B: Pregunta 12 Indique la desagregación requerida'}),
        queryset=DesagregacionReq.objects.all())

    pm_b_13 = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Módulo B: Pregunta 13 Indique la desagregación geográfica requerida'}),
        queryset=DesagregacionGeoReq.objects.all())

    pm_b_13_geo = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Módulo B: Pregunta 13 Indique la desagregación geográfica requerida'}),
        queryset=DesagregacionGeoReqGeografica.objects.all())

    pm_b_13_zona = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Módulo B: Pregunta 13 Indique la desagregación por zona requerida'}),
        queryset=DesagregacionGeoReqZona.objects.all())

    pm_b_14 = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Módulo B: Pregunta 14 Periodicidad de difusión requerida'}),
        queryset=PeriodicidadDifusion.objects.all())

    pm_c_1 = forms.CharField(required=False, widget=forms.Textarea(attrs={'class': 'form-control form-control-sm', 'title': 'Módulo C:  Pregunta 1 Observaciones'}))

    pm_d_1_anexos = forms.FileField(
        label='Select a file',
        help_text='max. 42 megabytes',
        required=False,
        widget=forms.ClearableFileInput(attrs={"title": "Módulo D: Anexos"}))

    compl_dem_a = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Complemento de demanda'}),
        queryset=DemandaInsaprio.objects.all())

    compl_dem_a_text = forms.CharField(required=False, widget=forms.Textarea(attrs={'class': 'form-control form-control-sm', 'title': 'Complemento de demanda'}))
    
    comple_dem_b = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Complemento de demanda'}),
        queryset=PlanaccionSuplirDem.objects.all())
    
    compl_dem_b_text = forms.CharField(required=False, widget=forms.Textarea(attrs={'class': 'form-control form-control-sm', 'title': 'Complemento de demanda'}))

    validacion_ddi = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(attrs={'title': 'Validación de demanda'}),
        queryset=EsDemandadeInfor.objects.all())

    validacion_ddi_text = forms.CharField(required=False, widget=forms.Textarea(attrs={'class': 'form-control form-control-sm', 'title': 'Validación de demanda'}))

    class Meta:
        model = demandaInfor
        fields = [
            'nombre_est',
            'area_tem',
            'tema_prin',
            'tema_comp',
            'comite_est',
            'quien_identddi',
            'entidad_qiddi',
            'entidad_sol',
            'otra_entidad_qiddi',
            'entidad_cons_sol',
            'dependencia',
            'nombre_jef_dep',
            'cargo_jef_dep',
            'correo_jef_dep',
            'telefono_jef_dep',
            'pers_req',
            'cargo_pers_req',
            'correo_pers_req',
            'telefono_pers_req',
            'pm_b_1',
            'pm_b_2',
            'pm_b_3',
            'pm_b_4',
            'pm_b_5',
            'pm_b_6',
            'total_si_no_pmb7',
            'entidad_pm_b7',
            'otro_entidad_pm_b7',
            'ooee_pm_b7',
            'otro_ooee_pm_b7',
            'rraa_pm_b7',
            'otro_rraa_pm_b7',
            'pm_b_8',
            'pm_b_9_anexar',
            'pm_b_10',
            'pm_b_10_otro',
            'pm_b_11_1',
            'otra_cual_11_1_d',
            'pm_b_12',
            'pm_b_13',
            'pm_b_13_geo',
            'pm_b_13_zona',
            'pm_b_14',
            'pm_c_1',
            'pm_d_1_anexos',
            'codigo_ddi',
            'compl_dem_a',
            'compl_dem_a_text',
            'comple_dem_b',
            'compl_dem_b_text',
            'validacion_ddi',
            'validacion_ddi_text',
        ]

    def clean(self):
        otra_entidad_qiddi = self.cleaned_data.get('otra_entidad_qiddi')
        dependencia = self.cleaned_data.get('dependencia')
        nombre_jef_dep = self.cleaned_data.get('nombre_jef_dep')
        cargo_jef_dep = self.cleaned_data.get('cargo_jef_dep')
        pers_req = self.cleaned_data.get('pers_req')
        cargo_pers_req = self.cleaned_data.get('cargo_pers_req')
        pm_b_1 = self.cleaned_data.get('pm_b_1')
        pm_b_2 = self.cleaned_data.get('pm_b_2')
        otro_entidad_pm_b7 = self.cleaned_data.get('otro_entidad_pm_b7')
        otro_ooee_pm_b7 = self.cleaned_data.get('otro_ooee_pm_b7')
        otro_rraa_pm_b7 = self.cleaned_data.get('otro_rraa_pm_b')
        pm_b_10_otro = self.cleaned_data.get('pm_b_10_otro')
        otra_cual_11_1_d  = self.cleaned_data.get('otra_cual_11_1_d')
        pm_c_1 = self.cleaned_data.get('pm_c_1')

        if otra_entidad_qiddi:
            if len(otra_entidad_qiddi) > 300: 
                self._errors['otra_entidad_qiddi'] = self.error_class([ 
                    'Máximo 300 caracteres requeridos'])

        if dependencia:
            if len(dependencia) > 150: 
                self._errors['dependencia'] = self.error_class([ 
                    'Máximo 150 caracteres requeridos'])

        if nombre_jef_dep:
            if len(nombre_jef_dep) > 150: 
                self._errors['nombre_jef_dep'] = self.error_class([ 
                    'Máximo 150 caracteres requeridos'])

        if cargo_jef_dep:
            if len(cargo_jef_dep) > 150: 
                self._errors['cargo_jef_dep'] = self.error_class([ 
                    'Máximo 150 caracteres requeridos'])

        if pers_req:
            if len(pers_req) > 150: 
                self._errors['pers_req'] = self.error_class([ 
                    'Máximo 150 caracteres requeridos'])

        if cargo_pers_req:
            if len(cargo_pers_req) > 150: 
                self._errors['cargo_pers_req'] = self.error_class([ 
                    'Máximo 150 caracteres requeridos'])

        if pm_b_1:
            if len(pm_b_1) > 2000: 
                self._errors['pm_b_1'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos']) 

        if pm_b_2:
            if len(pm_b_2) > 4000: 
                self._errors['pm_b_2'] = self.error_class([ 
                    'Máximo 4000 caracteres requeridos'])

        if otro_entidad_pm_b7:
            if len(otro_entidad_pm_b7) > 1000: 
                self._errors['otro_entidad_pm_b7'] = self.error_class([ 
                    'Máximo 1000 caracteres requeridos'])

        if otro_ooee_pm_b7:
            if len(otro_ooee_pm_b7) > 1000: 
                self._errors['otro_ooee_pm_b7'] = self.error_class([ 
                    'Máximo 1000 caracteres requeridos'])
        
        if otro_rraa_pm_b7:
            if len(otro_rraa_pm_b7) > 1000: 
                self._errors['otro_rraa_pm_b7'] = self.error_class([ 
                    'Máximo 1000 caracteres requeridos'])

        if pm_b_10_otro:
            if len(pm_b_10_otro) > 900: 
                self._errors['pm_b_10_otro'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if otra_cual_11_1_d:
            if len(otra_cual_11_1_d) > 800: 
                self._errors['otra_cual_11_1_d'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        if pm_c_1:
            if len(pm_c_1) > 3000: 
                self._errors['pm_c_1'] = self.error_class([ 
                    'Máximo 3000 caracteres requeridos'])


        return(self.cleaned_data)

    def clean_telefono_jef_dep(self):

        telefono_jef_dep = self.cleaned_data['telefono_jef_dep']
        
        if telefono_jef_dep == '': 
            telefono_jef_dep = None 

        if telefono_jef_dep:
            if len(str(telefono_jef_dep)) < 7 or len(str(telefono_jef_dep)) > 10 or len(str(telefono_jef_dep)) == 8 or len(str(telefono_jef_dep)) == 9:
                raise forms.ValidationError(
                    _("el número telefónico  que ingreso '%s' debe contener 10 digitos para celular o 7 para teléfono fijo. " % telefono_jef_dep))
        return telefono_jef_dep

    def clean_correo_jef_dep(self):
        correo_jef_dep = self.cleaned_data['correo_jef_dep']
        if correo_jef_dep:
            if "@" not in correo_jef_dep:
                raise forms.ValidationError(
                    _("El correo electrónico que ingreso '%s' debe contener @." % correo_jef_dep))

            if len(correo_jef_dep) > 150:
                raise forms.ValidationError(
                    _("El correo electrónico que ingreso '%s' debe contener menos de 150 caracteres" % correo_jef_dep))

        return correo_jef_dep    


    def clean_telefono_pers_req(self):

        telefono_pers_req = self.cleaned_data['telefono_pers_req']

        if telefono_pers_req == '': 
            telefono_pers_req = None 
        
        if telefono_pers_req:
            if len(str(telefono_pers_req)) < 7 or len(str(telefono_pers_req)) > 10  or len(str(telefono_pers_req)) == 8 or len(str(telefono_pers_req)) == 9:
                raise forms.ValidationError(
                    _("el número telefónico  que ingreso '%s' debe contener 10 digitos para celular o 7 para teléfono fijo. " % telefono_pers_req))
        return telefono_pers_req 

    def clean_correo_pers_req(self):
        correo_pers_req = self.cleaned_data['correo_pers_req']
        if correo_pers_req:
            if "@" not in correo_pers_req:
                raise forms.ValidationError(
                    _("El correo electrónico que ingreso '%s' debe contener @." % correo_pers_req))

            if len(correo_pers_req) > 150:
                raise forms.ValidationError(
                    _("El correo electrónico que ingreso '%s' debe contener menos de 150 caracteres" % correo_pers_req))

        return correo_pers_req

    def clean_pm_b_9_anexar(self):
        pm_b_9_anexar = self.cleaned_data['pm_b_9_anexar']
        if pm_b_9_anexar:
            file_type_annexes = str(pm_b_9_anexar).split('.')[-1]
            file_type_annexes = file_type_annexes.lower()
            if file_type_annexes not in FILE_TYPES_ANNEXES:
                raise forms.ValidationError(
                    _("Archivo No valido '%s'  la extensión debe ser XLS, XLSX, PDF, DOC O DOCX " % pm_b_9_anexar))
        return pm_b_9_anexar 

    def clean_pm_d_1_anexos(self):
        pm_d_1_anexos = self.cleaned_data['pm_d_1_anexos']
        if pm_d_1_anexos:
            file_type_annexes = str(pm_d_1_anexos).split('.')[-1]
            file_type_annexes = file_type_annexes.lower()
            if file_type_annexes not in FILE_TYPES_ANNEXES:
                raise forms.ValidationError(
                    _("Archivo No valido '%s'  la extensión debe ser XLS, XLSX, PDF, DOC O DOCX " % pm_d_1_anexos))
        return pm_d_1_anexos

##pregunta 3 modulo b
class EditUtilizaInfEstTextForm(forms.ModelForm):

    otro_cual = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otro  ¿Cuál?', 'class': 'form-control form-control-sm mb-1', 'title': '¿Para qué se utiliza la información estadística?'}))

    class Meta:
        model = UtilizaInfEstText
        fields = ['otro_cual']

    def clean(self):

        otro_cual = self.cleaned_data.get('otro_cual')
        
        if otro_cual:
            if len(otro_cual) > 900: 
                self._errors['otro_cual'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        return(self.cleaned_data)

##pregunta 4 modulo b
class EditUsPrincInfEstTextForm(forms.ModelForm):

    orginter_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Organismos internacionales', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: ¿Cuáles podrían ser los usuarios principales de la información solicitada?'}))

    ministerios_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Ministerios', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: ¿Cuáles podrían ser los usuarios principales de la información solicitada?'}))

    orgcontrol_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Organismos de Control', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: ¿Cuáles podrían ser los usuarios principales de la información solicitada?'}))

    oentidadesordenal_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otras entidades del orden Nacional', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: ¿Cuáles podrían ser los usuarios principales de la información solicitada?'}))

    entidadesordenterr_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Entidades de orden Territorial', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: ¿Cuáles podrían ser los usuarios principales de la información solicitada?'}))

    gremios_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Gremios', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: ¿Cuáles podrían ser los usuarios principales de la información solicitada?'}))

    entiprivadas_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Entidades privadas', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: ¿Cuáles podrían ser los usuarios principales de la información solicitada?'}))

    dependenentidad_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Dependencias de la misma entidad', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: ¿Cuáles podrían ser los usuarios principales de la información solicitada?'}))

    academia_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Academia', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: ¿Cuáles podrían ser los usuarios principales de la información solicitada?'}))

    otro_cual_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otra ¿cuál?', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: ¿Cuáles podrían ser los usuarios principales de la información solicitada?'}))

    class Meta:
        model = UsuariosPrincInfEstText
        fields = ['orginter_text', 'ministerios_text', 'orgcontrol_text', 'oentidadesordenal_text',
                    'entidadesordenterr_text', 'gremios_text', 'entiprivadas_text', 'dependenentidad_text',
                    'academia_text', 'otro_cual_text']

    def clean(self):

        orginter_text = self.cleaned_data.get('orginter_text') 
        ministerios_text = self.cleaned_data.get('ministerios_text')
        orgcontrol_text = self.cleaned_data.get('orgcontrol_text')
        oentidadesordenal_text = self.cleaned_data.get('oentidadesordenal_text')
        entidadesordenterr_text = self.cleaned_data.get('entidadesordenterr_text')
        gremios_text = self.cleaned_data.get('gremios_text')
        entiprivadas_text = self.cleaned_data.get('entiprivadas_text')
        dependenentidad_text = self.cleaned_data.get('dependenentidad_text')
        academia_text = self.cleaned_data.get('academia_text')
        otro_cual_text = self.cleaned_data.get('otro_cual_text')
        
        if orginter_text:
            if len(orginter_text) > 900: 
                self._errors['orginter_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if ministerios_text:
            if len(ministerios_text) > 900: 
                self._errors['ministerios_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if orgcontrol_text:
            if len(orgcontrol_text) > 900: 
                self._errors['orgcontrol_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if oentidadesordenal_text:
            if len(oentidadesordenal_text) > 900: 
                self._errors['oentidadesordenal_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if entidadesordenterr_text:
            if len(entidadesordenterr_text) > 900: 
                self._errors['entidadesordenterr_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if gremios_text:
            if len(gremios_text) > 900: 
                self._errors['gremios_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if entiprivadas_text:
            if len(entiprivadas_text) > 900: 
                self._errors['entiprivadas_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if dependenentidad_text:
            if len(dependenentidad_text) > 900: 
                self._errors['dependenentidad_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if academia_text:
            if len(academia_text) > 900: 
                self._errors['academia_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if otro_cual_text:
            if len(otro_cual_text) > 900: 
                self._errors['otro_cual_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        return(self.cleaned_data)

##pregunta 5 modulo b
class EditNormasTextForm(forms.ModelForm):

    const_pol_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Constitución Política', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: La información solicitada responde a las siguientes normas'})) 
    ley_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Ley', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: La información solicitada responde a las siguientes normas'}))
    decreto_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Decreto (Nacional, departamental, etc.)', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: La información solicitada responde a las siguientes normas'}))
    otra_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otra (Resolución, ordenanza, acuerdo)', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: La información solicitada responde a las siguientes normas'}))

    class Meta:
        model = NormasText
        fields = ['const_pol_text', 'ley_text', 'decreto_text', 'otra_text']

    def clean(self):

        const_pol_text = self.cleaned_data.get('const_pol_text') 
        ley_text = self.cleaned_data.get('const_pol_text') 
        decreto_text = self.cleaned_data.get('const_pol_text') 
        otra_text = self.cleaned_data.get('const_pol_text') 
        
        if const_pol_text:
            if len(const_pol_text) > 900: 
                self._errors['const_pol_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if ley_text:
            if len(ley_text) > 900: 
                self._errors['ley_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if decreto_text:
            if len(decreto_text) > 900: 
                self._errors['decreto_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if otra_text:
            if len(otra_text) > 900: 
                self._errors['otra_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        return(self.cleaned_data)

##pregunta 6 modulo b
class EditInfSolRespRequerimientoTextForm(forms.ModelForm):

    planalDes_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Plan Nacional de Desarrollo (Capítulo y línea)', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: La información solicitada responde a los siguientes requerimientos'})) 
    
    cuentasecomacroec_text =  forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Cuentas económicas y macroeconómicas', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: La información solicitada responde a los siguientes requerimientos'})) 

    plansecterrcom_text =  forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Plan Sectorial, Territorial o CONPES', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: La información solicitada responde a los siguientes requerimientos'})) 

    objdessost_text =  forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Objetivos de Desarrollo Sostenible (ODS)', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: La información solicitada responde a los siguientes requerimientos'})) 

    orgcooper_text =  forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Organización para la Cooperación y el Desarrollo Económicos (OCDE)', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: La información solicitada responde a los siguientes requerimientos'})) 

    otroscomprInt_text =  forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otros compromisos internacionales', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: La información solicitada responde a los siguientes requerimientos'})) 

    otros_text =  forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otro(s)', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: La información solicitada responde a los siguientes requerimientos'})) 


    class Meta:
        model = InfSolRespRequerimientoText
        fields = ['planalDes_text', 'cuentasecomacroec_text', 'plansecterrcom_text', 'objdessost_text',
                    'orgcooper_text', 'otroscomprInt_text', 'otros_text']

    def clean(self):

        planalDes_text = self.cleaned_data.get('planalDes_text') 
        cuentasecomacroec_text = self.cleaned_data.get('cuentasecomacroec_text') 
        plansecterrcom_text = self.cleaned_data.get('plansecterrcom_text') 
        objdessost_text = self.cleaned_data.get('objdessost_text') 
        orgcooper_text = self.cleaned_data.get('orgcooper_text') 
        otroscomprInt_text = self.cleaned_data.get('otroscomprInt_text') 
        otros_text = self.cleaned_data.get('otros_text') 
        
        if planalDes_text:
            if len(planalDes_text) > 900: 
                self._errors['planalDes_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if cuentasecomacroec_text:
            if len(cuentasecomacroec_text) > 900: 
                self._errors['cuentasecomacroec_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if plansecterrcom_text:
            if len(plansecterrcom_text) > 900: 
                self._errors['plansecterrcom_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if objdessost_text:
            if len(objdessost_text) > 900: 
                self._errors['objdessost_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if orgcooper_text:
            if len(orgcooper_text) > 900: 
                self._errors['orgcooper_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])
 
        if otroscomprInt_text:
            if len(otroscomprInt_text) > 900: 
                self._errors['otroscomprInt_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        if otros_text:
            if len(otros_text) > 900: 
                self._errors['otros_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        return(self.cleaned_data)

## pregunta 8 modulo B
class EditTipoRequerimientoTextForm(forms.ModelForm):

    otros_c_text = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otro  ¿Cuál?', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: ¿Qué tipo de requerimiento es?'}))

    class Meta:
        model = TipoRequerimientoText
        fields = ['otros_c_text']

    def clean(self):

        otros_c_text = self.cleaned_data.get('otros_c_text')
        
        if otros_c_text:
            if len(otros_c_text) > 900: 
                self._errors['otros_c_text'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        return(self.cleaned_data)

###pregunta 12 modulo b
class EditDesagregacionReqTextForm(forms.ModelForm):

    otra_cual_text_a = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otro  ¿Cuál?', 'class': 'form-control form-control-sm mb-1', 'title': 'Módulo B: Indique la desagregación requerida'}))

    class Meta:
        model = DesagregacionReqText
        fields = ['otra_cual_text_a']

    def clean(self):

        otra_cual_text_a = self.cleaned_data.get('otra_cual_text_a')
        
        if otra_cual_text_a:
            if len(otra_cual_text_a) > 900: 
                self._errors['otra_cual_text_a'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        return(self.cleaned_data)

###pregunta 13 modulo b
class EditDesaReqGeogTextForm(forms.ModelForm):

    otra_cual_text_b = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otro  ¿Cuál?', 'class': 'form-control form-control-sm mb-1', 'title': 'Indique la desagregación geográfica requerida'}))

    class Meta:
        model = DesagregacionGeoReqGeograficaText
        fields = ['otra_cual_text_b']

    def clean(self):

        otra_cual_text_b = self.cleaned_data.get('otra_cual_text_b')
        
        if otra_cual_text_b:
            if len(otra_cual_text_b) > 900: 
                self._errors['otra_cual_text_b'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        return(self.cleaned_data)

###pregunta 14 modulo b
class EditPerDifusionTextForm(forms.ModelForm):

    otra_cual_text_c = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otro  ¿Cuál?', 'class': 'form-control form-control-sm mb-1', 'title': 'Periodicidad de difusión requerida'}))

    class Meta:
        model = PeriodicidadDifusionText
        fields = ['otra_cual_text_c']

    def clean(self):

        otra_cual_text_c = self.cleaned_data.get('otra_cual_text_c')
        
        if otra_cual_text_c:
            if len(otra_cual_text_c) > 900: 
                self._errors['otra_cual_text_c'] = self.error_class([ 
                    'Máximo 900 caracteres requeridos'])

        return(self.cleaned_data)

## pregunta 11
class EditSolReqInformacionForm(forms.ModelForm):  ##EdittestPreguntaOnceForm

    active = forms.BooleanField(required=False, initial=False, widget=forms.CheckboxInput(
        attrs={"id": "id_active", "data-toggle": "toggle", "data-size": "small", " data-on": "Si",  "data-off": "No", "data-onstyle": "success", "data-offstyle": "danger"}))
    
    ooee_pmb11 = forms.ModelChoiceField(required=False, label="Operaciones estadísticas",  queryset=ooee_service.objects.all(),
            widget=forms.Select(attrs={'class': 'form-control form-control-sm',}))

    opc_apr_oe = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple,
        queryset=OpcAprovooee.objects.all())
    
    rraa_pmb11 = forms.ModelChoiceField(required=False, label="Registros administrativos",  queryset=rraa_service.objects.all(),
            widget=forms.Select(attrs={'class': 'form-control form-control-sm',}))

    opc_apr_ra = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple,
        queryset=OpcAprovrraa.objects.all())

    genera_nuev = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple,
        queryset=OpcAprovGenNueva.objects.all())

    inc_var_cual = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'a)2.1. Incluir variables/preguntas ¿Cuál(es)?', 'class': 'form-control form-control-sm mb-1', 
        'title': 'Módulo B: pregunta 11 a)2.1. Incluir variables/preguntas ¿Cuál(es)?'}))
    
    cam_preg_cual = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'a)2.2. Cambiar la formulación de alguna(s) pregunta(s) ¿Cuál(es)?', 'class': 'form-control form-control-sm mb-1', 
        'title': 'Módulo B: pregunta 11 a)2.2. Cambiar la formulación de alguna(s) pregunta(s) ¿Cuál(es)?'}))

    am_des_tem_cual =  forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'a)2.3. Ampliar la desagregación temática ¿Cuál(es)?', 'class': 'form-control form-control-sm mb-1', 
        'title': 'Módulo B: pregunta 11 a)2.3. Ampliar la desagregación temática ¿Cuál(es)?'}))

    am_des_geo_cual = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'a)2.4. Ampliar la desagregación geográfica ¿Cuál(es)?', 'class': 'form-control form-control-sm mb-1', 
        'title': 'Módulo B: pregunta 11 a)2.4. Ampliar la desagregación geográfica ¿Cuál(es)?'}))

    dif_resul_cual = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'a)2.5. Aumentar la frecuencia en la difusión de los resultados ¿Cuál(es)?', 'class': 'form-control form-control-sm mb-1', 
        'title': 'Módulo B: pregunta 11 a)2.5. Aumentar la frecuencia en la difusión de los resultados ¿Cuál(es)?'}))

    opc_aprov_cual = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'a)2.6. Otra ¿Cuál(es)?', 'class': 'form-control form-control-sm mb-1', 
        'title': 'Módulo B: a)2.6. Otra ¿Cuál(es)?'}))

    inc_varia_cual = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'b)2.1.  Inclusión de variables/preguntas', 'class': 'form-control form-control-sm mb-1', 
        'title': 'Módulo B: pregunta 11 Inclusión de variables/preguntas'}))
    
    camb_pregu_cual = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'b)2.2. Cambios en la formulación de alguna(s) pregunta(s)', 'class': 'form-control form-control-sm mb-1', 
        'title': 'Módulo B: Pregunta 11 b)2.2. Cambios en la formulación de alguna(s) pregunta(s)'}))

    otros_aprov_ra = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'b)2.3. Otro ', 'class': 'form-control form-control-sm mb-1', 
        'title': 'Módulo B: pregunta 11 b)2.3. Otro '}))

    nueva_oe = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'c)1. Operación estadística nueva', 'class': 'form-control form-control-sm mb-1', 
        'title': 'Módulo B: pregunta 11 c)1. Operación estadística nueva'}))
    
    indi_cual = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'c)2. Indicador ¿Cuál(es)?', 'class': 'form-control form-control-sm mb-1', 
        'title': 'Módulo B: pregunta 11 c)2. Indicador ¿Cuál(es)?'}))

    gen_nueva = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'c)3. Otra  ¿Cuál?', 'class': 'form-control form-control-sm mb-1', 
        'title': 'Módulo B: pregunta 11 c)3. Otra  ¿Cuál?'}))

    class Meta:
        model = SolReqInformacion
        fields = ['active', 'ooee_pmb11','opc_apr_oe','rraa_pmb11','opc_apr_ra','genera_nuev',
                    'inc_var_cual','cam_preg_cual','am_des_tem_cual', 'am_des_geo_cual','dif_resul_cual',
                    'opc_aprov_cual','inc_varia_cual','camb_pregu_cual','otros_aprov_ra','nueva_oe','indi_cual','gen_nueva',
                    ]

    def clean(self):

        inc_var_cual = self.cleaned_data.get('inc_var_cual')
        cam_preg_cual = self.cleaned_data.get('cam_preg_cual')
        am_des_tem_cual = self.cleaned_data.get('am_des_tem_cual')
        am_des_geo_cual = self.cleaned_data.get('am_des_geo_cual')
        dif_resul_cual = self.cleaned_data.get('dif_resul_cual')
        opc_aprov_cual = self.cleaned_data.get('opc_aprov_cual')

        inc_varia_cual = self.cleaned_data.get('inc_varia_cual')
        camb_pregu_cual = self.cleaned_data.get('camb_pregu_cual')
        otros_aprov_ra = self.cleaned_data.get('otros_aprov_ra')

        nueva_oe = self.cleaned_data.get('nueva_oe')
        indi_cual = self.cleaned_data.get('indi_cual')
        gen_nueva = self.cleaned_data.get('gen_nueva')

        
        if inc_var_cual:
            if len(inc_var_cual) > 800: 
                self._errors['inc_var_cual'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        if cam_preg_cual:
            if len(cam_preg_cual) > 800: 
                self._errors['cam_preg_cual'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        if am_des_tem_cual:
            if len(am_des_tem_cual) > 800: 
                self._errors['am_des_tem_cual'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        if am_des_geo_cual:
            if len(am_des_geo_cual) > 800: 
                self._errors['am_des_geo_cual'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        if dif_resul_cual:
            if len(dif_resul_cual) > 800: 
                self._errors['dif_resul_cual'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        if opc_aprov_cual:
            if len(opc_aprov_cual) > 800: 
                self._errors['opc_aprov_cual'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])
        
        if inc_varia_cual:
            if len(inc_varia_cual) > 800: 
                self._errors['inc_varia_cual'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        if camb_pregu_cual:
            if len(camb_pregu_cual) > 800: 
                self._errors['camb_pregu_cual'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        if otros_aprov_ra:
            if len(otros_aprov_ra) > 800: 
                self._errors['otros_aprov_ra'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])
        
        if nueva_oe:
            if len(nueva_oe) > 800: 
                self._errors['nueva_oe'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        if indi_cual:
            if len(indi_cual) > 800: 
                self._errors['indi_cual'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        if gen_nueva:
            if len(gen_nueva) > 800: 
                self._errors['gen_nueva'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])
        
        return(self.cleaned_data)

### pregunta 11 

class CrearConsumidorInfoForm(forms.ModelForm):

    nombre_ec = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Nombre entidad consumidora de información', 'class': 'form-control form-control-sm mb-1'})) 
    nit_ec = forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Nit', 'class': 'form-control'}))
    tipo_entidad_ec = forms.ModelChoiceField(label="Tipo entidad", empty_label=None, queryset=TipoEntidadddi.objects.all(), widget=forms.Select(attrs={'class':'form-control'})) 
    direccion_ec = forms.CharField(required=False, widget=forms.TextInput(attrs={'class': 'form-control'}))
    telefono_ec = forms.IntegerField(required=False, widget=forms.NumberInput(attrs={'placeholder': '999-9999', 'class': 'form-control col-9'}))  
    pagina_web_ec = forms.URLField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Pagina Web','class': 'form-control'}))
    estado =  forms.ModelChoiceField(empty_label=None, queryset=ddi_estado.objects.all(), widget=forms.Select(attrs={'class':'form-control'}))

    ##Responsable de la solicitud
    nombre_resp = forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Nombre','class': 'form-control col-8'}))
    cargo_resp = forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Cargo', 'class': 'form-control col-4'}))
    correo_resp = forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Correo electrónico','class': 'form-control'}))
    telefono_resp = forms.IntegerField(required=False, widget=forms.NumberInput(attrs={'placeholder': '999-9999','class': 'form-control col-9'}))
    extension_resp = forms.IntegerField(required=False, widget=forms.NumberInput(attrs={'placeholder': 'Ext.','class': 'form-control col-3'}))

    
    class Meta:
        model = consumidores_info
        fields = ['nombre_ec', 'nit_ec', 'tipo_entidad_ec', 'tipo_entidad_ec', 'direccion_ec', 'telefono_ec', 'pagina_web_ec',
                    'estado', 'nombre_resp', 'cargo_resp', 'correo_resp',  'telefono_resp', 'extension_resp']

    def clean(self):

        nombre_ec = self.cleaned_data.get('nombre_ec')
        nombre_resp = self.cleaned_data.get('nombre_resp')
        cargo_resp = self.cleaned_data.get('cargo_resp')

        if nombre_ec:
            if len(nombre_ec) > 200: 
                self._errors['nombre_ec'] = self.error_class([ 
                    'Máximo 200 caracteres requeridos'])

        if nombre_resp:
            if len(nombre_resp) > 200: 
                self._errors['nombre_resp'] = self.error_class([ 
                    'Máximo 200 caracteres requeridos'])

        if cargo_resp:
            if len(cargo_resp) > 500: 
                self._errors['cargo_resp'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])
        
        return(self.cleaned_data)

    def clean_nit_ec(self):
        nit_ec = self.cleaned_data['nit_ec']
        if nit_ec:
            if len(nit_ec) >  12:
                raise forms.ValidationError(_("El nit agregado '%s' debe contener 12 caracteres." % nit_ec))
        return nit_ec 


    def clean_telefono_ec(self):
        telefono_ec = self.cleaned_data['telefono_ec']
        if telefono_ec:
            if len(str(telefono_ec)) < 7 or len(str(telefono_ec)) > 10 :
                raise forms.ValidationError(_("el número telefónico  que ingreso '%s' debe contener 10 digitos para celular o 7 para teléfono fijo. " % telefono_ec))
                
        return telefono_ec


    def clean_pagina_web_ec(self):
        pagina_web_ec = self.cleaned_data['pagina_web_ec']
        validate = URLValidator()
        if pagina_web_ec:
            try: 
                validate(pagina_web_ec)
                print("pagina valida")
                return pagina_web_ec
            except ValidationError:
                print("invalida")
               

    ##responsable Fields

    def clean_correo_resp(self):
        correo_resp = self.cleaned_data['correo_resp']
        if correo_resp:
            if "@" not in correo_resp:
                raise forms.ValidationError(
                    _("El correo electronico que ingreso '%s' debe contener @." % correo_resp))

        return correo_resp

    def clean_telefono_resp(self):
        telefono_resp = self.cleaned_data['telefono_resp']
        if telefono_resp:
            if len(str(telefono_resp)) < 7 or len(str(telefono_resp)) > 10 :
                raise forms.ValidationError(_("el número telefónico  que ingreso '%s' debe contener 10 digitos para celular o 7 para teléfono fijo. " % telefono_resp))
        return telefono_resp

    
    def clean_extension_resp(self):
        extension_resp = self.cleaned_data['extension_resp']
        if extension_resp:
            if len(str(extension_resp)) > 6:
                raise forms.ValidationError(_("La extensión ingresada '%s' no debe  debe contener más de 6 digitos." % extension_resp))
        return extension_resp


class EditarConsumidorInfoForm(forms.ModelForm):

    nombre_ec = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Nombre entidad consumidora de información', 'class': 'form-control form-control-sm mb-1'})) 
    nit_ec = forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Nit', 'class': 'form-control'}))
    tipo_entidad_ec = forms.ModelChoiceField(label="Tipo entidad", empty_label=None, queryset=TipoEntidadddi.objects.all(), widget=forms.Select(attrs={'class':'form-control'})) 
    direccion_ec = forms.CharField(required=False, widget=forms.TextInput(attrs={'class': 'form-control'}))
    telefono_ec = forms.IntegerField(required=False, widget=forms.NumberInput(attrs={'placeholder': '999-9999', 'class': 'form-control col-9'}))  
    pagina_web_ec = forms.URLField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Pagina Web','class': 'form-control'}))
    estado =  forms.ModelChoiceField(empty_label=None, queryset=ddi_estado.objects.all(), widget=forms.Select(attrs={'class':'form-control'}))

    ##Responsable de la solicitud
    nombre_resp = forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Nombre','class': 'form-control col-8'}))
    cargo_resp = forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Cargo', 'class': 'form-control col-4'}))
    correo_resp = forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Correo electrónico','class': 'form-control'}))
    telefono_resp = forms.IntegerField(required=False, widget=forms.NumberInput(attrs={'placeholder': '999-9999','class': 'form-control col-9'}))
    extension_resp = forms.IntegerField(required=False, widget=forms.NumberInput(attrs={'placeholder': 'Ext.','class': 'form-control col-3'}))

    
    class Meta:
        model = consumidores_info
        fields = ['nombre_ec', 'nit_ec', 'tipo_entidad_ec', 'tipo_entidad_ec', 'direccion_ec', 'telefono_ec', 'pagina_web_ec',
                    'estado', 'nombre_resp', 'cargo_resp', 'correo_resp',  'telefono_resp', 'extension_resp']

    def clean(self):

        nombre_ec = self.cleaned_data.get('nombre_ec')
        nombre_resp = self.cleaned_data.get('nombre_resp')
        cargo_resp = self.cleaned_data.get('cargo_resp')

        if nombre_ec:
            if len(nombre_ec) > 200: 
                self._errors['nombre_ec'] = self.error_class([ 
                    'Máximo 200 caracteres requeridos'])

        if nombre_resp:
            if len(nombre_resp) > 200: 
                self._errors['nombre_resp'] = self.error_class([ 
                    'Máximo 200 caracteres requeridos'])

        if cargo_resp:
            if len(cargo_resp) > 500: 
                self._errors['cargo_resp'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])
               
        return(self.cleaned_data)

    def clean_nit_ec(self):
        nit_ec = self.cleaned_data['nit_ec']
        if nit_ec:
            if len(nit_ec) >  12:
                raise forms.ValidationError(_("El nit agregado '%s' debe contener 12 caracteres." % nit_ec))
        return nit_ec 


    def clean_telefono_ec(self):
        telefono_ec = self.cleaned_data['telefono_ec']
        if telefono_ec:
            if len(str(telefono_ec)) < 7 or len(str(telefono_ec)) > 10 :
                raise forms.ValidationError(_("el número telefónico  que ingreso '%s' debe contener 10 digitos para celular o 7 para teléfono fijo. " % telefono_ec))
                
        return telefono_ec


    def clean_pagina_web_ec(self):
        pagina_web_ec = self.cleaned_data['pagina_web_ec']
        validate = URLValidator()
        if pagina_web_ec:
            try: 
                validate(pagina_web_ec)
                print("pagina valida")
                return pagina_web_ec
            except ValidationError:
                print("invalida")
               

    ##responsable Fields

    def clean_correo_resp(self):
        correo_resp = self.cleaned_data['correo_resp']
        if correo_resp:
            if "@" not in correo_resp:
                raise forms.ValidationError(
                    _("El correo electronico que ingreso '%s' debe contener @." % correo_resp))

        return correo_resp

    def clean_telefono_resp(self):
        telefono_resp = self.cleaned_data['telefono_resp']
        if telefono_resp:
            if len(str(telefono_resp)) < 7 or len(str(telefono_resp)) > 10 :
                raise forms.ValidationError(_("el número telefónico  que ingreso '%s' debe contener 10 digitos para celular o 7 para teléfono fijo. " % telefono_resp))
        return telefono_resp

    
    def clean_extension_resp(self):
        extension_resp = self.cleaned_data['extension_resp']
        if extension_resp:
            if len(str(extension_resp)) > 6:
                raise forms.ValidationError(_("La extensión ingresada '%s' no debe  debe contener más de 6 digitos." % extension_resp))
        return extension_resp



