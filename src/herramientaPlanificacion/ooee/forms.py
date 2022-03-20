from django import forms
from django.core import validators
from django.core.validators import validate_integer
from django.core.exceptions import ValidationError
from .models import OoeeState, AreaTematica, Tema, TemaCompartido, SalaEspecializada, FasesProceso, MB_EntidadFases, Norma, MB_Norma, Requerimientos, MB_Requerimientos, \
    PrinUsuarios, MB_PrinUsuarios, OperacionEstadistica, OoeeLog, Comment, UnidadObservacion, MC_UnidadObservacion, TipoOperacion, \
    ObtencionDato, MC_ObtencionDato, MuestreoProbabilistico, MC_MuestreoProbabilistico, MuestreoNoProbabilistico, MC_MuestreoNoProbabilistico, \
    TipoMarco, MC_TipoMarco, DocsDesarrollo, MC_DocsDesarrollo, ConceptosEstandarizados, MC_ConceptosEstandarizados, Clasificaciones, MC_Clasificaciones, \
    CoberturaGeografica, MC_CoberturaGeografica, DesagregacionInformacion, MC_DesagregacionInformacion, DesagregacionGeografica, DesagregacionZona, \
    DesagregacionGrupos, FuenteFinanciacion, MC_FuenteFinanciacion, MC_listaVariable, MedioDatos, MD_MedioDatos, PeriodicidadOe, MD_PeriodicidadOe, \
    HerramProcesamiento, MD_HerramProcesamiento,  AnalisisResultados, ME_AnalisisResultados, MediosDifusion, MF_MediosDifusion, FechaPublicacion, \
    MF_FechaPublicacion, FrecuenciaDifusion, MF_FrecuenciaDifusion, ProductosDifundir, MF_ProductosDifundir, OtrosProductos, MF_OtrosProductos, MF_ResultadosSimilares, \
    MF_HPSistemaInfo, MC_ResultadoEstadistico, TipoNovedad, NovedadActualizacion, EstadoActualizacion, EstadoCritica,  Critica, EstadoEvaluacion, ResultadoEvaluacion, \
    PlanDeMejoramiento, ResultadoEvalmatriz, SeguimientoAnual, EvaluacionCalidad, ListaDeOOEE

from entities.models import Entidades_oe
from rraa.models import RegistroAdministrativo
from django.utils.translation import gettext as _
# url validators
from django.core.validators import URLValidator
from django.forms import modelformset_factory

# forms by Created ooee

CHOICES_STATEOOOE = (
    ('activa', 'Activa'),
    ('inactiva', 'Inactiva'),
)

CHOICES_VALIDATIONOOOE = (
    ('oficial', 'Oficial'),
    ('no oficial', 'No Oficial'),
)

FILE_TYPES = ['xls', 'xlsx']
FILE_TYPES_ANNEXES = ['xls', 'xlsx', 'pdf', 'doc', 'docx']


class CreateOEForm(forms.ModelForm):

    codigo_oe = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Código', 'class': 'form-control form-control-sm'}))
    entidad = forms.ModelChoiceField(label="Nombre de la entidad", empty_label=None, queryset=Entidades_oe.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))
    area_tematica = forms.ModelChoiceField(label="Area Temática", empty_label=None, queryset=AreaTematica.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))
    tema = forms.ModelChoiceField(label="Tema", empty_label=None, queryset=Tema.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))
    tema_compartido = forms.ModelMultipleChoiceField(
        required=False,  queryset=TemaCompartido.objects.all(), widget=forms.CheckboxSelectMultiple)
    entidad_resp2 = forms.ModelChoiceField(required=False, label="Nombre de la entidad",  queryset=Entidades_oe.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))
    entidad_resp3 = forms.ModelChoiceField(required=False, label="Nombre de la entidad",  queryset=Entidades_oe.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))
    nombre_dep = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Dependencia', 'class': 'form-control form-control-sm'}))
    nombre_dir = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Director', 'class': 'form-control form-control-sm'}))
    cargo_dir = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Cargo', 'class': 'form-control form-control-sm'}))
    correo_dir = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Correo electrónico', 'class': 'form-control form-control-sm'}))
    tel_dir = forms.IntegerField(required=False, widget=forms.NumberInput(
        attrs={'placeholder': '999-9999', 'class': 'form-control form-control-sm'}))
    nombre_resp = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Responsable', 'class': 'form-control form-control-sm'}))
    cargo_resp = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Cargo', 'class': 'form-control form-control-sm'}))
    correo_resp = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Correo electrónico', 'class': 'form-control form-control-sm'}))
    tel_resp = forms.IntegerField(required=False, widget=forms.NumberInput(
        attrs={'placeholder': '999-9999', 'class': 'form-control form-control-sm'}))
    nombre_oe = forms.CharField(widget=forms.TextInput(
        attrs={'placeholder': 'Nombre de la operación estadística', 'class': 'form-control form-control-sm'}))
    objetivo_oe = forms.CharField(widget=forms.Textarea(
        attrs={'placeholder': 'Objetivo de la operación estadística', 'class': 'form-control form-control-sm'}))

    nombre_est = forms.ModelChoiceField(label="Estado",  queryset=OoeeState.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm', 'onchange':'fondo(this)'}))

    fase = forms.BooleanField(required=False, initial=False, widget=forms.CheckboxInput(
        attrs={"id": "id_fase", "data-toggle": "toggle", "data-size": "small", " data-on": "Si",  "data-off": "No", "data-onstyle": "success", "data-offstyle": "danger"}))

    norma = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=Norma.objects.all())
    requerimientos = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=Requerimientos.objects.all())
    pri_usuarios = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=PrinUsuarios.objects.all())
    pob_obje = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': '¿Cuál es la población objetivo de la operación estadística?', 'class': 'form-control form-control-sm'}))
    uni_observacion = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=UnidadObservacion.objects.all())
    tipo_operacion = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(
        attrs={'class': 'list-group list-group-horizontal-sm'}), queryset=TipoOperacion.objects.all())
    obt_dato = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=ObtencionDato.objects.all())
    
    rraa_lista = forms.ModelMultipleChoiceField(
        required=False, widget=forms.SelectMultiple(attrs={"class":"selectpicker", "data-live-search":"true", "title": "Módulo C Pregunta 4: Donde se obtienen los datos RRAA"}), queryset=RegistroAdministrativo.objects.all())
    
    ooee_lista = forms.ModelMultipleChoiceField(
        required=False, widget=forms.SelectMultiple(attrs={"class":"selectpicker", "data-live-search":"true", "title": "Módulo C Pregunta 4: Donde se obtienen los datos OOEE"}), queryset=ListaDeOOEE.objects.all())

    tipo_probabilistico = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=MuestreoProbabilistico.objects.all())
    tipo_no_probabilistico = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=MuestreoNoProbabilistico.objects.all())
    marco_estad = forms.BooleanField(required=False, initial=False,  widget=forms.CheckboxInput(
        attrs={"id": "id_marco", "data-toggle": "toggle", "data-size": "small", " data-on": "Si",  "data-off": "No", "data-onstyle": "success", "data-offstyle": "danger"}))
    tipo_marco = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=TipoMarco.objects.all())
    docs_des = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=DocsDesarrollo.objects.all())
    lista_conc = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=ConceptosEstandarizados.objects.all())
    nome_clas = forms.BooleanField(required=False, initial=False, widget=forms.CheckboxInput(
        attrs={"id": "id_nome_clas", "data-toggle": "toggle", "data-size": "small", " data-on": "Si",  "data-off": "No", "data-onstyle": "success", "data-offstyle": "danger"}))
    nombre_cla = forms.ModelMultipleChoiceField(
        required=False,  queryset=Clasificaciones.objects.all(), widget=forms.CheckboxSelectMultiple)
    cob_geo = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(
        attrs={'class': 'list-group'}), queryset=CoberturaGeografica.objects.all())
    opc_desag = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(
        attrs={'class': 'list-group list-group-horizontal-sm'}), queryset=DesagregacionInformacion.objects.all())
    des_geo = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=DesagregacionGeografica.objects.all())
    des_zona = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(
        attrs={'class': 'list-group list-group-horizontal-sm'}), queryset=DesagregacionZona.objects.all())
    des_grupo = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=DesagregacionGrupos.objects.all())
    ca_anual = forms.IntegerField(required=False, widget=forms.NumberInput(
        attrs={'placeholder': '$ Costo anual', 'class': 'form-control form-control-sm'}))
    cb_anual = forms.BooleanField(
        required=False, initial=True, label='No sabe')
    fuentes = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=FuenteFinanciacion.objects.all())
    variable_file = forms.FileField(
        label='Select a file',
        help_text='max. 42 megabytes',
        required=False,
    )
    med_obt = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=MedioDatos.objects.all())
    periodicidad = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=PeriodicidadOe.objects.all())
    h_proc = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=HerramProcesamiento.objects.all())
    descrip_proces = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Haga una breve descripción de la manera cómo se realiza el procesamiento de los datos', 'class': 'form-control form-control-sm'}))
    a_resul = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=AnalisisResultados.objects.all())
    m_dif = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=MediosDifusion.objects.all())
    res_est_url = forms.URLField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Pagina Web', 'class': 'form-control form-control-sm'}))
    dispo_desde = forms.DateField(
        required=False,
        widget=forms.DateInput(format='%Y-%m-%d',
                               attrs={'class': 'form-control datepicker'}),
        input_formats=('%Y-%m-%d', )
        )
    dispo_hasta = forms.DateField(
        required=False,
        widget=forms.DateInput(format='%Y-%m-%d',
                               attrs={'class': 'form-control datepicker'}),
        input_formats=('%Y-%m-%d', )
        )
    f_publi = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=FechaPublicacion.objects.all())
    fre_dif = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=FrecuenciaDifusion.objects.all())
    pro_dif = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=ProductosDifundir.objects.all())
    otro_prod = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=OtrosProductos.objects.all())
    conoce_otra = forms.BooleanField(required=False, initial=False, widget=forms.CheckboxInput(
        attrs={"id": "id_conoce_otra", "data-toggle": "toggle", "data-size": "small", " data-on": "Si",  "data-off": "No", "data-onstyle": "success", "data-offstyle": "danger"}))
    hp_siste_infor = forms.BooleanField(required=False, initial=False, widget=forms.CheckboxInput(
        attrs={"id": "id_hp_siste_infor", "data-toggle": "toggle", "data-size": "small", " data-on": "Si",  "data-off": "No", "data-onstyle": "success", "data-offstyle": "danger"}))
    observaciones = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': '(Indique el número de pregunta si es necesario ampliar, aclarar o complementar la respuesta)', 'class': 'form-control form-control-sm'}))

    anexos = forms.FileField(
        label='Select a file',
        help_text='max. 42 megabytes',
        required=False,
    )
    estado_oe_tematico = forms.ChoiceField(required=False, label="Estado de la operación estadística", choices=CHOICES_STATEOOOE, widget=forms.Select(
        attrs={'class': 'form-control form-control-sm'}))
    validacion_oe_tematico = forms.ChoiceField(required=False, label="Validación de la operación estadística",  choices=CHOICES_VALIDATIONOOOE, widget=forms.Select(
        attrs={'class': 'form-control form-control-sm'}))

    class Meta:
        model = OperacionEstadistica
        fields = [
            'codigo_oe',
            'entidad',
            'area_tematica',
            'tema',
            'tema_compartido',
            'entidad_resp2',
            'entidad_resp3',
            'nombre_dep',
            'nombre_dir',
            'cargo_dir',
            'correo_dir',
            'tel_dir',
            'nombre_resp',
            'cargo_resp',
            'correo_resp',
            'tel_resp',
            'nombre_oe',
            'objetivo_oe',
            'nombre_est',
            'fase',
            'norma',
            'requerimientos',
            'pri_usuarios',
            'pob_obje',
            'uni_observacion',
            'tipo_operacion',
            'obt_dato',
            'rraa_lista',
            'ooee_lista',
            'tipo_probabilistico',
            'tipo_no_probabilistico',
            'marco_estad',
            'tipo_marco',
            'docs_des',
            'lista_conc',
            'nome_clas',
            'nombre_cla',
            'cob_geo',
            'opc_desag',
            'des_geo',
            'des_zona',
            'des_grupo',
            'ca_anual',
            'cb_anual',
            'fuentes',
            'variable_file',
            'med_obt',
            'periodicidad',
            'h_proc',
            'descrip_proces',
            'a_resul',
            'm_dif',
            'res_est_url',
            'dispo_desde',
            'dispo_hasta',
            'f_publi',
            'fre_dif',
            'pro_dif',
            'otro_prod',
            'conoce_otra',
            'hp_siste_infor',
            'observaciones',
            'anexos',
            'estado_oe_tematico',
            'validacion_oe_tematico'

        ]

    def clean(self):

        nombre_dep = self.cleaned_data.get('nombre_dep') 
        nombre_dir = self.cleaned_data.get('nombre_dir') 
        cargo_dir = self.cleaned_data.get('cargo_dir')
        nombre_resp = self.cleaned_data.get('nombre_resp')
        cargo_resp = self.cleaned_data.get('cargo_resp')
        objetivo_oe = self.cleaned_data.get('objetivo_oe')
        pob_obje  = self.cleaned_data.get('pob_obje')
        descrip_proces = self.cleaned_data.get('descrip_proces')
        observaciones = self.cleaned_data.get('observaciones')

        if nombre_dep:
            if len(nombre_dep) > 80: 
                self._errors['nombre_dep'] = self.error_class([ 
                    'Máximo 80 caracteres requeridos'])
        
        if nombre_dir:
            if len(nombre_dir) > 80: 
                self._errors['nombre_dir'] = self.error_class([ 
                    'Máximo 80 caracteres requeridos'])

        if cargo_dir:
            if len(cargo_dir) > 80: 
                self._errors['cargo_dir'] = self.error_class([ 
                    'Máximo 80 caracteres requeridos'])

        if nombre_resp:
            if len(nombre_resp) > 80: 
                self._errors['nombre_resp'] = self.error_class([ 
                    'Máximo 80 caracteres requeridos'])

        if cargo_resp:
            if len(cargo_resp) > 80: 
                self._errors['cargo_resp'] = self.error_class([ 
                    'Máximo 80 caracteres requeridos'])

        if objetivo_oe:
            if len(objetivo_oe) > 800: 
                self._errors['objetivo_oe'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos']) 

        if pob_obje:
            if len(pob_obje) > 2000: 
                self._errors['pob_obje'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos']) 
        
        if descrip_proces:
            if len(descrip_proces) > 3000: 
                self._errors['descrip_proces'] = self.error_class([ 
                    'Máximo 3000 caracteres requeridos'])

        if observaciones:
            if len(observaciones) > 2000: 
                self._errors['observaciones'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos']) 

        # return any errors if found 
        return(self.cleaned_data)


    def clean_nombre_oe(self):
        nombre_oe = self.cleaned_data['nombre_oe']
        if OperacionEstadistica.objects.filter(nombre_oe=nombre_oe).exists():
            raise forms.ValidationError(_("La operación estadística  '%s' ya existe." % nombre_oe))
    
        if nombre_oe:
            if len(nombre_oe) > 600: 
                self._errors['nombre_oe'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])
        
        return nombre_oe


    def clean_tel_dir(self):
        tel_dir = self.cleaned_data['tel_dir']
        if tel_dir:
            if len(str(tel_dir)) < 7 or len(str(tel_dir)) > 10:
                raise forms.ValidationError(
                    _("el número telefónico  que ingreso '%s' debe contener 10 digitos para celular o 7 para teléfono fijo. " % tel_dir))
        return tel_dir

    def clean_tel_resp(self):
        tel_resp = self.cleaned_data['tel_resp']
        if tel_resp:
            if len(str(tel_resp)) < 7 or len(str(tel_resp)) > 10:
                raise forms.ValidationError(
                    _("el número telefónico  que ingreso '%s' debe contener 10 digitos para celular o 7 para teléfono fijo. " % tel_resp))
        return tel_resp

    def clean_correo_dir(self):
        correo_dir = self.cleaned_data['correo_dir']
        if correo_dir:
            if "@" not in correo_dir:
                raise forms.ValidationError(
                    _("El correo electrónico que ingreso '%s' debe contener @." % correo_dir))
            if  len(correo_dir) > 80:
                raise forms.ValidationError(
                    _("El correo electrónico que ingreso '%s' debe contener menos de 80 caracteres" % correo_dir))
        return correo_dir

    def clean_correo_resp(self):
        correo_resp = self.cleaned_data['correo_resp']
        if correo_resp:
            if "@" not in correo_resp:
                raise forms.ValidationError(
                    _("El correo electrónico que ingreso '%s' debe contener @." % correo_resp))
            if  len(correo_resp) > 80:
                raise forms.ValidationError(
                    _("El correo electrónico que ingreso '%s' debe contener menos de 80 caracteres" % correo_resp))
        return correo_resp

    def clean_variable_file(self):
        variable_file = self.cleaned_data['variable_file']
        if variable_file:
            print("formulario", variable_file)
            file_type = str(variable_file).split('.')[-1]
            file_type = file_type.lower()
            if file_type not in FILE_TYPES:
                raise forms.ValidationError(
                    _("Archivo No valido '%s'  la extensión debe ser XLS o XLSX" % variable_file))

        return variable_file

    def clean_res_est_url(self):
        res_est_url = self.cleaned_data['res_est_url']
        if res_est_url:
            validate = URLValidator()
            try:
                validate(res_est_url)
                return res_est_url
            except ValidationError:
                return None 

                

    def clean_anexos(self):
        anexos = self.cleaned_data['anexos']
        if anexos:
            file_type_annexes = str(anexos).split('.')[-1]
            file_type_annexes = file_type_annexes.lower()
            if file_type_annexes not in FILE_TYPES_ANNEXES:
                raise forms.ValidationError(
                    _("Archivo No valido '%s'  la extensión debe ser XLS, XLSX, PDF, DOC O DOCX " % anexos))
        return anexos


# pregunta 14 del  modulo C

class MC_listaVariableForm(forms.ModelForm):

    lista_var = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Ingrese las variables que maneja la operación', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MC_listaVariable
        fields = [
            'lista_var'
        ]

    def clean(self):

        lista_var = self.cleaned_data.get('lista_var')
        
        if lista_var:
            if len(lista_var) > 250: 
                self._errors['lista_var'] = self.error_class([ 
                    'Máximo 250 caracteres requeridos'])

         # return any errors if found 
        return(self.cleaned_data)


MC_listaVariableFormset = modelformset_factory(
    MC_listaVariable,
    form=MC_listaVariableForm,
    can_delete=True,
    extra=1,
    fields=('lista_var',),
)

# Modulo C pregunta 15


class MC_ResultadoEstadisticoForm(forms.ModelForm):
    resultEstad = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'indicadores o resultados agregados', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MC_ResultadoEstadistico
        fields = [
            'resultEstad'
        ]

    def clean(self):

        resultEstad = self.cleaned_data.get('resultEstad')
        
        if resultEstad:
            if len(resultEstad) > 250: 
                self._errors['resultEstad'] = self.error_class([ 
                    'Máximo 250 caracteres requeridos'])

         # return any errors if found 
        return(self.cleaned_data)


# pregunta 15 del  modulo C
MC_resultadoEstadisticoFormset = modelformset_factory(
    MC_ResultadoEstadistico,
    form=MC_ResultadoEstadisticoForm,
    fields=('resultEstad',),
    can_delete=True,
    extra=1,
)

# modulo B pregunta  3


class CreateEntidadFaseForm(forms.ModelForm):
    # nombre_entifas = forms.ModelChoiceField(label="Nombre de la entidad", empty_label=None, queryset=MB_EntidadFases.objects.all(
    # ), widget=forms.Select(attrs={'id': 'entidad-fase', 'class': 'form-control form-control-sm'}))
    nombre_entifas = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'escriba el nombre de la entidad', 'class': 'form-control form-control-sm'}))
    fases = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(
        attrs={'class': 'list-group list-group-horizontal-sm selectFase'}), queryset=FasesProceso.objects.all())

    class Meta:
        model = MB_EntidadFases
        fields = [
            'nombre_entifas', 'fases'
        ]

    def clean(self):

        nombre_entifas = self.cleaned_data.get('nombre_entifas')
        
        if nombre_entifas:
            if len(nombre_entifas) > 300: 
                self._errors['nombre_entifas'] = self.error_class([ 
                    'Máximo 300 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)


MB_EntidadFasesFormset = modelformset_factory(
    MB_EntidadFases,
    form=CreateEntidadFaseForm,
    fields=('nombre_entifas', 'fases',),
    extra=1,
    can_delete=True,
)

# modulo B pregunta 4


class MB_NormaForm(forms.ModelForm):

    cp_d = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': 'Describa el requerimiento', 'class': 'form-control form-control-sm mb-1'}))
    ley_d = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Describa el requerimiento', 'class': 'form-control form-control-sm mb-1'}))
    decreto_d = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Describa el requerimiento', 'class': 'form-control form-control-sm mb-1'}))
    otra_d = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Describa el requerimiento', 'class': 'form-control form-control-sm mb-1'}))

    class Meta:
        model = MB_Norma
        fields = ['cp_d', 'ley_d', 'decreto_d',  'otra_d']

    def clean(self):

        cp_d = self.cleaned_data.get('cp_d') 
        ley_d = self.cleaned_data.get('ley_d') 
        decreto_d = self.cleaned_data.get('decreto_d')
        otra_d = self.cleaned_data.get('otra_d')
        
        if cp_d:
            if len(cp_d) > 2000: 
                self._errors['cp_d'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos']) 
        
        if ley_d:
            if len(ley_d) > 2000: 
                self._errors['ley_d'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos']) 
        
        if decreto_d:
            if len(decreto_d) > 2000: 
                self._errors['decreto_d'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos']) 

        if otra_d:
            if len(otra_d) > 2000: 
                self._errors['otra_d'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos']) 

        # return any errors if found 
        return(self.cleaned_data)


# modulo B pregunta 5

class MB_RequerimientoForm(forms.ModelForm):

    ri_ods = forms.CharField(required=False, widget=forms.TextInput(attrs={
                             'placeholder': 'Describa el requerimiento especifico', 'class': 'form-control form-control-sm'}))
    ri_ocde = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Describa el requerimiento especifico', 'class': 'form-control form-control-sm'}))
    ri_ci = forms.CharField(required=False, widget=forms.TextInput(attrs={
                            'placeholder': 'Describa el requerimiento especifico', 'class': 'form-control form-control-sm'}))
    ri_pnd = forms.CharField(required=False, widget=forms.TextInput(attrs={
                             'placeholder': 'Describa el requerimiento especifico', 'class': 'form-control form-control-sm'}))
    ri_cem = forms.CharField(required=False, widget=forms.TextInput(attrs={
                             'placeholder': 'Describa el requerimiento especifico', 'class': 'form-control form-control-sm'}))
    ri_pstc = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Describa el requerimiento especifico', 'class': 'form-control form-control-sm'}))
    ri_otro = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Describa el requerimiento especifico', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MB_Requerimientos
        fields = ['ri_ods', 'ri_ocde', 'ri_ci',
                  'ri_pnd', 'ri_cem', 'ri_pstc', 'ri_otro']

    def clean(self):
        
        ri_ods = self.cleaned_data.get('ri_ods') 
        ri_ocde = self.cleaned_data.get('ri_ocde') 
        ri_ci = self.cleaned_data.get('ri_ci')                  
        ri_pnd = self.cleaned_data.get('ri_pnd')
        ri_cem = self.cleaned_data.get('ri_cem') 
        ri_pstc = self.cleaned_data.get('ri_pstc')
        ri_otro = self.cleaned_data.get('ri_otro') 
        
        if ri_ods:
            if len(ri_ods) > 2000: 
                self._errors['ri_ods'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])
        
        if ri_ocde:
            if len(ri_ocde) > 2000: 
                self._errors['ri_ocde'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        if ri_ci:
            if len(ri_ci) > 2000: 
                self._errors['ri_ci'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])
        
        if ri_pnd:
            if len(ri_pnd) > 2000: 
                self._errors['ri_pnd'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])
        
        if ri_cem:
            if len(ri_cem) > 2000: 
                self._errors['ri_cem'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])
 
        if ri_pstc:
            if len(ri_pstc) > 2000: 
                self._errors['ri_pstc'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])
        
        if ri_otro:
            if len(ri_otro) > 2000: 
                self._errors['ri_otro'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)

# modulo B pregunta 6


class MB_PrinUsuariosForm(forms.ModelForm):

    org_int = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))
    pres_rep = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))
    misnit = forms.CharField(required=False, widget=forms.TextInput(attrs={
                             'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))
    org_cont = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))
    o_ent_o_nac = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))
    ent_o_terr = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))
    gremios = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))
    ent_privadas = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))
    dep_misma_entidad = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))
    academia = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MB_PrinUsuarios
        fields = ['org_int', 'pres_rep', 'misnit', 'org_cont', 'o_ent_o_nac', 'ent_o_terr',  'gremios',
                  'ent_privadas', 'dep_misma_entidad', 'academia'
                  ]

    def clean(self):
        
        org_int = self.cleaned_data.get('org_int')
        pres_rep = self.cleaned_data.get('pres_rep')
        misnit = self.cleaned_data.get('misnit')
        org_cont = self.cleaned_data.get('org_cont')
        o_ent_o_nac = self.cleaned_data.get('o_ent_o_nac')
        ent_o_terr = self.cleaned_data.get('ent_o_terr')
        gremios = self.cleaned_data.get('gremios')
        ent_privadas = self.cleaned_data.get('ent_privadas')
        dep_misma_entidad = self.cleaned_data.get('dep_misma_entidad')
        academia = self.cleaned_data.get('academia')
            
        if org_int:
            if len(org_int) > 2000: 
                self._errors['org_int'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        if pres_rep:
            if len(pres_rep) > 2000: 
                self._errors['pres_rep'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        if misnit:
            if len(misnit) > 2000: 
                self._errors['misnit'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])
        
        if org_cont:
            if len(org_cont) > 2000: 
                self._errors['org_cont'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        if o_ent_o_nac:
            if len(o_ent_o_nac) > 2000: 
                self._errors['o_ent_o_nac'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])
        
        if ent_o_terr:
            if len(ent_o_terr) > 2000: 
                self._errors['ent_o_terr'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        if gremios:
            if len(gremios) > 2000: 
                self._errors['gremios'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        if ent_privadas:
            if len(ent_privadas) > 2000: 
                self._errors['ent_privadas'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])
          
        if dep_misma_entidad:
            if len(dep_misma_entidad) > 2000: 
                self._errors['dep_misma_entidad'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        if academia:
            if len(academia) > 2000: 
                self._errors['academia'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])
        
        # return any errors if found 
        return(self.cleaned_data)
        
# Modulo C pregunta 2

class MC_UnidadObservacionForm(forms.ModelForm):
    mc_otra = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MC_UnidadObservacion
        fields = ['mc_otra']

    def clean(self):
        
        mc_otra = self.cleaned_data.get('mc_otra')
        
        if mc_otra:
            if len(mc_otra) > 25000: 
                self._errors['mc_otra'] = self.error_class([ 
                    'Máximo 2500 caracteres requeridos'])

         # return any errors if found 
        return(self.cleaned_data)

# Modulo C pregunta 4


class MC_ObtencionDatoForm(forms.ModelForm):
    mc_ra_cual = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Enumere el registro administrativo', 'class': 'form-control form-control-sm mb-1medio'}))
    
    mc_ra_entidad = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Enumere la Entidad que realiza el registro', 'class': 'form-control form-control-sm'}))
    
    mc_oe_cual = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Enumere la operación estadística', 'class': 'form-control form-control-sm mb-1medio'}))
    
    mc_oe_entidad = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Enumere la Entidad que realiza la operación', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MC_ObtencionDato
        fields = ['mc_ra_cual',  'mc_ra_entidad',
                  'mc_oe_cual',  'mc_oe_entidad']

    def clean(self):
            
        mc_ra_cual = self.cleaned_data.get('mc_ra_cual') 
        mc_ra_entidad = self.cleaned_data.get('mc_ra_entidad')
        mc_oe_cual = self.cleaned_data.get('mc_oe_cual')
        mc_oe_entidad = self.cleaned_data.get('mc_oe_entidad')
        
        if mc_ra_cual:
            if len(mc_ra_cual) > 2000: 
                self._errors['mc_ra_cual'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])
        
        if mc_ra_entidad:
            if len(mc_ra_entidad) > 2000: 
                self._errors['mc_ra_entidad'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        if mc_oe_cual:
            if len(mc_oe_cual) > 2000: 
                self._errors['mc_oe_cual'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        if mc_oe_entidad:
            if len(mc_oe_entidad) > 2000: 
                self._errors['mc_oe_entidad'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)

    
# Modulo C pregunta 5 probabilistico

class MC_MuestreoProbabilisticoForm(forms.ModelForm):
    prob_otro = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MC_MuestreoProbabilistico
        fields = ['prob_otro']

    def clean(self):
            
        prob_otro = self.cleaned_data.get('prob_otro')
        
        if prob_otro:
            if len(prob_otro) > 150: 
                self._errors['prob_otro'] = self.error_class([ 
                    'Máximo 150 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)


# Modulo C pregunta 5  No probabilistico

class MC_MuestreoNoProbabilisticoForm(forms.ModelForm):
    no_prob_otro = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MC_MuestreoNoProbabilistico
        fields = ['no_prob_otro']

    def clean(self):
            
        no_prob_otro = self.cleaned_data.get('no_prob_otro')
        
        if no_prob_otro:
            if len(no_prob_otro) > 500:
                self._errors['no_prob_otro'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)


# Modulo C pregunta 6
class MC_TipoMarcoForm(forms.ModelForm):
    otro_tipo_marco = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MC_TipoMarco
        fields = ['otro_tipo_marco']

    def clean(self):
            
        otro_tipo_marco = self.cleaned_data.get('otro_tipo_marco')
        
        if otro_tipo_marco:
            if len(otro_tipo_marco) > 800: 
                self._errors['otro_tipo_marco'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)


# Modulo C pregunta 7
class MC_DocsDesarrolloForm(forms.ModelForm):
    otro_docs = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MC_DocsDesarrollo
        fields = ['otro_docs']

    def clean(self):
            
        otro_docs = self.cleaned_data.get('otro_docs')
        
        if otro_docs:
            if len(otro_docs) > 800: 
                self._errors['otro_docs'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)


# Modulo C pregunta 8
class MC_ConceptosEstandarizadosForm(forms.ModelForm):
    org_in_cuales = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm mb-1medio'}))
    ent_ordnac_cuales = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm mb-1medio'}))
    leye_dec_cuales = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm mb-1medio'}))
    otra_cual_conp = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm mb-1medio'}))
    ningu_pq = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Por qué?', 'class': 'form-control form-control-sm mb-1medio'}))

    class Meta:
        model = MC_ConceptosEstandarizados
        fields = ['org_in_cuales', 'ent_ordnac_cuales',
                  'leye_dec_cuales', 'otra_cual_conp', 'ningu_pq']

    def clean(self):
            
        org_in_cuales = self.cleaned_data.get('org_in_cuales')
        ent_ordnac_cuales = self.cleaned_data.get('ent_ordnac_cuales')
        leye_dec_cuales = self.cleaned_data.get('leye_dec_cuales')
        otra_cual_conp = self.cleaned_data.get('otra_cual_conp')
        ningu_pq = self.cleaned_data.get('otra_cual_conp')
        
        if org_in_cuales:
            if len(org_in_cuales) > 500: 
                self._errors['org_in_cuales'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])

        if ent_ordnac_cuales:
            if len(ent_ordnac_cuales) > 500: 
                self._errors['ent_ordnac_cuales'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])

        if leye_dec_cuales:
            if len(leye_dec_cuales) > 500: 
                self._errors['leye_dec_cuales'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])

        if otra_cual_conp:
            if len(otra_cual_conp) > 500: 
                self._errors['otra_cual_conp'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])

        if ningu_pq:
            if len(ningu_pq) > 500: 
                self._errors['ningu_pq'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])
        
        # return any errors if found 
        return(self.cleaned_data)

# Modulo C pregunta 9
class MC_ClasificacionesForm(forms.ModelForm):
    otra_cual_clas = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otra', 'class': 'form-control form-control-sm mb-1medio w-75 ml-3'}))
    no_pq = forms.CharField(required=False, widget=forms.TextInput(attrs={
                            'placeholder': '¿Por qué?', 'class': 'form-control form-control-sm mb-1medio'}))

    class Meta:
        model = MC_Clasificaciones
        fields = ['otra_cual_clas', 'no_pq']

    def clean(self):
            
        otra_cual_clas  = self.cleaned_data.get('otra_cual_clas')
        no_pq  = self.cleaned_data.get('no_pq')
        
        if otra_cual_clas:
            if len(otra_cual_clas) > 800: 
                self._errors['otra_cual_clas'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        if no_pq:
            if len(no_pq) > 800: 
                self._errors['no_pq'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)

# Modulo C pregunta 10
class MC_CoberturaGeograficaForm(forms.ModelForm):
    tot_regional = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuántos?', 'class': 'form-control form-control-sm'}))
    cual_regional = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm heigthTextArea'}))
    tot_dep = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuántos?', 'class': 'form-control form-control-sm'}))
    cual_dep = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm heigthTextArea'}))
    tot_are_metr = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuántos?', 'class': 'form-control form-control-sm'}))
    cual_are_metr = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm heigthTextArea'}))
    tot_mun = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuántos?', 'class': 'form-control form-control-sm'}))
    cual_mun = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm heigthTextArea'}))
    tot_otro = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuántos?', 'class': 'form-control form-control-sm'}))
    cual_otro = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm heigthTextArea'}))

    class Meta:
        model = MC_CoberturaGeografica
        fields = ['tot_regional', 'cual_regional',  'tot_dep', 'cual_dep',
                  'tot_are_metr', 'cual_are_metr', 'tot_mun', 'cual_mun', 'tot_otro', 'cual_otro']

    
    def clean(self):
            
        tot_regional = self.cleaned_data.get('tot_regional')
        cual_regional  = self.cleaned_data.get('cual_regional')
        tot_dep  = self.cleaned_data.get('tot_dep')
        cual_dep  = self.cleaned_data.get('cual_dep')
        tot_are_metr  = self.cleaned_data.get('tot_are_metr')
        cual_are_metr  = self.cleaned_data.get('cual_are_metr')
        tot_mun  = self.cleaned_data.get('tot_mun')
        cual_mun  = self.cleaned_data.get('cual_mun')
        tot_otro  = self.cleaned_data.get('tot_otro')
        cual_otro  = self.cleaned_data.get('cual_otro')
        
        if tot_regional:
            if len(tot_regional) > 40: 
                self._errors['tot_regional'] = self.error_class([ 
                    'Máximo 40 caracteres requeridos'])

        if cual_regional:
            if len(cual_regional) > 600: 
                self._errors['cual_regional'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])
        
        if tot_dep:
            if len(tot_dep) > 40: 
                self._errors['tot_dep'] = self.error_class([ 
                    'Máximo 40 caracteres requeridos'])

        if cual_dep:
            if len(cual_dep) > 600: 
                self._errors['cual_dep'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])
        
        if tot_are_metr:
            if len(tot_are_metr) > 40: 
                self._errors['tot_are_metr'] = self.error_class([ 
                    'Máximo 40 caracteres requeridos'])

        if cual_are_metr:
            if len(cual_are_metr) > 600: 
                self._errors['cual_are_metr'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])
        
        if tot_mun:
            if len(tot_mun) > 40: 
                self._errors['tot_mun'] = self.error_class([ 
                    'Máximo 40 caracteres requeridos'])

        if cual_mun:
            if len(cual_mun) > 600: 
                self._errors['cual_mun'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])

        if tot_otro:
            if len(tot_otro) > 600: 
                self._errors['tot_otro'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])

        if cual_otro:
            if len(cual_otro) > 600: 
                self._errors['cual_otro'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])
        
        # return any errors if found 
        return(self.cleaned_data)


# Modulo C pregunta 11

class MC_DesagregacionInformacionForm(forms.ModelForm):
    des_tot_regional = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuántos?', 'class': 'form-control form-control-sm'}))
    des_cual_regional = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))
    des_tot_dep = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuántos?', 'class': 'form-control form-control-sm'}))
    des_cual_dep = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))
    des_tot_are_metr = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuántos?', 'class': 'form-control form-control-sm'}))
    des_cual_are_metr = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))
    des_tot_mun = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuántos?', 'class': 'form-control form-control-sm'}))
    des_cual_mun = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))
    des_tot_otro = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuántos?', 'class': 'form-control form-control-sm'}))
    des_cual_otro = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))
    des_grupo_otro = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MC_DesagregacionInformacion
        fields = ['des_tot_regional', 'des_cual_regional',  'des_tot_dep', 'des_cual_dep', 'des_tot_are_metr', 'des_cual_are_metr', 'des_tot_mun', 'des_cual_mun',
                  'des_tot_otro', 'des_cual_otro', 'des_grupo_otro']

    def clean(self):
            
        des_tot_regional = self.cleaned_data.get('des_tot_regional')
        des_cual_regional = self.cleaned_data.get('des_cual_regional')
        des_tot_dep = self.cleaned_data.get('des_tot_dep')
        des_cual_dep = self.cleaned_data.get('des_cual_dep')
        des_tot_are_metr = self.cleaned_data.get('des_tot_are_metr')
        des_cual_are_metr = self.cleaned_data.get('des_cual_are_metr')
        des_tot_mun = self.cleaned_data.get('des_tot_mun')
        des_cual_mun = self.cleaned_data.get('des_cual_mun')
        des_tot_otro = self.cleaned_data.get('des_tot_otro')
        des_cual_otro = self.cleaned_data.get('des_cual_otro')
        des_grupo_otro = self.cleaned_data.get('des_grupo_otro')
        
        if des_tot_regional:
            if len(des_tot_regional) > 40: 
                self._errors['des_tot_regional'] = self.error_class([ 
                    'Máximo 40 caracteres requeridos'])

        if des_cual_regional:
            if len(des_cual_regional) > 600: 
                self._errors['des_cual_regional'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])

        if des_tot_dep:
            if len(des_tot_dep) > 40: 
                self._errors['des_tot_dep'] = self.error_class([ 
                    'Máximo 40 caracteres requeridos'])

        if des_cual_dep:
            if len(des_cual_dep) > 600: 
                self._errors['des_cual_dep'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])

        if des_tot_are_metr:
            if len(des_tot_are_metr) > 40: 
                self._errors['des_tot_are_metr'] = self.error_class([ 
                    'Máximo 40 caracteres requeridos'])

        if des_cual_are_metr:
            if len(des_cual_are_metr) > 600: 
                self._errors['des_cual_are_metr'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])

        if des_tot_mun:
            if len(des_tot_mun) > 40: 
                self._errors['des_tot_mun'] = self.error_class([ 
                    'Máximo 40 caracteres requeridos'])

        if des_cual_mun:
            if len(des_cual_mun) > 600: 
                self._errors['des_cual_mun'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])

        if des_tot_otro:
            if len(des_tot_otro) > 40: 
                self._errors['des_tot_otro'] = self.error_class([ 
                    'Máximo 40 caracteres requeridos'])

        if des_cual_otro:
            if len(des_cual_otro) > 600: 
                self._errors['des_cual_otro'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])       

        if des_grupo_otro:
            if len(des_grupo_otro) > 600: 
                self._errors['des_grupo_otro'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])       

        # return any errors if found 
        return(self.cleaned_data)

# Modulo C pregunta 13
class MC_FuenteFinanciacionForm(forms.ModelForm):
    r_otros = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MC_FuenteFinanciacion
        fields = ['r_otros']

    def clean(self):
            
        r_otros = self.cleaned_data.get('r_otros')
       
        if r_otros:
            if len(r_otros) > 500: 
                self._errors['r_otros'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)


# Modulo D pregunta 1

class MD_MedioDatosForm(forms.ModelForm):
    sis_info = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm'}))
    md_otro = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MD_MedioDatos
        fields = ['sis_info',  'md_otro']

    def clean(self):
            
        sis_info = self.cleaned_data.get('sis_info')
        md_otro = self.cleaned_data.get('md_otro')
       
        if sis_info:
            if len(sis_info) > 300: 
                self._errors['sis_info'] = self.error_class([ 
                    'Máximo 300 caracteres requeridos'])

        if md_otro:
            if len(md_otro) > 300: 
                self._errors['md_otro'] = self.error_class([ 
                    'Máximo 300 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)


# Modulo D pregunta 2

class MD_PeriodicidadOeForm(forms.ModelForm):
    per_otro = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MD_PeriodicidadOe
        fields = ['per_otro']

    def clean(self):
            
        per_otro = self.cleaned_data.get('per_otro')
    
        if per_otro:
            if len(per_otro) > 300: 
                self._errors['per_otro'] = self.error_class([ 
                    'Máximo 300 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)


# Modulo D pregunta 3

class MD_HerramProcesamientoForm(forms.ModelForm):
    herr_otro = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MD_HerramProcesamiento
        fields = ['herr_otro']

    def clean(self):
            
        herr_otro = self.cleaned_data.get('herr_otro')
    
        if herr_otro:
            if len(herr_otro) > 300: 
                self._errors['herr_otro'] = self.error_class([ 
                    'Máximo 300 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)

# modulo E pregunta 1

class ME_AnalisisResultadosForm(forms.ModelForm):
    ana_otro = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = ME_AnalisisResultados
        fields = ['ana_otro']

    def clean(self):
            
        ana_otro = self.cleaned_data.get('ana_otro')
    
        if ana_otro:
            if len(ana_otro) > 300: 
                self._errors['ana_otro'] = self.error_class([ 
                    'Máximo 300 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)

# Modulo F pregunta 1

class MF_MediosDifusionForm(forms.ModelForm):
    dif_otro = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MF_MediosDifusion
        fields = ['dif_otro']

    def clean(self):
            
        dif_otro = self.cleaned_data.get('dif_otro')
    
        if dif_otro:
            if len(dif_otro) > 300: 
                self._errors['dif_otro'] = self.error_class([ 
                    'Máximo 300 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)


# Modulo F pregunta 4

class MF_FechaPublicacionForm(forms.ModelForm):
    fecha = forms.DateField(
        required=False,
        widget=forms.DateInput(format='%Y-%m-%d',
                               attrs={'class': 'form-control datepicker'}),
        input_formats=('%Y-%m-%d', )
        )
    no_hay = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Por qué?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MF_FechaPublicacion
        fields = ['fecha', 'no_hay']

    def clean(self):
            
        no_hay = self.cleaned_data.get('no_hay')
    
        if no_hay:
            if len(no_hay) > 150: 
                self._errors['no_hay'] = self.error_class([ 
                    'Máximo 150 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)

# Modulo F pregunta 5


class MF_FrecuenciaDifusionForm(forms.ModelForm):
    no_definido = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MF_FrecuenciaDifusion
        fields = ['no_definido']

    def clean(self):
            
        no_definido = self.cleaned_data.get('no_definido')
    
        if no_definido:
            if len(no_definido) > 300: 
                self._errors['no_definido'] = self.error_class([ 
                    'Máximo 300 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)

# Modulo F pregunta 6


class MF_ProductosDifundirForm(forms.ModelForm):
    difundir_otro = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MF_ProductosDifundir
        fields = ['difundir_otro']

    def clean(self):
            
        difundir_otro = self.cleaned_data.get('difundir_otro')
    
        if difundir_otro:
            if len(difundir_otro) > 500: 
                self._errors['difundir_otro'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)
    

# Modulo F pregunta 7


class MF_OtrosProductosForm(forms.ModelForm):
    ser_hist_desde = forms.DateField(
        required=False,
        widget=forms.DateInput(format='%Y-%m-%d',
                               attrs={'class': 'form-control datepicker'}),
        input_formats=('%Y-%m-%d', )
        )
    ser_hist_hasta = forms.DateField(
        required=False,
        widget=forms.DateInput(format='%Y-%m-%d',
                               attrs={'class': 'form-control datepicker'}),
        input_formats=('%Y-%m-%d', )
        )
    microdatos_desde = forms.DateField(
        required=False,
        widget=forms.DateInput(format='%Y-%m-%d',
                               attrs={'class': 'form-control datepicker'}),
        input_formats=('%Y-%m-%d', )
        )
    microdatos_hasta = forms.DateField(
        required=False,
        widget=forms.DateInput(format='%Y-%m-%d',
                               attrs={'class': 'form-control datepicker'}),
        input_formats=('%Y-%m-%d', )
        )
    op_url = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'URL?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MF_OtrosProductos
        fields = ['ser_hist_desde', 'ser_hist_hasta',
                  'microdatos_desde', 'microdatos_hasta', 'op_url']

    def clean(self):
            
        op_url = self.cleaned_data.get('op_url')

        if op_url:
            if len(op_url) > 300: 
                self._errors['op_url'] = self.error_class([ 
                    'Máximo 300 caracteres requeridos'])
         # return any errors if found 
        return(self.cleaned_data)


# Modulo F pregunta 8
class MF_ResultadosSimilaresForm(forms.ModelForm):
    rs_entidad = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Entidad', 'class': 'form-control form-control-sm'}))
    rs_oe = forms.CharField(required=False, widget=forms.TextInput(attrs={
                            'placeholder': 'Operación estadística / Indicadores', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MF_ResultadosSimilares
        fields = ['rs_entidad', 'rs_oe']

    def clean(self):
            
        rs_entidad = self.cleaned_data.get('rs_entidad')
        rs_oe = self.cleaned_data.get('rs_oe')
    
        if rs_entidad:
            if len(rs_entidad) > 500: 
                self._errors['rs_entidad'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])

        if rs_oe:
            if len(rs_oe) > 500: 
                self._errors['rs_oe'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)


# Modulo F pregunta 9
class MF_HPSistemaInfoForm(forms.ModelForm):
    si_cual = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MF_HPSistemaInfo
        fields = ['si_cual']

    def clean(self):
            
        si_cual = self.cleaned_data.get('si_cual')
    
        if si_cual:
            if len(si_cual) > 500: 
                self._errors['si_cual'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)

# forms edit OOEE


class EditOEForm(forms.ModelForm):

    codigo_oe = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Código', 'class': 'form-control form-control-sm'}))
    entidad = forms.ModelChoiceField(label="Nombre de la entidad", empty_label=None, queryset=Entidades_oe.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm', "title": "A. Identificación: Nombre de entidad"}))
    area_tematica = forms.ModelChoiceField(label="Area Temática", empty_label=None, queryset=AreaTematica.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm', "title": "A. Identificación: Área Temática"}))
    tema = forms.ModelChoiceField(label="Tema", empty_label=None, queryset=Tema.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm', "title": "A. Identificación: Tema"}))
    tema_compartido = forms.ModelMultipleChoiceField(
        required=False,  queryset=TemaCompartido.objects.all(), widget=forms.CheckboxSelectMultiple(attrs={"title": "A. Identificación: Tema Compartido"}))
    sala_esp = forms.ModelMultipleChoiceField(
        required=False,  queryset=SalaEspecializada.objects.all(), widget=forms.CheckboxSelectMultiple(attrs={"title": "A. Identificación: Sala Especializada"}))
    entidad_resp2 = forms.ModelChoiceField(required=False, label="Nombre de la entidad",  queryset=Entidades_oe.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm', "title": "A. Identificación: Entidad Responsable 2"}))
    entidad_resp3 = forms.ModelChoiceField(required=False, label="Nombre de la entidad",  queryset=Entidades_oe.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm', "title": "A. Identificación: Entidad Responsable 3"}))
    nombre_dep = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Dependencia', 'class': 'form-control form-control-sm', "title": "A. Identificación: Nombre Dependencia"}))
    nombre_dir = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Director', 'class': 'form-control form-control-sm', "title": "A. Identificación: Nombre Director"}))
    cargo_dir = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Cargo', 'class': 'form-control form-control-sm', "title": "A. Identificación: Cargo Director"}))
    correo_dir = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Correo electrónico', 'class': 'form-control form-control-sm', "title": "A. Identificación: Correo Director "}))
    tel_dir = forms.IntegerField(required=False, widget=forms.NumberInput(
        attrs={'placeholder': '999-9999', 'class': 'form-control form-control-sm', "title": "A. Identificación: Teléfono Director"}))
    nombre_resp = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Responsable', 'class': 'form-control form-control-sm', "title": "A. Identificación: Nombre Responsable"}))
    cargo_resp = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Cargo', 'class': 'form-control form-control-sm', "title": "A. Identificación: Cargo Responsable"}))
    correo_resp = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Correo electrónico', 'class': 'form-control form-control-sm', "title": "A. Identificación: Correo Responsable"}))
    tel_resp = forms.IntegerField(required=False, widget=forms.NumberInput(
        attrs={'placeholder': '999-9999', 'class': 'form-control form-control-sm', "title": "A. Identificación: Teléfono Responsable"}))
    nombre_oe = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Nombre de la operación estadística', 'class': 'form-control form-control-sm', "title": "Módulo B Pregunta 1: Nombre de la operación estadística"}))
    objetivo_oe = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Objetivo de la operación estadística', 'class': 'form-control form-control-sm', "title": "Módulo B Pregunta 2: Objetivo de la operación estadística"}))
    nombre_est = forms.ModelChoiceField(label="Estado", empty_label=None, queryset=OoeeState.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm', 'onchange':'fondo(this)', "title": "Estado de la operación estadística" }))

    fase = forms.BooleanField(required=False, initial=True, widget=forms.CheckboxInput(
        attrs={"id": "id_fase", "data-toggle": "toggle", "data-size": "small", " data-on": "Si",  "data-off": "No", "data-onstyle": "success", "data-offstyle": "danger", "title": "Módulo B Pregunta 3: Entidad(es) que participa(n) en alguna de las fases del proceso estadístico"}))
    #fase = forms.BooleanField(required=False, initial=True)
    norma = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title": "Módulo B Pregunta 4:  Que Norma(s) soporta la producción de información de la Operación Estadística"}), queryset=Norma.objects.all())
    requerimientos = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title": "Módulo B Pregunta 5: Que requerimientos satisface la Operación Estadística"}), queryset=Requerimientos.objects.all())
    pri_usuarios = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title": "Módulo B Pregunta 6: Principales Usuarios de la Operación Estadística"}), queryset=PrinUsuarios.objects.all())
    pob_obje = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': '¿Cuál es la población objetivo de la operación estadística?', 'class': 'form-control form-control-sm', "title": "Módulo C Pregunta 1: Población objetivo de la Operación Estadística"}))
    uni_observacion = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title": "Módulo C Pregunta 2:  Unidad de observación de la Operación Estadística"}), queryset=UnidadObservacion.objects.all())
    tipo_operacion = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={'class': 'list-group list-group-horizontal-sm', "title": "Módulo C Pregunta 3:  Tipo de Operación Estadística"}), queryset=TipoOperacion.objects.all())
    obt_dato = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title": "Módulo C Pregunta 4:  Donde se obtienen los datos"}), queryset=ObtencionDato.objects.all())
    
    rraa_lista = forms.ModelMultipleChoiceField(
        required=False, widget=forms.SelectMultiple(attrs={"class":"selectpicker", "data-live-search":"true", "title": "Módulo C Pregunta 4: Donde se obtienen los datos"}), queryset=RegistroAdministrativo.objects.all())
    
    ooee_lista = forms.ModelMultipleChoiceField(
        required=False, widget=forms.SelectMultiple(attrs={"class":"selectpicker", "data-live-search":"true", "title": "Módulo C Pregunta 4: Donde se obtienen los datos OOEE"}), queryset=ListaDeOOEE.objects.all())

    tipo_probabilistico = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title": "Módulo C Pregunta 5: Tipo de muestreo probabilistico"}), queryset=MuestreoProbabilistico.objects.all())
    tipo_no_probabilistico = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title": "Módulo C Pregunta 5: Tipo de muestreo no probabilistico"}), queryset=MuestreoNoProbabilistico.objects.all())
    marco_estad = forms.BooleanField(required=False, initial=False,  widget=forms.CheckboxInput(
        attrs={"id": "id_marco", "data-toggle": "toggle", "data-size": "small", " data-on": "Si",  "data-off": "No", "data-onstyle": "success", "data-offstyle": "danger", "title": "Módulo C Pregunta 6:  Marco estadístico para identificar y ubicar las unidades de observación"}))
    tipo_marco = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title": "Módulo C Pregunta 6:  Tipo de marco estadístico"}), queryset=TipoMarco.objects.all())
    docs_des = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title": "Módulo C Pregunta 7: Cuáles son los documentos que se elaboran para el desarrollo de la Operación Estadística"}), queryset=DocsDesarrollo.objects.all())
    lista_conc = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title": "Módulo C Pregunta 8:  Que conceptos estandarizados utiliza la Operación Estadística"}), queryset=ConceptosEstandarizados.objects.all())
    nome_clas = forms.BooleanField(required=False, initial=False, widget=forms.CheckboxInput(
        attrs={"id": "id_nome_clas", "data-toggle": "toggle", "data-size": "small", " data-on": "Si",  "data-off": "No", "data-onstyle": "success", "data-offstyle": "danger", "title": "Módulo C Pregunta 9: Se utiliza nomenclaturas y/o clasificaciones?"}))
    #nome_clas = forms.BooleanField(required=False, initial=True)
    nombre_cla = forms.ModelMultipleChoiceField(
        required=False,  queryset=Clasificaciones.objects.all(), widget=forms.CheckboxSelectMultiple(attrs={"title": "Módulo C Pregunta 9:  Cuales clasificaciones y/o nomenclaturas"}))
    cob_geo = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(
        attrs={'class': 'list-group', "title": "Módulo C Pregunta 10: Cobertura geográfica de la Operación Estadística"}), queryset=CoberturaGeografica.objects.all())
    opc_desag = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(
        attrs={'class': 'list-group list-group-horizontal-sm', "title": "Módulo C Pregunta 11: Desagregación de la información estadística"}), queryset=DesagregacionInformacion.objects.all())
    des_geo = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title": "Módulo C Pregunta 11.1: Desagregación Geográfica"}), queryset=DesagregacionGeografica.objects.all())
    des_zona = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(
        attrs={'class': 'list-group list-group-horizontal-sm', "title": "Módulo C Pregunta 11.2: Desagregación Zona"}), queryset=DesagregacionZona.objects.all())
    des_grupo = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title": "Módulo C Pregunta 11.3:  Desagregación Grupo"}), queryset=DesagregacionGrupos.objects.all())
    ca_anual = forms.IntegerField(required=False, widget=forms.NumberInput(
        attrs={'placeholder': '$ Costo anual', 'class': 'form-control form-control-sm', "title": "Módulo C Pregunta 12: Costo Anual"}))
    cb_anual = forms.BooleanField(
        required=False, initial=True, label='No sabe', widget=forms.CheckboxInput(attrs={"title": "Módulo C Pregunta 12: Costo Anual"}))

    fuentes = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title": "Módulo C Pregunta 13: Fuentes de financiación de la Operación Estadística"}), queryset=FuenteFinanciacion.objects.all())
    variable_file = forms.FileField(
        label='Select a file',
        help_text='max. 42 megabytes',
        required=False,
        widget=forms.ClearableFileInput(attrs={"title": "Módulo C pregunta 14: Si tiene diccionario de datos o listado de variables con su descripción anéxelo en formato excel"})
    )
    med_obt = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title": "Módulo D Pregunta 1:  Medio de obtención de datos"}), queryset=MedioDatos.objects.all())
    periodicidad = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title": "Módulo D Pregunta 2:  Periodicidad de recolección o acopio de la Operación Estadística"}), queryset=PeriodicidadOe.objects.all())
    h_proc = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title": "Módulo D Pregunta 3:  Herramientas utilizadas en el procesamiento de datos"}), queryset=HerramProcesamiento.objects.all())
    descrip_proces = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Haga una breve descripción de la manera cómo se realiza el procesamiento de los datos', 'class': 'form-control form-control-sm', "title": "Módulo D Pregunta 4:  Descripción de cómo se realiza el procesamiento de los datos"}))
    a_resul = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title": "Módulo e Pregunta 1: Tipo de análisis que se realiza a los resultados obtenidos"}), queryset=AnalisisResultados.objects.all())
    m_dif = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title": "Módulo F Pregunta 1:  Medio(s) que difunde los resultados estadísticos"}), queryset=MediosDifusion.objects.all())
    res_est_url = forms.URLField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Pagina Web', 'class': 'form-control form-control-sm', "title": "Módulo F Pregunta 2: Página web donde se encuentran los resultados estadísticos"}))
    dispo_desde = forms.DateField(
        required=False,
        widget=forms.DateInput(format='%Y-%m-%d',
                               attrs={'class': 'form-control datepicker',
                               'title': 'Módulo F Pregunta 3: Fechas de disponibilidad de los resultados estadísticos (Agregados e indicadores) fecha desde'}),
        input_formats=('%Y-%m-%d', )
        ) 
    
    
    dispo_hasta = forms.DateField(
        required=False,
        widget=forms.DateInput(format='%Y-%m-%d',
                               attrs={'class': 'form-control datepicker',
                               'title': 'Módulo F Pregunta 3: Fechas de disponibilidad de los resultados estadísticos (Agregados e indicadores) fecha hasta'}),
        input_formats=('%Y-%m-%d', )
        )  
    
    f_publi = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title": "Módulo F Pregunta 4:  Próxima fecha de publicación de resultados estadísticos"}), queryset=FechaPublicacion.objects.all())
    fre_dif = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title": "Módulo F Pregunta 5: Frecuencia de difusión de los resultados estadísticos"}), queryset=FrecuenciaDifusion.objects.all())
    pro_dif = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title": "Módulo F Pregunta 6: Productos que utiliza para difundir los resultados estadísticos"}), queryset=ProductosDifundir.objects.all())
    otro_prod = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title": "Módulo F Pregunta 7: Productos estadísticos de la OE que están disponibles para consulta de los usuarios"}), queryset=OtrosProductos.objects.all())

    conoce_otra = forms.BooleanField(required=False, initial=False, widget=forms.CheckboxInput(
        attrs={"id": "id_conoce_otra", "data-toggle": "toggle", "data-size": "small", " data-on": "Si",  "data-off": "No", "data-onstyle": "success", "data-offstyle": "danger", "title": "Módulo F Pregunta 8: Otra entidad que produce resultados similares a los de la operación estadística"}))
    hp_siste_infor = forms.BooleanField(required=False, initial=False, widget=forms.CheckboxInput(
        attrs={"id": "id_hp_siste_infor", "data-toggle": "toggle", "data-size": "small", " data-on": "Si",  "data-off": "No", "data-onstyle": "success", "data-offstyle": "danger", "title": "Módulo F Pregunta 9: Hace parte de algún sistema de información"}))
    observaciones = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': '(Indique el número de pregunta si es necesario ampliar, aclarar o complementar la respuesta)', 'class': 'form-control form-control-sm', "title": "Módulo G Pregunta 1: Observaciones"}))

    anexos = forms.FileField(
        label='Select a file',
        help_text='max. 42 megabytes',
        required=False,
        widget=forms.ClearableFileInput(attrs={"title": "Módulo H: Archivo de anexos"})
    )
    estado_oe_tematico = forms.ChoiceField(required=False, label="estado oe", choices=CHOICES_STATEOOOE, widget=forms.Select(
        attrs={'class': 'form-control form-control-sm', "title": "Preguntas de Control: Estado de la operación estadística"}))
    validacion_oe_tematico = forms.ChoiceField(required=False, label="validación oe", choices=CHOICES_VALIDATIONOOOE, widget=forms.Select(
        attrs={'class': 'form-control form-control-sm', "title": "Preguntas de Control: Validación de la operación estadística"}))

    class Meta:
        model = OperacionEstadistica
        fields = [
            'codigo_oe',
            'entidad',
            'area_tematica',
            'tema',
            'tema_compartido',
            'sala_esp',
            'entidad_resp2',
            'entidad_resp3',
            'nombre_dep',
            'nombre_dir',
            'cargo_dir',
            'correo_dir',
            'tel_dir',
            'nombre_resp',
            'cargo_resp',
            'correo_resp',
            'tel_resp',
            'nombre_oe',
            'objetivo_oe',
            'nombre_est',
            'fase',
            'norma',
            'requerimientos',
            'pri_usuarios',
            'pob_obje',
            'uni_observacion',
            'tipo_operacion',
            'obt_dato',
            'rraa_lista',
            'ooee_lista',
            'tipo_probabilistico',
            'tipo_no_probabilistico',
            'marco_estad',
            'tipo_marco',
            'docs_des',
            'lista_conc',
            'nome_clas',
            'nombre_cla',
            'cob_geo',
            'opc_desag',
            'des_geo',
            'des_zona',
            'des_grupo',
            'ca_anual',
            'cb_anual',
            'fuentes',
            'variable_file',
            'med_obt',
            'periodicidad',
            'h_proc',
            'descrip_proces',
            'a_resul',
            'm_dif',
            'res_est_url',
            'dispo_desde',
            'dispo_hasta',
            'f_publi',
            'fre_dif',
            'pro_dif',
            'otro_prod',
            'conoce_otra',
            'hp_siste_infor',
            'observaciones',
            'anexos',
            'estado_oe_tematico',
            'validacion_oe_tematico'
        ]

    def clean(self):
        nombre_dep = self.cleaned_data.get('nombre_dep') 
        nombre_dir = self.cleaned_data.get('nombre_dir') 
        cargo_dir = self.cleaned_data.get('cargo_dir')
        nombre_resp = self.cleaned_data.get('nombre_resp')
        cargo_resp = self.cleaned_data.get('cargo_resp')
        nombre_oe = self.cleaned_data.get('nombre_oe') 
        objetivo_oe = self.cleaned_data.get('objetivo_oe')
        pob_obje  = self.cleaned_data.get('pob_obje')
        descrip_proces = self.cleaned_data.get('descrip_proces')
        observaciones = self.cleaned_data.get('observaciones')

        if nombre_dep:
            if len(nombre_dep) > 80: 
                self._errors['nombre_dep'] = self.error_class([ 
                    'Máximo 80 caracteres requeridos'])
        
        if nombre_dir:
            if len(nombre_dir) > 80: 
                self._errors['nombre_dir'] = self.error_class([ 
                    'Máximo 80 caracteres requeridos'])

        if cargo_dir:
            if len(cargo_dir) > 80: 
                self._errors['cargo_dir'] = self.error_class([ 
                    'Máximo 80 caracteres requeridos'])

        if nombre_resp:
            if len(nombre_resp) > 80: 
                self._errors['nombre_resp'] = self.error_class([ 
                    'Máximo 80 caracteres requeridos'])

        if cargo_resp:
            if len(cargo_resp) > 80: 
                self._errors['cargo_resp'] = self.error_class([ 
                    'Máximo 80 caracteres requeridos'])

        if nombre_oe:
            if len(nombre_oe) > 600: 
                self._errors['nombre_oe'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos']) 
        
        if objetivo_oe:
            if len(objetivo_oe) > 800: 
                self._errors['objetivo_oe'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos']) 

        if pob_obje:
            if len(pob_obje) > 2000: 
                self._errors['pob_obje'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos']) 
        
        if descrip_proces:
            if len(descrip_proces) > 3000: 
                self._errors['descrip_proces'] = self.error_class([ 
                    'Máximo 3000 caracteres requeridos'])

        if observaciones:
            if len(observaciones) > 2000: 
                self._errors['observaciones'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos']) 

        # return any errors if found 
        return(self.cleaned_data)

    def clean_tel_dir(self):
        tel_dir = self.cleaned_data['tel_dir']
        if tel_dir:
            if len(str(tel_dir)) < 7 or len(str(tel_dir)) > 10:
                raise forms.ValidationError(
                    _("el número telefónico  que ingreso '%s' debe contener 10 digitos para celular o 7 para teléfono fijo. " % tel_dir))
        return tel_dir

    def clean_tel_resp(self):
        tel_resp = self.cleaned_data['tel_resp']
        if tel_resp:
            if len(str(tel_resp)) < 7 or len(str(tel_resp)) > 10:
                raise forms.ValidationError(
                    _("el número telefónico  que ingreso '%s' debe contener 10 digitos para celular o 7 para teléfono fijo. " % tel_resp))
        return tel_resp

    def clean_correo_dir(self):
        correo_dir = self.cleaned_data['correo_dir']
        if correo_dir:
            if "@" not in correo_dir:
                raise forms.ValidationError(
                    _("El correo electrónico que ingreso '%s' debe contener @." % correo_dir))

            if len(correo_dir) > 80:
                raise forms.ValidationError(
                    _("El correo electrónico que ingreso '%s' debe contener menos de 80 caracteres" % correo_dir))

        return correo_dir

    def clean_correo_resp(self):
        correo_resp = self.cleaned_data['correo_resp']
        if correo_resp:
            if "@" not in correo_resp:
                raise forms.ValidationError(
                    _("El correo electrónico que ingreso '%s' debe contener @." % correo_resp))

            if len(correo_resp) > 80:
                raise forms.ValidationError(
                    _("El correo electrónico que ingreso '%s' debe contener menos de 80 caracteres" % correo_resp))

        return correo_resp

    def clean_variable_file(self):
        variable_file = self.cleaned_data['variable_file']
        if variable_file:
            print("formulario", variable_file)
            file_type = str(variable_file).split('.')[-1]
            file_type = file_type.lower()
            if file_type not in FILE_TYPES:
                raise forms.ValidationError(
                    _("Archivo No valido '%s'  la extensión debe ser XLS o XLSX" % variable_file))

        return variable_file

    def clean_res_est_url(self):
        res_est_url = self.cleaned_data['res_est_url']
        validate = URLValidator()
        if res_est_url:
            try: 
                validate(res_est_url)
                print("pagina valida")
                return res_est_url
            except ValidationError:
                print("invalida")
                

    def clean_anexos(self):
        anexos = self.cleaned_data['anexos']
        if anexos:
            file_type_annexes = str(anexos).split('.')[-1]
            file_type_annexes = file_type_annexes.lower()
            if file_type_annexes not in FILE_TYPES_ANNEXES:
                raise forms.ValidationError(
                    _("Archivo No valido '%s'  la extensión debe ser XLS, XLSX, PDF, DOC O DOCX " % anexos))
        return anexos


# modulo B pregunta 4


class EditMB_NormaForm(forms.ModelForm):

    cp_d = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': 'Describa el requerimiento', 'class': 'form-control form-control-sm mb-1'}))
    ley_d = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Describa el requerimiento', 'class': 'form-control form-control-sm mb-1'}))
    decreto_d = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Describa el requerimiento', 'class': 'form-control form-control-sm mb-1'}))
    otra_d = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Describa el requerimiento', 'class': 'form-control form-control-sm mb-1'}))

    class Meta:
        model = MB_Norma
        fields = ['cp_d', 'ley_d', 'decreto_d',  'otra_d']

    def clean(self):

        cp_d = self.cleaned_data.get('cp_d') 
        ley_d = self.cleaned_data.get('ley_d') 
        decreto_d = self.cleaned_data.get('decreto_d')
        otra_d = self.cleaned_data.get('otra_d')
        
        if cp_d:
            if len(cp_d) > 2000: 
                self._errors['cp_d'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos']) 
        
        if ley_d:
            if len(ley_d) > 2000: 
                self._errors['ley_d'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos']) 
        
        if decreto_d:
            if len(decreto_d) > 2000: 
                self._errors['decreto_d'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos']) 

        if otra_d:
            if len(otra_d) > 2000: 
                self._errors['otra_d'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos']) 

        # return any errors if found 
        return(self.cleaned_data)


# modulo B pregunta 5

class EditMB_RequerimientoForm(forms.ModelForm):

    ri_ods = forms.CharField(required=False, widget=forms.TextInput(attrs={
                             'placeholder': 'Describa el requerimiento especifico', 'class': 'form-control form-control-sm'}))
    ri_ocde = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Describa el requerimiento especifico', 'class': 'form-control form-control-sm'}))
    ri_ci = forms.CharField(required=False, widget=forms.TextInput(attrs={
                            'placeholder': 'Describa el requerimiento especifico', 'class': 'form-control form-control-sm'}))
    ri_pnd = forms.CharField(required=False, widget=forms.TextInput(attrs={
                             'placeholder': 'Describa el requerimiento especifico', 'class': 'form-control form-control-sm'}))
    ri_cem = forms.CharField(required=False, widget=forms.TextInput(attrs={
                             'placeholder': 'Describa el requerimiento especifico', 'class': 'form-control form-control-sm'}))
    ri_pstc = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Describa el requerimiento especifico', 'class': 'form-control form-control-sm'}))
    ri_otro = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Describa el requerimiento especifico', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MB_Requerimientos
        fields = ['ri_ods', 'ri_ocde', 'ri_ci',
                  'ri_pnd', 'ri_cem', 'ri_pstc', 'ri_otro']

    def clean(self):
        
        ri_ods = self.cleaned_data.get('ri_ods') 
        ri_ocde = self.cleaned_data.get('ri_ocde') 
        ri_ci = self.cleaned_data.get('ri_ci')                  
        ri_pnd = self.cleaned_data.get('ri_pnd')
        ri_cem = self.cleaned_data.get('ri_cem') 
        ri_pstc = self.cleaned_data.get('ri_pstc')
        ri_otro = self.cleaned_data.get('ri_otro') 
        
        if ri_ods:
            if len(ri_ods) > 2000: 
                self._errors['ri_ods'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])
        
        if ri_ocde:
            if len(ri_ocde) > 2000: 
                self._errors['ri_ocde'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        if ri_ci:
            if len(ri_ci) > 2000: 
                self._errors['ri_ci'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])
        
        if ri_pnd:
            if len(ri_pnd) > 2000: 
                self._errors['ri_pnd'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])
        
        if ri_cem:
            if len(ri_cem) > 2000: 
                self._errors['ri_cem'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])
 
        if ri_pstc:
            if len(ri_pstc) > 2000: 
                self._errors['ri_pstc'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])
        
        if ri_otro:
            if len(ri_otro) > 2000: 
                self._errors['ri_otro'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)

# modulo B pregunta 6

class EditMB_PrinUsuariosForm(forms.ModelForm):

    org_int = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))
    pres_rep = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))
    misnit = forms.CharField(required=False, widget=forms.TextInput(attrs={
                             'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))
    org_cont = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))
    o_ent_o_nac = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))
    ent_o_terr = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))
    gremios = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))
    ent_privadas = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))
    dep_misma_entidad = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))
    academia = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MB_PrinUsuarios
        fields = ['org_int', 'pres_rep', 'misnit', 'org_cont', 'o_ent_o_nac', 'ent_o_terr',  'gremios',
                  'ent_privadas', 'dep_misma_entidad', 'academia'
                  ]
    
    def clean(self):
        
        org_int = self.cleaned_data.get('org_int')
        pres_rep = self.cleaned_data.get('pres_rep')
        misnit = self.cleaned_data.get('misnit')
        org_cont = self.cleaned_data.get('org_cont')
        o_ent_o_nac = self.cleaned_data.get('o_ent_o_nac')
        ent_o_terr = self.cleaned_data.get('ent_o_terr')
        gremios = self.cleaned_data.get('gremios')
        ent_privadas = self.cleaned_data.get('ent_privadas')
        dep_misma_entidad = self.cleaned_data.get('dep_misma_entidad')
        academia = self.cleaned_data.get('academia')
            
        if org_int:
            if len(org_int) > 2000: 
                self._errors['org_int'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        if pres_rep:
            if len(pres_rep) > 2000: 
                self._errors['pres_rep'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        if misnit:
            if len(misnit) > 2000: 
                self._errors['misnit'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])
        
        if org_cont:
            if len(org_cont) > 2000: 
                self._errors['org_cont'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        if o_ent_o_nac:
            if len(o_ent_o_nac) > 2000: 
                self._errors['o_ent_o_nac'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])
        
        if ent_o_terr:
            if len(ent_o_terr) > 2000: 
                self._errors['ent_o_terr'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        if gremios:
            if len(gremios) > 2000: 
                self._errors['gremios'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        if ent_privadas:
            if len(ent_privadas) > 2000: 
                self._errors['ent_privadas'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])
          
        if dep_misma_entidad:
            if len(dep_misma_entidad) > 2000: 
                self._errors['dep_misma_entidad'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        if academia:
            if len(academia) > 2000: 
                self._errors['academia'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])
        
        # return any errors if found 
        return(self.cleaned_data)


# Modulo C pregunta 2

class EditMC_UnidadObservacionForm(forms.ModelForm):
    mc_otra = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MC_UnidadObservacion
        fields = ['mc_otra']

    def clean(self):
        
        mc_otra = self.cleaned_data.get('mc_otra')
        
        if mc_otra:
            if len(mc_otra) > 2500: 
                self._errors['mc_otra'] = self.error_class([ 
                    'Máximo 2500 caracteres requeridos'])

         # return any errors if found 
        return(self.cleaned_data)

# Modulo C pregunta 4


class EditMC_ObtencionDatoForm(forms.ModelForm):
    
    mc_ra_cual = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Enumere el registro administrativo', 'class': 'form-control form-control-sm mb-1medio'}))
    
    mc_ra_entidad = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Enumere la Entidad que realiza el registro', 'class': 'form-control form-control-sm'}))
    
    mc_oe_cual = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Enumere la operación estadística', 'class': 'form-control form-control-sm mb-1medio'}))
    
    mc_oe_entidad = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Enumere la Entidad que realiza la operación', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MC_ObtencionDato
        fields = ['mc_ra_cual',  'mc_ra_entidad',
                  'mc_oe_cual',  'mc_oe_entidad']

    def clean(self):
            
        mc_ra_cual = self.cleaned_data.get('mc_ra_cual') 
        mc_ra_entidad = self.cleaned_data.get('mc_ra_entidad')
        mc_oe_cual = self.cleaned_data.get('mc_oe_cual')
        mc_oe_entidad = self.cleaned_data.get('mc_oe_entidad')
        
        if mc_ra_cual:
            if len(mc_ra_cual) > 2000: 
                self._errors['mc_ra_cual'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])
        
        if mc_ra_entidad:
            if len(mc_ra_entidad) > 2000: 
                self._errors['mc_ra_entidad'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        if mc_oe_cual:
            if len(mc_oe_cual) > 2000: 
                self._errors['mc_oe_cual'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        if mc_oe_entidad:
            if len(mc_oe_entidad) > 2000: 
                self._errors['mc_oe_entidad'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)


# Modulo C pregunta 5 probabilistico

class EditMC_MuestreoProbabilisticoForm(forms.ModelForm):
    prob_otro = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MC_MuestreoProbabilistico
        fields = ['prob_otro']

    def clean(self):
            
        prob_otro = self.cleaned_data.get('prob_otro')
        
        if prob_otro:
            if len(prob_otro) > 150: 
                self._errors['prob_otro'] = self.error_class([ 
                    'Máximo 150 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)


# Modulo C pregunta 5  No probabilistico

class EditMC_MuestreoNoProbabilisticoForm(forms.ModelForm):
    no_prob_otro = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MC_MuestreoNoProbabilistico
        fields = ['no_prob_otro']

    def clean(self):
            
        no_prob_otro = self.cleaned_data.get('no_prob_otro')
        
        if no_prob_otro:
            if len(no_prob_otro) > 500: 
                self._errors['no_prob_otro'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)


# Modulo C pregunta 6
class EditMC_TipoMarcoForm(forms.ModelForm):
    otro_tipo_marco = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MC_TipoMarco
        fields = ['otro_tipo_marco']

    def clean(self):
            
        otro_tipo_marco = self.cleaned_data.get('otro_tipo_marco')
        
        if otro_tipo_marco:
            if len(otro_tipo_marco) > 800: 
                self._errors['otro_tipo_marco'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)


# Modulo C pregunta 7
class EditMC_DocsDesarrolloForm(forms.ModelForm):
    otro_docs = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MC_DocsDesarrollo
        fields = ['otro_docs']

    def clean(self):
            
        otro_docs = self.cleaned_data.get('otro_docs')
        
        if otro_docs:
            if len(otro_docs) > 800: 
                self._errors['otro_docs'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)


# Modulo C pregunta 8
class EditMC_ConceptosEstandarizadosForm(forms.ModelForm):
    org_in_cuales = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm mb-1medio'}))
    ent_ordnac_cuales = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm mb-1medio'}))
    leye_dec_cuales = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm mb-1medio'}))
    otra_cual_conp = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm mb-1medio'}))
    ningu_pq = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Por qué?', 'class': 'form-control form-control-sm mb-1medio'}))

    class Meta:
        model = MC_ConceptosEstandarizados
        fields = ['org_in_cuales', 'ent_ordnac_cuales',
                  'leye_dec_cuales', 'otra_cual_conp', 'ningu_pq']

    def clean(self):
            
        org_in_cuales = self.cleaned_data.get('org_in_cuales')
        ent_ordnac_cuales = self.cleaned_data.get('ent_ordnac_cuales')
        leye_dec_cuales = self.cleaned_data.get('leye_dec_cuales')
        otra_cual_conp = self.cleaned_data.get('otra_cual_conp')
        ningu_pq = self.cleaned_data.get('otra_cual_conp')
        
        if org_in_cuales:
            if len(org_in_cuales) > 500: 
                self._errors['org_in_cuales'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])

        if ent_ordnac_cuales:
            if len(ent_ordnac_cuales) > 500: 
                self._errors['ent_ordnac_cuales'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])

        if leye_dec_cuales:
            if len(leye_dec_cuales) > 500: 
                self._errors['leye_dec_cuales'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])

        if otra_cual_conp:
            if len(otra_cual_conp) > 500: 
                self._errors['otra_cual_conp'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])

        if ningu_pq:
            if len(ningu_pq) > 500: 
                self._errors['ningu_pq'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])
        
        # return any errors if found 
        return(self.cleaned_data)


# Modulo C pregunta 9
class EditMC_ClasificacionesForm(forms.ModelForm):
    otra_cual_clas = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otra', 'class': 'form-control form-control-sm mb-1medio w-75 ml-3'}))
    no_pq = forms.CharField(required=False, widget=forms.TextInput(attrs={
                            'placeholder': '¿Por qué?', 'class': 'form-control form-control-sm mb-1medio'}))

    class Meta:
        model = MC_Clasificaciones
        fields = ['otra_cual_clas', 'no_pq']

    def clean(self):
            
        otra_cual_clas  = self.cleaned_data.get('otra_cual_clas')
        no_pq  = self.cleaned_data.get('no_pq')
        
        if otra_cual_clas:
            if len(otra_cual_clas) > 800: 
                self._errors['otra_cual_clas'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        if no_pq:
            if len(no_pq) > 800: 
                self._errors['no_pq'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)


# Modulo C pregunta 10
class EditMC_CoberturaGeograficaForm(forms.ModelForm):
    tot_regional = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuántos?', 'class': 'form-control form-control-sm'}))
    cual_regional = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm heigthTextArea'}))
    tot_dep = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuántos?', 'class': 'form-control form-control-sm'}))
    cual_dep = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm heigthTextArea'}))
    tot_are_metr = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuántos?', 'class': 'form-control form-control-sm'}))
    cual_are_metr = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm heigthTextArea'}))
    tot_mun = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuántos?', 'class': 'form-control form-control-sm'}))
    cual_mun = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm heigthTextArea'}))
    tot_otro = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuántos?', 'class': 'form-control form-control-sm'}))
    cual_otro = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm heigthTextArea'}))

    class Meta:
        model = MC_CoberturaGeografica
        fields = ['tot_regional', 'cual_regional',  'tot_dep', 'cual_dep',
                  'tot_are_metr', 'cual_are_metr', 'tot_mun', 'cual_mun', 'tot_otro', 'cual_otro']

    
    def clean(self):
            
        tot_regional = self.cleaned_data.get('tot_regional')
        cual_regional  = self.cleaned_data.get('cual_regional')
        tot_dep  = self.cleaned_data.get('tot_dep')
        cual_dep  = self.cleaned_data.get('cual_dep')
        tot_are_metr  = self.cleaned_data.get('tot_are_metr')
        cual_are_metr  = self.cleaned_data.get('cual_are_metr')
        tot_mun  = self.cleaned_data.get('tot_mun')
        cual_mun  = self.cleaned_data.get('cual_mun')
        tot_otro  = self.cleaned_data.get('tot_otro')
        cual_otro  = self.cleaned_data.get('cual_otro')
        
        if tot_regional:
            if len(tot_regional) > 40: 
                self._errors['tot_regional'] = self.error_class([ 
                    'Máximo 40 caracteres requeridos'])

        if cual_regional:
            if len(cual_regional) > 600: 
                self._errors['cual_regional'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])
        
        if tot_dep:
            if len(tot_dep) > 40: 
                self._errors['tot_dep'] = self.error_class([ 
                    'Máximo 40 caracteres requeridos'])

        if cual_dep:
            if len(cual_dep) > 600: 
                self._errors['cual_dep'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])
        
        if tot_are_metr:
            if len(tot_are_metr) > 40: 
                self._errors['tot_are_metr'] = self.error_class([ 
                    'Máximo 40 caracteres requeridos'])

        if cual_are_metr:
            if len(cual_are_metr) > 600: 
                self._errors['cual_are_metr'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])
        
        if tot_mun:
            if len(tot_mun) > 40: 
                self._errors['tot_mun'] = self.error_class([ 
                    'Máximo 40 caracteres requeridos'])

        if cual_mun:
            if len(cual_mun) > 600: 
                self._errors['cual_mun'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])

        if tot_otro:
            if len(tot_otro) > 600: 
                self._errors['tot_otro'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])

        if cual_otro:
            if len(cual_otro) > 600: 
                self._errors['cual_otro'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])
        
        # return any errors if found 
        return(self.cleaned_data)


# Modulo C pregunta 11

class EditMC_DesagregacionInformacionForm(forms.ModelForm):
    des_tot_regional = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuántos?', 'class': 'form-control form-control-sm'}))
    des_cual_regional = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))
    des_tot_dep = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuántos?', 'class': 'form-control form-control-sm'}))
    des_cual_dep = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))
    des_tot_are_metr = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuántos?', 'class': 'form-control form-control-sm'}))
    des_cual_are_metr = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))
    des_tot_mun = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuántos?', 'class': 'form-control form-control-sm'}))
    des_cual_mun = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))
    des_tot_otro = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuántos?', 'class': 'form-control form-control-sm'}))
    des_cual_otro = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))
    des_grupo_otro = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MC_DesagregacionInformacion
        fields = ['des_tot_regional', 'des_cual_regional',  'des_tot_dep', 'des_cual_dep', 'des_tot_are_metr', 'des_cual_are_metr', 'des_tot_mun', 'des_cual_mun',
                  'des_tot_otro', 'des_cual_otro', 'des_grupo_otro']

    def clean(self):
            
        des_tot_regional = self.cleaned_data.get('des_tot_regional')
        des_cual_regional = self.cleaned_data.get('des_cual_regional')
        des_tot_dep = self.cleaned_data.get('des_tot_dep')
        des_cual_dep = self.cleaned_data.get('des_cual_dep')
        des_tot_are_metr = self.cleaned_data.get('des_tot_are_metr')
        des_cual_are_metr = self.cleaned_data.get('des_cual_are_metr')
        des_tot_mun = self.cleaned_data.get('des_tot_mun')
        des_cual_mun = self.cleaned_data.get('des_cual_mun')
        des_tot_otro = self.cleaned_data.get('des_tot_otro')
        des_cual_otro = self.cleaned_data.get('des_cual_otro')
        des_grupo_otro = self.cleaned_data.get('des_grupo_otro')
        
        if des_tot_regional:
            if len(des_tot_regional) > 40: 
                self._errors['des_tot_regional'] = self.error_class([ 
                    'Máximo 40 caracteres requeridos'])

        if des_cual_regional:
            if len(des_cual_regional) > 600: 
                self._errors['des_cual_regional'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])

        if des_tot_dep:
            if len(des_tot_dep) > 40: 
                self._errors['des_tot_dep'] = self.error_class([ 
                    'Máximo 40 caracteres requeridos'])

        if des_cual_dep:
            if len(des_cual_dep) > 600: 
                self._errors['des_cual_dep'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])

        if des_tot_are_metr:
            if len(des_tot_are_metr) > 40: 
                self._errors['des_tot_are_metr'] = self.error_class([ 
                    'Máximo 40 caracteres requeridos'])

        if des_cual_are_metr:
            if len(des_cual_are_metr) > 600: 
                self._errors['des_cual_are_metr'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])

        if des_tot_mun:
            if len(des_tot_mun) > 40: 
                self._errors['des_tot_mun'] = self.error_class([ 
                    'Máximo 40 caracteres requeridos'])

        if des_cual_mun:
            if len(des_cual_mun) > 600: 
                self._errors['des_cual_mun'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])

        if des_tot_otro:
            if len(des_tot_otro) > 40: 
                self._errors['des_tot_otro'] = self.error_class([ 
                    'Máximo 40 caracteres requeridos'])

        if des_cual_otro:
            if len(des_cual_otro) > 600: 
                self._errors['des_cual_otro'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])       

        if des_grupo_otro:
            if len(des_grupo_otro) > 600: 
                self._errors['des_grupo_otro'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])       

        # return any errors if found 
        return(self.cleaned_data)


# Modulo C pregunta 13
class EditMC_FuenteFinanciacionForm(forms.ModelForm):
    r_otros = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MC_FuenteFinanciacion
        fields = ['r_otros']

    def clean(self):
            
        r_otros = self.cleaned_data.get('r_otros')
       
        if r_otros:
            if len(r_otros) > 500: 
                self._errors['r_otros'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)


# Modulo D pregunta 1


class EditMD_MedioDatosForm(forms.ModelForm):
    sis_info = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm'}))
    md_otro = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MD_MedioDatos
        fields = ['sis_info',  'md_otro']

    def clean(self):
            
        sis_info = self.cleaned_data.get('sis_info')
        md_otro = self.cleaned_data.get('md_otro')
       
        if sis_info:
            if len(sis_info) > 300: 
                self._errors['sis_info'] = self.error_class([ 
                    'Máximo 300 caracteres requeridos'])

        if md_otro:
            if len(md_otro) > 300: 
                self._errors['md_otro'] = self.error_class([ 
                    'Máximo 300 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)


# Modulo D pregunta 2

class EditMD_PeriodicidadOeForm(forms.ModelForm):
    per_otro = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MD_PeriodicidadOe
        fields = ['per_otro']

    def clean(self):
            
        per_otro = self.cleaned_data.get('per_otro')
    
        if per_otro:
            if len(per_otro) > 300: 
                self._errors['per_otro'] = self.error_class([ 
                    'Máximo 300 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)

# Modulo D pregunta 3


class EditMD_HerramProcesamientoForm(forms.ModelForm):
    herr_otro = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MD_HerramProcesamiento
        fields = ['herr_otro']

    def clean(self):
            
        herr_otro = self.cleaned_data.get('herr_otro')
    
        if herr_otro:
            if len(herr_otro) > 300: 
                self._errors['herr_otro'] = self.error_class([ 
                    'Máximo 300 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)

# modulo E pregunta 1


class EditME_AnalisisResultadosForm(forms.ModelForm):
    ana_otro = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = ME_AnalisisResultados
        fields = ['ana_otro']

    def clean(self):
            
        ana_otro = self.cleaned_data.get('ana_otro')
    
        if ana_otro:
            if len(ana_otro) > 300: 
                self._errors['ana_otro'] = self.error_class([ 
                    'Máximo 300 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)

# Modulo F pregunta 1


class EditMF_MediosDifusionForm(forms.ModelForm):
    dif_otro = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MF_MediosDifusion
        fields = ['dif_otro']

    def clean(self):
            
        dif_otro = self.cleaned_data.get('dif_otro')
    
        if dif_otro:
            if len(dif_otro) > 300: 
                self._errors['dif_otro'] = self.error_class([ 
                    'Máximo 300 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)


# Modulo F pregunta 4
class EditMF_FechaPublicacionForm(forms.ModelForm):
    fecha = forms.DateField(
        required=False,
        widget=forms.DateInput(format='%Y-%m-%d',
                               attrs={'class': 'form-control datepicker'}),
        input_formats=('%Y-%m-%d', )
        )
    no_hay = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Por qué?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MF_FechaPublicacion
        fields = ['fecha', 'no_hay']

    def clean(self):
            
        no_hay = self.cleaned_data.get('no_hay')
    
        if no_hay:
            if len(no_hay) > 150: 
                self._errors['no_hay'] = self.error_class([ 
                    'Máximo 150 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)


# Modulo F pregunta 5

class EditMF_FrecuenciaDifusionForm(forms.ModelForm):
    no_definido = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MF_FrecuenciaDifusion
        fields = ['no_definido']

    def clean(self):
            
        no_definido = self.cleaned_data.get('no_definido')
    
        if no_definido:
            if len(no_definido) > 300: 
                self._errors['no_definido'] = self.error_class([ 
                    'Máximo 300 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)

# Modulo F pregunta 6


class EditMF_ProductosDifundirForm(forms.ModelForm):
    difundir_otro = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MF_ProductosDifundir
        fields = ['difundir_otro']

    def clean(self):
            
        difundir_otro = self.cleaned_data.get('difundir_otro')
    
        if difundir_otro:
            if len(difundir_otro) > 500: 
                self._errors['difundir_otro'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)

# Modulo F pregunta 7


class EditMF_OtrosProductosForm(forms.ModelForm):
    ser_hist_desde = forms.DateField(
        required=False,
        widget=forms.DateInput(format='%Y-%m-%d',
                               attrs={'class': 'form-control datepicker'}),
        input_formats=('%Y-%m-%d', )
        )
    ser_hist_hasta = forms.DateField(
        required=False,
        widget=forms.DateInput(format='%Y-%m-%d',
                               attrs={'class': 'form-control datepicker'}),
        input_formats=('%Y-%m-%d', )
        )
    microdatos_desde = forms.DateField(
        required=False,
        widget=forms.DateInput(format='%Y-%m-%d',
                               attrs={'class': 'form-control datepicker'}),
        input_formats=('%Y-%m-%d', )
        )
    microdatos_hasta = forms.DateField(
        required=False,
        widget=forms.DateInput(format='%Y-%m-%d',
                               attrs={'class': 'form-control datepicker'}),
        input_formats=('%Y-%m-%d', )
        )
    op_url = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'URL?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MF_OtrosProductos
        fields = ['ser_hist_desde', 'ser_hist_hasta',
                  'microdatos_desde', 'microdatos_hasta', 'op_url']

    def clean(self):
            
        op_url = self.cleaned_data.get('op_url')

        if op_url:
            if len(op_url) > 300: 
                self._errors['op_url'] = self.error_class([ 
                    'Máximo 300 caracteres requeridos'])
         # return any errors if found 
        return(self.cleaned_data)

# Modulo F pregunta 8
class EditMF_ResultadosSimilaresForm(forms.ModelForm):
    rs_entidad = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Entidad', 'class': 'form-control form-control-sm'}))
    rs_oe = forms.CharField(required=False, widget=forms.TextInput(attrs={
                            'placeholder': 'Operación estadística / Indicadores', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MF_ResultadosSimilares
        fields = ['rs_entidad', 'rs_oe']

    def clean(self):
            
        rs_entidad = self.cleaned_data.get('rs_entidad')
        rs_oe = self.cleaned_data.get('rs_oe')
    
        if rs_entidad:
            if len(rs_entidad) > 500: 
                self._errors['rs_entidad'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])

        if rs_oe:
            if len(rs_oe) > 500: 
                self._errors['rs_oe'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)


# Modulo F pregunta 9
class EditMF_HPSistemaInfoForm(forms.ModelForm):
    si_cual = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál(es)?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MF_HPSistemaInfo
        fields = ['si_cual']

    def clean(self):
            
        si_cual = self.cleaned_data.get('si_cual')
    
        if si_cual:
            if len(si_cual) > 500: 
                self._errors['si_cual'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)


#### formulario para los comentarios  ##############
class CommentForm(forms.ModelForm):
    body = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Describa aquí las inconsistencias sobre la información diligenciada en el formulario', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = Comment
        fields = ['body']

    def clean(self):
            
        body = self.cleaned_data.get('body')
    
        if body:
            if len(body) > 1000:
                self._errors['body'] = self.error_class([ 
                    'Máximo 1000 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data) 

#### end formulario para los comentarios  ##############


# formulario para actualización novedades

class NovedadForm(forms.ModelForm):

    descrip_novedad = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Descripción de la novedad', 'class': 'form-control form-control-sm'}))
    
    novedad = forms.ModelChoiceField(required=False, label="tipo de novedad",  queryset=TipoNovedad.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))
    
    est_actualiz = forms.ModelChoiceField(required=False, label="estado de actualización",  queryset=EstadoActualizacion.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))

    class Meta:
        model = NovedadActualizacion
        fields = ['descrip_novedad', 'novedad', 'est_actualiz']


    def clean(self):
            
        descrip_novedad = self.cleaned_data.get('descrip_novedad')
    
        if descrip_novedad:
            if len(descrip_novedad) > 1000: 
                self._errors['descrip_novedad'] = self.error_class([ 
                    'Máximo 1000 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)


class CriticaForm(forms.ModelForm):

    descrip_critica = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Observaciones de la crítica', 'class': 'form-control form-control-sm'}))
    estado_crit = forms.ModelChoiceField(label="tipo de novedad", empty_label=None, queryset=EstadoCritica.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))

    class Meta:
        model = Critica
        fields = ['descrip_critica', 'estado_crit']

    def clean(self):
            
        descrip_critica = self.cleaned_data.get('descrip_critica')
    
        if descrip_critica:
            if len(descrip_critica) > 1000: 
                self._errors['descrip_critica'] = self.error_class([ 
                    'Máximo 1000 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)

#### end formulario para actualización novedades  ##############

## formulario evaluación de calidad #########

CHOICES_METODOLOGIA = (
    (None, '------------'),
    ('ntcpe 1000', 'NTCPE 1000'),
    ('matriz de requisitos', 'MATRIZ DE REQUISITOS'),
)


class EvaluacionCalidadForm(forms.ModelForm):

    est_evaluacion = forms.ModelChoiceField(required=False, label="Estado de la evaluación", queryset=EstadoEvaluacion.objects.all(
        ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))
    
    observ_est = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Escriba sus observaciones relacionadas con el estado de la evaluación.', 'class': 'form-control form-control-sm'}))

    res_evaluacion = forms.ModelChoiceField(required=False, label="Resultado de la evaluación NTCPE 1000",  queryset=ResultadoEvaluacion.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))
    
    res_mzrequi = forms.ModelChoiceField(required=False, label="Resultado de la evaluación MATRIZ DE REQUISITOS",  queryset=ResultadoEvalmatriz.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))
    
    observ_resul = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Observaciones relacionadas con el resultado de la evaluación', 'class': 'form-control form-control-sm'}))
    
    metodologia = forms.ChoiceField(required=False, label="Metodologia", choices=CHOICES_METODOLOGIA, widget=forms.Select(
        attrs={'class': 'form-control form-control-sm'}))
    
    pla_mejoramiento = forms.ModelChoiceField(required=False, label="Plan de mejoramiento",  queryset=PlanDeMejoramiento.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))
    
    seg_vig = forms.ModelChoiceField(required=False, label="Seguimiento anual (Vigilancia)",  queryset=SeguimientoAnual.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))

    obs_seg_anual = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Escriba sus observaciones relacionadas con el estado del seguimiento.', 'class': 'form-control form-control-sm'}))
    
    year_eva = forms.DateTimeField(
        input_formats=['%Y'],
        required=False,
        widget=forms.DateTimeInput(attrs={
            'class': 'form-control datetimepicker-input form-control-sm',
            'data-target': '##evaluacion_year'
        })
    )
    
    vigencia_desde = forms.DateTimeField(
        input_formats=['%d/%m/%Y'],
        required=False,
        widget=forms.DateTimeInput(attrs={
            'class': 'form-control datetimepicker-input form-control-sm',
            'data-target': '##evaluacion_vigencia_desde'
        })
    )
    
    vigencia_hasta = forms.DateTimeField(
        input_formats=['%d/%m/%Y'],
        required=False,
        widget=forms.DateTimeInput(attrs={
            'class': 'form-control datetimepicker-input form-control-sm',
            'data-target': '##evaluacion_vigencia_hasta'
        })
    )

    class Meta:
        model = EvaluacionCalidad
        fields = ['est_evaluacion', 'observ_est', 'res_evaluacion', 'res_mzrequi',  'observ_resul', 'metodologia', 'pla_mejoramiento',
        'seg_vig', 'obs_seg_anual', 'year_eva', 'vigencia_desde', 'vigencia_hasta']


    def clean(self):
            
        observ_est = self.cleaned_data.get('observ_est')
        observ_resul = self.cleaned_data.get('observ_resul')
        obs_seg_anual = self.cleaned_data.get('obs_seg_anual')
    
        if observ_est:
            if len(observ_est) > 600: 
                self._errors['observ_est'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])

        if observ_resul:
            if len(observ_resul) > 600: 
                self._errors['observ_resul'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])

        if obs_seg_anual:
            if len(obs_seg_anual) > 600: 
                self._errors['obs_seg_anual'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)


## End crear formulario evaluación de calidad #########

## Editar formulario evaluación de calidad #########

class EditEvaluacionCalidadForm(forms.ModelForm):

    est_evaluacion = forms.ModelChoiceField(required=False, label="Estado de la evaluación", queryset=EstadoEvaluacion.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))
    
    observ_est = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Escriba sus observaciones relacionadas con el estado de la evaluación.', 'class': 'form-control form-control-sm'}))

    res_evaluacion = forms.ModelChoiceField(required=False, label="Resultado de la evaluación NTCPE 1000",  queryset=ResultadoEvaluacion.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))
    
    res_mzrequi = forms.ModelChoiceField(required=False, label="Resultado de la evaluación MATRIZ DE REQUISITOS",  queryset=ResultadoEvalmatriz.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))
    
    observ_resul = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Observaciones relacionadas con el resultado de la evaluación', 'class': 'form-control form-control-sm'}))
    
    metodologia = forms.ChoiceField(required=False, label="Metodologia", choices=CHOICES_METODOLOGIA, widget=forms.Select(
        attrs={'class': 'form-control form-control-sm'}))
    
    pla_mejoramiento = forms.ModelChoiceField(required=False, label="Plan de mejoramiento",  queryset=PlanDeMejoramiento.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))
    
    seg_vig = forms.ModelChoiceField(required=False, label="Seguimiento anual (Vigilancia)",  queryset=SeguimientoAnual.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))

    obs_seg_anual = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Escriba sus observaciones relacionadas con el estado del seguimiento.', 'class': 'form-control form-control-sm'}))
    
    year_eva = forms.DateTimeField(
        input_formats=['%d/%m/%Y'],
        required=False,
        widget=forms.DateTimeInput(attrs={
            'class': 'form-control datetimepicker-input form-control-sm',
            'data-target': '##evaluacion_year'
        })
    )
    
    vigencia_desde = forms.DateTimeField(
        input_formats=['%d/%m/%Y'],
        required=False,
        widget=forms.DateTimeInput(attrs={
            'class': 'form-control datetimepicker-input form-control-sm',
            'data-target': '##evaluacion_vigencia_desde'
        })
    )
    
    vigencia_hasta = forms.DateTimeField(
        input_formats=['%d/%m/%Y'],
        required=False,
        widget=forms.DateTimeInput(attrs={
            'class': 'form-control datetimepicker-input form-control-sm',
            'data-target': '##evaluacion_vigencia_hasta'
        })
    )

    class Meta:
        model = EvaluacionCalidad
        fields = ['post_oe', 'est_evaluacion', 'observ_est', 'res_evaluacion', 'res_mzrequi','observ_resul', 'metodologia', 'pla_mejoramiento', 
        'seg_vig', 'obs_seg_anual', 'year_eva', 'vigencia_desde', 'vigencia_hasta']


    def clean(self):
            
        observ_est = self.cleaned_data.get('observ_est')
        observ_resul = self.cleaned_data.get('observ_resul')
        obs_seg_anual = self.cleaned_data.get('obs_seg_anual')
    
        if observ_est:
            if len(observ_est) > 600: 
                self._errors['observ_est'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])

        if observ_resul:
            if len(observ_resul) > 600: 
                self._errors['observ_resul'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])

        if obs_seg_anual:
            if len(obs_seg_anual) > 600: 
                self._errors['obs_seg_anual'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)


## end Editar formulario evaluación de calidad #########


