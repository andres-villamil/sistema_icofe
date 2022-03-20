from django import forms
from django.core import validators
from django.core.validators import validate_integer
from django.core.exceptions import ValidationError

from .models import Entidades_oe, RegistroAdministrativo, MB_NormaRRAA, MB_DocumentoMetodologRRAA, MB_VariableRecolectada, MB_ConceptosEstandarizadosRRAA, \
    MB_ClasificacionesRRAA, MB_RecoleccionDatosRRAA, MB_FrecuenciaRecoleccionDato, MB_HerramientasUtilProcesa,  MB_EntidadesAccesoRRAA, \
        MB_SeguridadInform, MB_FrecuenciaAlmacenamientobd, MB_CoberturaGeograficaRRAA, MB_IndicadorResultadoAgregado, \
            MB_NoAccesoDatos, TipoNovedadRRAA, EstadoActualizacionRRAA, SeguimientoPlanFortalecimiento, EstadoCriticaRRAA, \
                EstadoRegistroAd, UsoDeDatosRRAA, AreaTematicaRRAA, TemaRRAA, TemaCompartidoRRAA, NormaRRAA, DocumentoMetodologRRAA, ConceptoEstandarizadoRRAA, \
                    ClasificacionesRRAA, RecoleccionDatoRRAA, FrecuenciaRecoleccionDato, HerramientasUtilProcesa, \
                        SeguridadInform, FrecuenciaAlmacenamientobd, CoberturaGeograficaRRAA, UsoDeDatosRRAA, RelacioneEntidadesAcceso, NoAccesoDatos, \
                            CommentRRAA, NovedadActualizacionRRAA, CriticaRRAA, FortalecimientoRRAA

from entities.models import Entidades_oe
from django.utils.translation import gettext as _
# url validators
from django.core.validators import URLValidator
## formset
from django.forms import modelformset_factory


FILE_TYPES_VARECOL = ['xls', 'xlsx']


class CreateRAForm(forms.ModelForm):

    # A. Identificación
    codigo_rraa = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Código', 'class': 'form-control form-control-sm'}))
    
    entidad_pri = forms.ModelChoiceField(label="Nombre de la entidad", empty_label=None, queryset=Entidades_oe.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm  '}))

    otras_entidades = forms.ModelMultipleChoiceField(
        required=False, widget=forms.SelectMultiple(attrs={"class":"selectpicker", "data-live-search":"true", 'title': 'Seleccione'}), queryset=Entidades_oe.objects.all())

    area_temat = forms.ModelChoiceField(label="Area Temática", empty_label=None, queryset=AreaTematicaRRAA.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm '}))

    tema = forms.ModelChoiceField(label="Tema", empty_label=None, queryset=TemaRRAA.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm '}))
    
    sist_estado = forms.ModelChoiceField(label="Estado", queryset=EstadoRegistroAd.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm', 'onchange':'fondo(this)'})) 
            
    #Dependencia Responsable
    nom_dep = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Dependencia', 'class': 'form-control form-control-sm '}))

    nom_dir = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Director', 'class': 'form-control form-control-sm '}))

    car_dir = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Cargo', 'class': 'form-control form-control-sm '}))

    cor_dir = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Correo electrónico', 'class': 'form-control form-control-sm '}))

    telef_dir = forms.IntegerField(required=False, widget=forms.NumberInput(
        attrs={'placeholder': '999-9999', 'class': 'form-control form-control-sm '}))

    #Temático o Responsable Técnico
    nom_resp = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Responsable', 'class': 'form-control form-control-sm '}))

    carg_resp = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Cargo', 'class': 'form-control form-control-sm '}))

    cor_resp = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Correo electrónico', 'class': 'form-control form-control-sm '}))

    telef_resp = forms.IntegerField(required=False, widget=forms.NumberInput(
        attrs={'placeholder': '999-9999', 'class': 'form-control form-control-sm '}))

    tema_compart = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={'class': ''}), queryset=TemaCompartidoRRAA.objects.all())

    # B. CARACTERIZACIÓN DE REGISTROS ADMINISTRATIVOS

    nombre_ra = forms.CharField(widget=forms.TextInput(
        attrs={'placeholder': 'Nombre del registro administrativo', 'class': 'form-control form-control-sm '}))

    objetivo_ra = forms.CharField(widget=forms.Textarea(
        attrs={'placeholder': 'Objetivo del registro administrativo', 'class': 'form-control form-control-sm '}))

    norma_ra = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={'class': ''}), queryset=NormaRRAA.objects.all())

    pq_secreo = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': '(Describa la razón por la cuál se creó el registro administrativo)', 'class': 'form-control form-control-sm '}))

    fecha_ini_rec = forms.DateField(
        required=False,
        widget=forms.DateInput(format='%Y-%m-%d',
                               attrs={'class': 'form-control datepicker'}),
        input_formats=('%Y-%m-%d', )
        )

    fecha_ult_rec = forms.DateField(
        required=False,
        widget=forms.DateInput(format='%Y-%m-%d',
                               attrs={'class': 'form-control datepicker'}),
        input_formats=('%Y-%m-%d', )
        )

    doc_met_ra = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=DocumentoMetodologRRAA.objects.all())

    variableRecol_file = forms.FileField(
        label='Select a file',
        help_text='max. 42 megabytes',
        required=False,
    )

    con_est_ra = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=ConceptoEstandarizadoRRAA.objects.all())  

    clas_s_n = forms.BooleanField(required=False, initial=False, widget=forms.CheckboxInput(
        attrs={"id": "id_clas_s_n", "data-toggle": "toggle", "data-size": "small", " data-on": "Si",  "data-off": "No", "data-onstyle": "success", "data-offstyle": "danger"}))

    nomb_cla = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=ClasificacionesRRAA.objects.all())   
    
    recole_dato = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=RecoleccionDatoRRAA.objects.all())  
                
    fre_rec_dato = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=FrecuenciaRecoleccionDato.objects.all()) 
    
    herr_u_pro = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=HerramientasUtilProcesa.objects.all()) 
                
    seg_inf = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=SeguridadInform.objects.all())

    almacen_bd_s_n = forms.BooleanField(required=False, initial=False, widget=forms.CheckboxInput(
        attrs={"id": "id_almacen_bd_s_n", "data-toggle": "toggle", "data-size": "small", " data-on": "Si",  "data-off": "No", "data-onstyle": "success", "data-offstyle": "danger"}))

    frec_alm_bd = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=FrecuenciaAlmacenamientobd.objects.all())
                
    cob_geograf = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=CoberturaGeograficaRRAA.objects.all())
                
    uso_de_datos = forms.ModelChoiceField(required=False, queryset=UsoDeDatosRRAA.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))
       
    user_exte_acceso = forms.BooleanField(required=False, initial=False, widget=forms.CheckboxInput(
        attrs={"id": "id_user_exte_acceso", "data-toggle": "toggle", "data-size": "small", " data-on": "Si",  "data-off": "No", "data-onstyle": "success", "data-offstyle": "danger"}))
    
    #pregunta 18 aqui

    no_hay_acceso = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple, queryset=NoAccesoDatos.objects.all())
    
    #modulo c
    observacion = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': '(Indique el número de pregunta si es necesario ampliar, aclarar o complementar la respuesta)', 'class': 'form-control form-control-sm'}))


    class Meta:
        model = RegistroAdministrativo
        fields = [
            'codigo_rraa',
            'entidad_pri',
            'otras_entidades',
            'area_temat',
            'tema', 
            'sist_estado',
            'nom_dep', 
            'nom_dir', 
            'car_dir', 
            'cor_dir', 
            'telef_dir',
            'nom_resp', 
            'carg_resp', 
            'cor_resp',
            'telef_resp',
            'tema_compart', 
            'nombre_ra', 
            'objetivo_ra',
            'norma_ra', 
            'pq_secreo',
            'fecha_ini_rec',
            'fecha_ult_rec',
            'doc_met_ra',
            'variableRecol_file', 
            'con_est_ra',  
            'clas_s_n',
            'nomb_cla', 
            'recole_dato', 
            'fre_rec_dato',
            'herr_u_pro',
            'seg_inf',
            'almacen_bd_s_n', 
            'frec_alm_bd',
            'cob_geograf',     
            'uso_de_datos', 
            'user_exte_acceso',   
            'no_hay_acceso',
            'observacion', 
        ]

    def clean(self):

        nom_dep = self.cleaned_data.get('nom_dep')
        nom_dir = self.cleaned_data.get('nom_dir') 
        car_dir = self.cleaned_data.get('car_dir') 
        nom_resp = self.cleaned_data.get('nom_resp') 
        carg_resp = self.cleaned_data.get('carg_resp') 
        objetivo_ra = self.cleaned_data.get('objetivo_ra')
        pq_secreo = self.cleaned_data.get('pq_secreo')
        observacion = self.cleaned_data.get('observacion')
        
        if nom_dep:
            if len(nom_dep) > 120: 
                self._errors['nom_dep'] = self.error_class([ 
                    'Máximo 120 caracteres requeridos'])
        
        if nom_dir:
            if len(nom_dir) > 80: 
                self._errors['nom_dir'] = self.error_class([ 
                    'Máximo 80 caracteres requeridos'])

        if car_dir:
            if len(car_dir) > 120: 
                self._errors['car_dir'] = self.error_class([ 
                    'Máximo 120 caracteres requeridos'])
        
        if nom_resp:
            if len(nom_resp) > 120: 
                self._errors['nom_resp'] = self.error_class([ 
                    'Máximo 120 caracteres requeridos'])
        
        if carg_resp:
            if len(carg_resp) > 120: 
                self._errors['carg_resp'] = self.error_class([ 
                    'Máximo 120 caracteres requeridos'])

        if objetivo_ra:
            if len(objetivo_ra) > 1500: 
                self._errors['objetivo_ra'] = self.error_class([ 
                    'Máximo 1500 caracteres requeridos'])

        if pq_secreo:
            if len(pq_secreo) > 2000: 
                self._errors['pq_secreo'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        if observacion:
            if len(observacion) > 3000: 
                self._errors['observacion'] = self.error_class([ 
                    'Máximo 3000 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)

    
    def clean_nombre_ra(self):
        nombre_ra = self.cleaned_data['nombre_ra']
        if RegistroAdministrativo.objects.filter(nombre_ra=nombre_ra).exists():
            raise forms.ValidationError(_("El registro administrativo  '%s' ya existe." % nombre_ra))

        if nombre_ra:
            if len(nombre_ra) > 600: 
                self._errors['nombre_ra'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])

       	return nombre_ra


    def clean_telef_dir(self):
        telef_dir = self.cleaned_data['telef_dir']
        if telef_dir:
            if len(str(telef_dir)) < 7 or len(str(telef_dir)) > 10:
                raise forms.ValidationError(
                    _("el número telefónico  que ingreso '%s' debe contener 10 digitos para celular o 7 para teléfono fijo. " % telef_dir))
        return telef_dir

    def clean_telef_resp(self):
        telef_resp = self.cleaned_data['telef_resp']
        if telef_resp:
            if len(str(telef_resp)) < 7 or len(str(telef_resp)) > 10:
                raise forms.ValidationError(
                    _("el número telefónico  que ingreso '%s' debe contener 10 digitos para celular o 7 para teléfono fijo. " % telef_resp))
        return telef_resp

    def clean_cor_dir(self):
        cor_dir = self.cleaned_data['cor_dir']
        if cor_dir:
            if "@" not in cor_dir:
                raise forms.ValidationError(
                    _("El correo electrónico que ingreso '%s' debe contener @." % cor_dir))
            if  len(cor_dir) > 80:
                raise forms.ValidationError(
                    _("El correo electrónico que ingreso '%s' debe contener menos de 80 caracteres" % cor_dir))
        return cor_dir

    def clean_cor_resp(self):
        cor_resp = self.cleaned_data['cor_resp']
        if cor_resp:
            if "@" not in cor_resp:
                raise forms.ValidationError(
                    _("El correo electrónico que ingreso '%s' debe contener @." % cor_resp))
            if  len(cor_resp) > 80:
                raise forms.ValidationError(
                    _("El correo electrónico que ingreso '%s' debe contener menos de 80 caracteres" % cor_resp))
        return cor_resp


    def clean_variableRecol_file(self):
        variableRecol_file = self.cleaned_data['variableRecol_file']
        if variableRecol_file:
            file_type = str(variableRecol_file).split('.')[-1]
            file_type = file_type.lower()
            if file_type not in FILE_TYPES_VARECOL:
                raise forms.ValidationError(
                    _("Archivo No valido '%s' la extensión debe ser XLS o XLSX" % variableRecol_file))

        return variableRecol_file


## Modulo B pregunta 3
class MB_NormaRRAAForm(forms.ModelForm):

    cp_ra = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': 'Describa el requerimiento', 'class': 'form-control form-control-sm mb-1'}))  
    
    ley_ra = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': 'Describa el requerimiento', 'class': 'form-control form-control-sm mb-1'}))
                      
    decreto_ra = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': 'Describa el requerimiento', 'class': 'form-control form-control-sm mb-1'}))

    otra_ra =  forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': 'Describa el requerimiento', 'class': 'form-control form-control-sm mb-1'}))
    
    class Meta:
        model = MB_NormaRRAA
        fields = ['cp_ra', 'ley_ra', 'decreto_ra', 'otra_ra']

    def clean(self):

        cp_ra = self.cleaned_data.get('cp_ra')
        ley_ra = self.cleaned_data.get('ley_ra')
        decreto_ra = self.cleaned_data.get('ley_ra')
        otra_ra = self.cleaned_data.get('ley_ra')
    
        if cp_ra:
            if len(cp_ra) > 2000: 
                self._errors['cp_ra'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        if ley_ra:
            if len(ley_ra) > 2000: 
                self._errors['ley_ra'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        if decreto_ra:
            if len(decreto_ra) > 2000: 
                self._errors['decreto_ra'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        if otra_ra:
            if len(otra_ra) > 2000: 
                self._errors['otra_ra'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])


## Modulo B pregunta 6
class MB_DocumentoMetodologRRAAForm(forms.ModelForm):
    
    otra_doc_cual = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm mb-1'}))

    class Meta:
        model = MB_DocumentoMetodologRRAA
        fields = ['otra_doc_cual']

    def clean(self):

        otra_doc_cual = self.cleaned_data.get('otra_doc_cual')
        
        if otra_doc_cual:
            if len(otra_doc_cual) > 2000: 
                self._errors['otra_doc_cual'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])


# modulo B pregunta 7  ***formsets
class MB_VariableRecolectadaForm(forms.ModelForm):

    variableRec = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'title': 'Módulo B Pregunta 7: Indique cuáles son las variables recolectadas con el registro administrativo','placeholder': 'Indique cuáles son las variables recolectadas con el registro administrativo', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MB_VariableRecolectada
        fields = [
            'variableRec'
        ]

    def clean(self):

        variableRec = self.cleaned_data.get('variableRec')
        
        if variableRec:
            if len(variableRec) > 250: 
                self._errors['variableRec'] = self.error_class([ 
                    'Máximo 250 caracteres requeridos'])

         # return any errors if found 
        return(self.cleaned_data)

# modulo B pregunta 7  ***formsets
MB_VariableRecolectadaFormset = modelformset_factory(
    MB_VariableRecolectada,
    form=MB_VariableRecolectadaForm,
    fields=('variableRec',),
    can_delete=True,
    extra=1,
)

    

# modulo B pregunta 8
class MB_ConceptosEstandarizadosRRAAForm(forms.ModelForm):
    
    org_in_cual = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm mb-1'}))

    ent_ordnac_cual = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm mb-1'}))

    leye_dec_cual = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm mb-1'}))

    otra_ce_cual = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm mb-1'}))

    class Meta:
        model = MB_ConceptosEstandarizadosRRAA
        fields = ['org_in_cual', 'ent_ordnac_cual', 'leye_dec_cual', 'otra_ce_cual' ]

    def clean(self):

        org_in_cual = self.cleaned_data.get('org_in_cual')
        ent_ordnac_cual = self.cleaned_data.get('ent_ordnac_cual')
        leye_dec_cual = self.cleaned_data.get('leye_dec_cual')
        otra_ce_cual = self.cleaned_data.get('otra_ce_cual')

        
        if org_in_cual:
            if len(org_in_cual) > 1000: 
                self._errors['org_in_cual'] = self.error_class([ 
                    'Máximo 1000 caracteres requeridos'])

        if ent_ordnac_cual:
            if len(ent_ordnac_cual) > 1000: 
                self._errors['ent_ordnac_cual'] = self.error_class([ 
                    'Máximo 1000 caracteres requeridos'])

        if leye_dec_cual:
            if len(leye_dec_cual) > 1000: 
                self._errors['leye_dec_cual'] = self.error_class([ 
                    'Máximo 1000 caracteres requeridos'])

        if otra_ce_cual:
            if len(otra_ce_cual) > 1000: 
                self._errors['otra_ce_cual'] = self.error_class([ 
                    'Máximo 1000 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)


# Modulo B pregunta 9
class MB_ClasificacionesRRAAForm(forms.ModelForm):

    otra_cual_clas = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otra', 'class': 'form-control form-control-sm mb-1medio w-75 ml-3'}))

    no_pq = forms.CharField(required=False, widget=forms.TextInput(attrs={
                            'placeholder': '¿Por qué?', 'class': 'form-control form-control-sm mb-1medio'}))

    class Meta:
        model = MB_ClasificacionesRRAA
        fields = ['otra_cual_clas', 'no_pq']

    def clean(self):
            
        otra_cual_clas  = self.cleaned_data.get('otra_cual_clas')
        no_pq  = self.cleaned_data.get('no_pq')
        
        if otra_cual_clas:
            if len(otra_cual_clas) > 600: 
                self._errors['otra_cual_clas'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])

        if no_pq:
            if len(no_pq) > 600: 
                self._errors['no_pq'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)
 


# Modulo B pregunta 10
class MB_RecoleccionDatosRRAAForm(forms.ModelForm):

    sistema_inf_cual = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm mb-1'}))

    otro_c = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': 'Otro (s) ¿cuál(es)?', 'class': 'form-control form-control-sm mb-1'}))

    class Meta:
        model = MB_RecoleccionDatosRRAA
        fields = ['sistema_inf_cual', 'otro_c']

    def clean(self):

        sistema_inf_cual = self.cleaned_data.get('sistema_inf_cual')
        otro_c = self.cleaned_data.get('otro_c')
        
        if sistema_inf_cual:
            if len(sistema_inf_cual) > 500: 
                self._errors['sistema_inf_cual'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])

        if otro_c:
            if len(otro_c) > 500: 
                self._errors['otro_c'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])


        # return any errors if found 
        return(self.cleaned_data)

# Modulo B pregunta 11
class MB_FrecuenciaRecoleccionDatoForm(forms.ModelForm):

    otra_cual_fre = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿cuál(es)?', 'class': 'form-control form-control-sm mb-1'}))

    class Meta:
        model = MB_FrecuenciaRecoleccionDato
        fields = ['otra_cual_fre']

    def clean(self):

        otra_cual_fre = self.cleaned_data.get('otra_cual_fre')
        
        if otra_cual_fre:
            if len(otra_cual_fre) > 800: 
                self._errors['otra_cual_fre'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)

# Modulo B pregunta 12
class MB_HerramientasUtilProcesaForm(forms.ModelForm):

    otra_herram = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿cuál(es)?', 'class': 'form-control form-control-sm mb-1'}))

    class Meta:
        model = MB_HerramientasUtilProcesa
        fields = ['otra_herram']

    def clean(self):

        otra_herram = self.cleaned_data.get('otra_herram')
        
        if otra_herram:
            if len(otra_herram) > 800: 
                self._errors['otra_herram'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        return(self.cleaned_data)


# Modulo B pregunta 13
class MB_SeguridadInformForm(forms.ModelForm):

    otra_cual_s =  forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿cuál?', 'class': 'form-control form-control-sm mb-1'}))

    class Meta:
        model = MB_SeguridadInform
        fields = ['otra_cual_s']

    def clean(self):

        otra_cual_s = self.cleaned_data.get('otra_cual_s')
        
        if otra_cual_s:
            if len(otra_cual_s) > 800: 
                self._errors['otra_cual_s'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        return(self.cleaned_data)


# Modulo B pregunta 14
class MB_FrecuenciaAlmacenamientobdForm(forms.ModelForm):
    
    otra_alm_bd = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿cuál?', 'class': 'form-control form-control-sm mb-1'}))

    class Meta:
        model = MB_FrecuenciaAlmacenamientobd
        fields = ['otra_alm_bd']

    def clean(self):

        otra_alm_bd = self.cleaned_data.get('otra_alm_bd')
        
        if otra_alm_bd:
            if len(otra_alm_bd) > 500: 
                self._errors['otra_alm_bd'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])

        return(self.cleaned_data)


    # Modulo B pregunta 15
class MB_CoberturaGeograficaRRAAForm(forms.ModelForm):

    cual_regio = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿Cuál (es)?', 'class': 'form-control form-control-sm mb-1'}))

    cual_depa = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿Cuál (es)?', 'class': 'form-control form-control-sm mb-1'}))

    cual_are_metrop = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿Cuál (es)?', 'class': 'form-control form-control-sm mb-1'}))

    cual_munic = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿Cuál (es)?', 'class': 'form-control form-control-sm mb-1'}))

    cual_otro = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿Cuál (es)?', 'class': 'form-control form-control-sm mb-1'}))

    class Meta:
        model = MB_CoberturaGeograficaRRAA
        fields = ['cual_regio', 'cual_depa', 'cual_are_metrop', 'cual_munic', 'cual_otro' ]

    def clean(self):

        cual_regio = self.cleaned_data.get('cual_regio')
        cual_depa = self.cleaned_data.get('cual_depa')
        cual_are_metrop = self.cleaned_data.get('cual_are_metrop')
        cual_munic = self.cleaned_data.get('cual_munic')
        cual_otro = self.cleaned_data.get('cual_otro')
        
        if cual_regio:
            if len(cual_regio) > 600: 
                self._errors['cual_regio'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])

        if cual_depa:
            if len(cual_depa) > 600: 
                self._errors['cual_depa'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])

        if cual_are_metrop:
            if len(cual_are_metrop) > 600: 
                self._errors['cual_are_metrop'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])
            
        if cual_munic:
            if len(cual_munic) > 1500: 
                self._errors['cual_munic'] = self.error_class([ 
                    'Máximo 1500 caracteres requeridos'])

        if cual_otro:
            if len(cual_otro) > 600: 
                self._errors['cual_otro'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])


        return(self.cleaned_data)


# Modulo B pregunta 16 ***formset

class MB_IndicadorResultadoAgregadoForm(forms.ModelForm):

    ind_res_agre = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuáles resultados se generan?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MB_IndicadorResultadoAgregado
        fields = [
            'ind_res_agre'
        ]

    def clean(self):

        ind_res_agre = self.cleaned_data.get('ind_res_agre')
        
        if ind_res_agre:
            if len(ind_res_agre) > 350: 
                self._errors['ind_res_agre'] = self.error_class([ 
                    'Máximo 350 caracteres requeridos'])

         # return any errors if found 
        return(self.cleaned_data)

# modulo B pregunta 16  ***formsets
MB_IndicadorResultadoAgregadoFormset = modelformset_factory(
    MB_IndicadorResultadoAgregado,
    form= MB_IndicadorResultadoAgregadoForm,
    fields=('ind_res_agre',),
    can_delete=True,
    extra=1,
)



# Modulo B pregunta 18 pendiente ***formsets
class CreateEntidadesAccesoRRAAForm(forms.ModelForm):
   
    nomb_entidad_acc = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'escriba el nombre de la entidad', 'class': 'form-control form-control-sm'}))
    
    otro_cual =  forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otro, ¿cuál?', 'class': 'form-control form-control-sm'}))

    opcion_pr = forms.ModelMultipleChoiceField(required=False, widget=forms.CheckboxSelectMultiple(
        attrs={'class': 'list-group list-group-horizontal-sm selectFase'}), queryset=RelacioneEntidadesAcceso.objects.all())

    class Meta:
        model = MB_EntidadesAccesoRRAA
        fields = ['nomb_entidad_acc',  'otro_cual', 'opcion_pr']


    def clean(self):

        nomb_entidad_acc = self.cleaned_data.get('nomb_entidad_acc')
        otro_cual = self.cleaned_data.get('otro_cual')
        
        if nomb_entidad_acc:
            if len(nomb_entidad_acc) > 300: 
                self._errors['nomb_entidad_acc'] = self.error_class([ 
                    'Máximo 300 caracteres requeridos'])

        if otro_cual:
            if len(otro_cual) > 300: 
                self._errors['otro_cual'] = self.error_class([ 
                    'Máximo 300 caracteres requeridos'])                    

         # return any errors if found 
        return(self.cleaned_data)

MB_EntidadesAccesoFormset = modelformset_factory(
    MB_EntidadesAccesoRRAA,
    form=CreateEntidadesAccesoRRAAForm,
    fields=('nomb_entidad_acc', 'otro_cual', 'opcion_pr',),
    extra=1,
    can_delete=True,
)



# Modulo B pregunta 19

class MB_NoAccesoDatosForm(forms.ModelForm):

    otra_no_acceso = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál (es)?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MB_NoAccesoDatos
        fields = ['otra_no_acceso']

    def clean(self):

        otra_no_acceso = self.cleaned_data.get('otra_no_acceso')
        
        if otra_no_acceso:
            if len(otra_no_acceso) > 800: 
                self._errors['otra_no_acceso'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        return(self.cleaned_data)


#############################  formularios de editar RRAA ##########################################


class EditRAForm(forms.ModelForm):
    
    # A. Identificación
    codigo_rraa = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Código', 'class': 'form-control form-control-sm' }))
    
    entidad_pri = forms.ModelChoiceField(label="Nombre de la entidad", empty_label=None, queryset=Entidades_oe.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm', "title": "A. Identificación: Nombre de entidad"}))

    otras_entidades = forms.ModelMultipleChoiceField(
        required=False, widget=forms.SelectMultiple(attrs={"class":"selectpicker", "data-live-search":"true", 'title': 'A. Identificación: otras entidades responsables'}), queryset=Entidades_oe.objects.all())

    area_temat = forms.ModelChoiceField(label="Area Temática", empty_label=None, queryset=AreaTematicaRRAA.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm', "title": "A. Identificación: Área temática"}))

    tema = forms.ModelChoiceField(label="Tema", empty_label=None, queryset=TemaRRAA.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm', "title": "A. Identificación: Tema"}))
    
    sist_estado = forms.ModelChoiceField(label="Estado", empty_label=None, queryset=EstadoRegistroAd.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm', 'onchange':'fondo(this)', 'title': 'El registro cambio de estado'}))
  
    #Dependencia Responsable
    nom_dep = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Dependencia', 'class': 'form-control form-control-sm', "title": "A. Identificación: Nombre de la dependencia"}))

    nom_dir = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Director', 'class': 'form-control form-control-sm', "title": "A. Identificación: Nombre del director"}))

    car_dir = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Cargo', 'class': 'form-control form-control-sm', "title": "A. Identificación: Cargo del director"}))

    cor_dir = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Correo electrónico', 'class': 'form-control form-control-sm', "title": "A. Identificación: Correo del director"}))

    telef_dir = forms.IntegerField(required=False, widget=forms.NumberInput(
        attrs={'placeholder': '999-9999', 'class': 'form-control form-control-sm', "title": "A. Identificación: Teléfono del director"}))

    #Temático o Responsable Técnico
    nom_resp = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Responsable', 'class': 'form-control form-control-sm', "title": "A. Identificación: Nombre del responsable"}))

    carg_resp = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Cargo', 'class': 'form-control form-control-sm', "title": "A. Identificación: Cargo del responsable"}))

    cor_resp = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Correo electrónico', 'class': 'form-control form-control-sm', "title": "A. Identificación: Correo del responsable"}))

    telef_resp = forms.IntegerField(required=False, widget=forms.NumberInput(
        attrs={'placeholder': '999-9999', 'class': 'form-control form-control-sm', "title": "A. Identificación: Teléfono del responsable"}))

    tema_compart = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={'class': '', "title": "A. Identificación: Tema compartido"}), queryset=TemaCompartidoRRAA.objects.all())

    # B. CARACTERIZACIÓN DE REGISTROS ADMINISTRATIVOS

    nombre_ra = forms.CharField(widget=forms.TextInput(
        attrs={'placeholder': 'Nombre del registro administrativo', 'class': 'form-control form-control-sm', "title" : "Módulo B Pregunta 1: Nombre del registro administrativo"}))

    objetivo_ra = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Objetivo', 'class': 'form-control form-control-sm', "title" : "Módulo B Pregunta 2: Objetivo del registro administrativo"}))

    norma_ra = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title" : "Módulo B Pregunta 3: Bajo cuál(es) normas se soporta la creación del registro administrativo"}), queryset=NormaRRAA.objects.all())

    pq_secreo = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': '(Describa la razón por la cuál se creó el registro administrativo)', 'class': 'form-control form-control-sm', "title" : "Módulo B Pregunta 4: Describa la razón por la cuál se creó el registro administrativo"}))

    fecha_ini_rec = forms.DateField(
        required=False,
        widget=forms.DateInput(format='%Y-%m-%d',
                               attrs={'class': 'form-control datepicker', 
                               "title" : "Módulo B Pregunta 5: Indique desde y hasta cuando se ha recolectado el registro administrativo"}),
        input_formats=('%Y-%m-%d', )
        )

    fecha_ult_rec = forms.DateField(
        required=False,
        widget=forms.DateInput(format='%Y-%m-%d',
                               attrs={'class': 'form-control datepicker', 
                               "title" : "Módulo B Pregunta 5: Indique desde y hasta cuando se ha recolectado el registro administrativo"}),
        input_formats=('%Y-%m-%d', )
        )

    doc_met_ra = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title" : "Módulo B Pregunta 6: Indique desde y hasta cuando se ha recolectado el registro administrativo"}),
        queryset=DocumentoMetodologRRAA.objects.all())

    variableRecol_file = forms.FileField(
        label='Select a file',
        help_text='max. 42 megabytes',
        required=False,
        widget=forms.ClearableFileInput(attrs={"title": "Módulo B pregunta 7: Archivo Indique cuáles son las variables recolectadas con el registro administrativo"})
    )

    con_est_ra = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title" : "Módulo B Pregunta 8: Indique si el registro administrativo utiliza conceptos estandarizados provenientes"}),
        queryset=ConceptoEstandarizadoRRAA.objects.all())

    clas_s_n = forms.BooleanField(required=False, initial=False, widget=forms.CheckboxInput(
        attrs={"id": "id_clas_s_n", "data-toggle": "toggle", "data-size": "small", " data-on": "Si",  "data-off": "No", "data-onstyle": "success", "data-offstyle": "danger",
        "title" : "Módulo B Pregunta 9: El registro administrativo utiliza nomenclaturas y/o clasificaciones"}))

    nomb_cla = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title" : "Módulo B Pregunta 9: El registro administrativo utiliza nomenclaturas y/o clasificaciones"}), queryset=ClasificacionesRRAA.objects.all())   
    
    recole_dato = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title" : "Módulo B Pregunta 10: ¿Cuál es el medio de obtención o recolección de los datos?"}), queryset=RecoleccionDatoRRAA.objects.all())  
                
    fre_rec_dato = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title" : "Módulo B Pregunta 11: ¿Con qué frecuencia se recolectan los datos? "}), queryset=FrecuenciaRecoleccionDato.objects.all()) 
    
    herr_u_pro = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title" : "Módulo B Pregunta 12: Indique cuáles de las siguientes herramientas son utilizadas en el procesamiento de los datos"}), queryset=HerramientasUtilProcesa.objects.all()) 
                
    seg_inf = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title" : "Módulo B Pregunta 13: ¿Con cuáles herramientas cuenta para garantizar la seguridad de la información del registro administrativo?"}), queryset=SeguridadInform.objects.all())

    almacen_bd_s_n = forms.BooleanField(required=False, initial=False, widget=forms.CheckboxInput(
        attrs={"id": "id_almacen_bd_s_n", "data-toggle": "toggle", "data-size": "small", " data-on": "Si",  "data-off": "No", "data-onstyle": "success", "data-offstyle": "danger",
        "title": "Módulo B Pregunta 14: ¿La información recolectada es acopiada o almacenada en una base de datos?"}))

    frec_alm_bd = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title": "Módulo B Pregunta 14: ¿La información recolectada es acopiada o almacenada en una base de datos?"}), queryset=FrecuenciaAlmacenamientobd.objects.all())
                
    cob_geograf = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title": "Módulo B Pregunta 15: ¿Cuál es la cobertura geográfica del registro administrativo?"}), queryset=CoberturaGeograficaRRAA.objects.all())
                
    uso_de_datos = forms.ModelChoiceField(required=False, queryset=UsoDeDatosRRAA.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm', "title": "Módulo B Pregunta 16: ¿La entidad hace uso de los datos del registro administrativo para generar información estadística?" }))
       
    user_exte_acceso = forms.BooleanField(required=False, initial=False, widget=forms.CheckboxInput(
        attrs={"id": "id_user_exte_acceso", "data-toggle": "toggle", "data-size": "small", " data-on": "Si",  "data-off": "No", "data-onstyle": "success", "data-offstyle": "danger",
        "title": "Módulo B Pregunta 17: ¿Usuarios externos a la entidad tienen acceso a los datos del registro administrativo?"}))
    
    no_hay_acceso = forms.ModelMultipleChoiceField(
        required=False, widget=forms.CheckboxSelectMultiple(attrs={"title": "Módulo B Pregunta 19: ¿Cuál es la razón principal por la cual no se permite el acceso a los datos del registro administrativo?"}), queryset=NoAccesoDatos.objects.all())
    
    #modulo c
    observacion = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': '(Indique el número de pregunta si es necesario ampliar, aclarar o complementar la respuesta)', 'class': 'form-control form-control-sm',
        "title": "Módulo C: Observaciones"}))

    #### complemento RRAA

    ra_activo = forms.BooleanField(required=False, initial=False, widget=forms.CheckboxInput(
        attrs={"id": "id_ra_activo", "data-toggle": "toggle", "data-size": "small", " data-on": "Si",  
        "data-off": "No", "data-onstyle": "success", "data-offstyle": "danger", "title": "Campos de control: Registro administrativo activo" }))

    user_dane = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'usuario DANE', 
        'class': 'form-control form-control-sm', "title": "Campos de control: Usuario DANE"}))

    responde_ods = forms.BooleanField(required=False, initial=False, widget=forms.CheckboxInput(
        attrs={"id": "id_responde_ods", "data-toggle": "toggle", "data-size": "small", " data-on": "Si",  
        "data-off": "No", "data-onstyle": "success", "data-offstyle": "danger", "title": "Campos de control: Responde a requerimientos ODS"}))

    indicador_ods =  forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Indicador ODS a que da respuesta', 
        'class': 'form-control form-control-sm', "title": "Indicador ODS a que da respuesta"})) 
    
    class Meta:
        model = RegistroAdministrativo
        fields = [
            'codigo_rraa',
            'entidad_pri',
            'otras_entidades',
            'area_temat',
            'tema', 
            'sist_estado',
            'nom_dep', 
            'nom_dir', 
            'car_dir', 
            'cor_dir', 
            'telef_dir',
            'nom_resp', 
            'carg_resp', 
            'cor_resp',
            'telef_resp', 
            'tema_compart',
            'nombre_ra', 
            'objetivo_ra',
            'norma_ra', 
            'pq_secreo',
            'fecha_ini_rec',
            'fecha_ult_rec',
            'doc_met_ra',
            'variableRecol_file',
            'con_est_ra',  
            'clas_s_n',
            'nomb_cla', 
            'recole_dato', 
            'fre_rec_dato',
            'herr_u_pro',
            'seg_inf',
            'almacen_bd_s_n', 
            'frec_alm_bd',
            'cob_geograf',     
            'uso_de_datos', 
            'user_exte_acceso',   
            #pregunta 18 aqui
            'no_hay_acceso',
            'observacion',
            'ra_activo',
            'user_dane',
            'responde_ods',
            'indicador_ods',
        ]

    def clean(self):

        nom_dep = self.cleaned_data.get('nom_dep')
        nom_dir = self.cleaned_data.get('nom_dir') 
        car_dir = self.cleaned_data.get('car_dir') 
        nom_resp = self.cleaned_data.get('nom_resp') 
        carg_resp = self.cleaned_data.get('carg_resp') 
        nombre_ra = self.cleaned_data.get('nombre_ra') 
        objetivo_ra = self.cleaned_data.get('objetivo_ra')
        pq_secreo = self.cleaned_data.get('pq_secreo')
        observacion = self.cleaned_data.get('observacion')
        user_dane = self.cleaned_data.get('user_dane')
        indicador_ods = self.cleaned_data.get('indicador_ods')
        
        if nom_dep:
            if len(nom_dep) > 120: 
                self._errors['nom_dep'] = self.error_class([ 
                    'Máximo 120 caracteres requeridos'])
        
        if nom_dir:
            if len(nom_dir) > 80: 
                self._errors['nom_dir'] = self.error_class([ 
                    'Máximo 80 caracteres requeridos'])

        if car_dir:
            if len(car_dir) > 120: 
                self._errors['car_dir'] = self.error_class([ 
                    'Máximo 120 caracteres requeridos'])
        
        if nom_resp:
            if len(nom_resp) > 120: 
                self._errors['nom_resp'] = self.error_class([ 
                    'Máximo 120 caracteres requeridos'])
        
        if carg_resp:
            if len(carg_resp) > 120: 
                self._errors['carg_resp'] = self.error_class([ 
                    'Máximo 120 caracteres requeridos'])

        if nombre_ra:
            if len(nombre_ra) > 600: 
                self._errors['nombre_ra'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])

        if objetivo_ra:
            if len(objetivo_ra) > 1500: 
                self._errors['objetivo_ra'] = self.error_class([ 
                    'Máximo 1500 caracteres requeridos'])

        if pq_secreo:
            if len(pq_secreo) > 2000: 
                self._errors['pq_secreo'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        if observacion:
            if len(observacion) > 3000: 
                self._errors['observacion'] = self.error_class([ 
                    'Máximo 3000 caracteres requeridos'])

        if user_dane:
            if len(user_dane) > 500: 
                self._errors['user_dane'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])

        if indicador_ods:
            if len(indicador_ods) > 1000: 
                self._errors['indicador_ods'] = self.error_class([ 
                'Máximo 1000 caracteres requeridos'])
        
    
        # return any errors if found 
        return(self.cleaned_data)


    def clean_telef_dir(self):
        telef_dir = self.cleaned_data['telef_dir']
        if telef_dir:
            if len(str(telef_dir)) < 7  or  len(str(telef_dir)) > 10:
                raise forms.ValidationError(
                    _("el número telefónico  que ingreso '%s' debe contener 10 digitos para celular o 7 para teléfono fijo. " % telef_dir))
        return telef_dir


    def clean_telef_resp(self):
        telef_resp = self.cleaned_data['telef_resp']
        if telef_resp:
            if len(str(telef_resp)) < 7 or len(str(telef_resp)) > 10:
                raise forms.ValidationError(
                    _("el número telefónico  que ingreso '%s' debe contener 10 digitos para celular o 7 para teléfono fijo. " % telef_resp))
        return telef_resp

    def clean_cor_dir(self):
        cor_dir = self.cleaned_data['cor_dir']
        if cor_dir:
            if "@" not in cor_dir:
                raise forms.ValidationError(
                    _("El correo electrónico que ingreso '%s' debe contener @." % cor_dir))
            if  len(cor_dir) > 80:
                raise forms.ValidationError(
                    _("El correo electrónico que ingreso '%s' debe contener menos de 80 caracteres" % cor_dir))
        return cor_dir

    def clean_cor_resp(self):
        cor_resp = self.cleaned_data['cor_resp']
        if cor_resp:
            if "@" not in cor_resp:
                raise forms.ValidationError(
                    _("El correo electrónico que ingreso '%s' debe contener @." % cor_resp))
            if  len(cor_resp) > 80:
                raise forms.ValidationError(
                    _("El correo electrónico que ingreso '%s' debe contener menos de 80 caracteres" % cor_resp))
        return cor_resp


    def clean_variableRecol_file(self):
        variableRecol_file = self.cleaned_data['variableRecol_file']
        if variableRecol_file:
            file_type = str(variableRecol_file).split('.')[-1]
            file_type = file_type.lower()
            if file_type not in FILE_TYPES_VARECOL:
                raise forms.ValidationError(
                    _("Archivo No valido '%s'  la extensión debe ser XLS o XLSX" % variableRecol_file))

        return variableRecol_file


## Modulo B pregunta 3
class EditMB_NormaRRAAForm(forms.ModelForm):

    cp_ra = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': 'Describa el requerimiento', 'class': 'form-control form-control-sm mb-1'}))  
    
    ley_ra = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': 'Describa el requerimiento', 'class': 'form-control form-control-sm mb-1'}))
                      
    decreto_ra = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': 'Describa el requerimiento', 'class': 'form-control form-control-sm mb-1'}))

    otra_ra =  forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': 'Describa el requerimiento', 'class': 'form-control form-control-sm mb-1'}))
    
    class Meta:
        model = MB_NormaRRAA
        fields = ['cp_ra', 'ley_ra', 'decreto_ra', 'otra_ra']

    def clean(self):

        cp_ra = self.cleaned_data.get('cp_ra')
        ley_ra = self.cleaned_data.get('ley_ra')
        decreto_ra = self.cleaned_data.get('ley_ra')
        otra_ra = self.cleaned_data.get('ley_ra')
    
        if cp_ra:
            if len(cp_ra) > 2000: 
                self._errors['cp_ra'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        if ley_ra:
            if len(ley_ra) > 2000: 
                self._errors['ley_ra'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        if decreto_ra:
            if len(decreto_ra) > 2000: 
                self._errors['decreto_ra'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        if otra_ra:
            if len(otra_ra) > 2000: 
                self._errors['otra_ra'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        return(self.cleaned_data)


## Modulo B pregunta 6
class EditMB_DocumentoMetodologRRAAForm(forms.ModelForm):
    
    otra_doc_cual = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm mb-1'}))

    class Meta:
        model = MB_DocumentoMetodologRRAA
        fields = ['otra_doc_cual']

    def clean(self):

        otra_doc_cual = self.cleaned_data.get('otra_doc_cual')
        
        if otra_doc_cual:
            if len(otra_doc_cual) > 2000: 
                self._errors['otra_doc_cual'] = self.error_class([ 
                    'Máximo 2000 caracteres requeridos'])

        return(self.cleaned_data)


# modulo B pregunta 8
class EditMB_ConceptosEstandarizadosRRAAForm(forms.ModelForm):
    
    org_in_cual = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm mb-1'}))

    ent_ordnac_cual = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm mb-1'}))

    leye_dec_cual = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm mb-1'}))

    otra_ce_cual = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm mb-1'}))

    class Meta:
        model = MB_ConceptosEstandarizadosRRAA
        fields = ['org_in_cual', 'ent_ordnac_cual', 'leye_dec_cual', 'otra_ce_cual' ]

    def clean(self):

        org_in_cual = self.cleaned_data.get('org_in_cual')
        ent_ordnac_cual = self.cleaned_data.get('ent_ordnac_cual')
        leye_dec_cual = self.cleaned_data.get('leye_dec_cual')
        otra_ce_cual = self.cleaned_data.get('otra_ce_cual')

        
        if org_in_cual:
            if len(org_in_cual) > 1000: 
                self._errors['org_in_cual'] = self.error_class([ 
                    'Máximo 1000 caracteres requeridos'])

        if ent_ordnac_cual:
            if len(ent_ordnac_cual) > 1000: 
                self._errors['ent_ordnac_cual'] = self.error_class([ 
                    'Máximo 1000 caracteres requeridos'])

        if leye_dec_cual:
            if len(leye_dec_cual) > 1000: 
                self._errors['leye_dec_cual'] = self.error_class([ 
                    'Máximo 1000 caracteres requeridos'])

        if otra_ce_cual:
            if len(otra_ce_cual) > 1000: 
                self._errors['otra_ce_cual'] = self.error_class([ 
                    'Máximo 1000 caracteres requeridos'])

        return(self.cleaned_data)


# Modulo B pregunta 9
class EditMB_ClasificacionesRRAAForm(forms.ModelForm):

    otra_cual_clas = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': 'Otra', 'class': 'form-control form-control-sm mb-1medio w-75 ml-3'}))

    no_pq = forms.CharField(required=False, widget=forms.TextInput(attrs={
                            'placeholder': '¿Por qué?', 'class': 'form-control form-control-sm mb-1medio'}))

    class Meta:
        model = MB_ClasificacionesRRAA
        fields = ['otra_cual_clas', 'no_pq']

    def clean(self):
            
        otra_cual_clas  = self.cleaned_data.get('otra_cual_clas')
        no_pq  = self.cleaned_data.get('no_pq')
        
        if otra_cual_clas:
            if len(otra_cual_clas) > 600: 
                self._errors['otra_cual_clas'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])

        if no_pq:
            if len(no_pq) > 600: 
                self._errors['no_pq'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)



# Modulo B pregunta 10
class EditMB_RecoleccionDatosRRAAForm(forms.ModelForm):

    sistema_inf_cual = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿Cuál?', 'class': 'form-control form-control-sm mb-1'}))

    otro_c = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': 'Otro (s) ¿cuál(es)?', 'class': 'form-control form-control-sm mb-1'}))

    class Meta:
        model = MB_RecoleccionDatosRRAA
        fields = ['sistema_inf_cual', 'otro_c']

    def clean(self):

        sistema_inf_cual = self.cleaned_data.get('sistema_inf_cual')
        otro_c = self.cleaned_data.get('otro_c')
        
        if sistema_inf_cual:
            if len(sistema_inf_cual) > 500: 
                self._errors['sistema_inf_cual'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])

        if otro_c:
            if len(otro_c) > 500: 
                self._errors['otro_c'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])

        return(self.cleaned_data)

# Modulo B pregunta 11
class EditMB_FrecuenciaRecoleccionDatoForm(forms.ModelForm):

    otra_cual_fre = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿cuál(es)?', 'class': 'form-control form-control-sm mb-1'}))

    class Meta:
        model = MB_FrecuenciaRecoleccionDato
        fields = ['otra_cual_fre']

    def clean(self):

        otra_cual_fre = self.cleaned_data.get('otra_cual_fre')
        
        if otra_cual_fre:
            if len(otra_cual_fre) > 800: 
                self._errors['otra_cual_fre'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        return(self.cleaned_data)

# Modulo B pregunta 12
class EditMB_HerramientasUtilProcesaForm(forms.ModelForm):

    otra_herram = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿cuál(es)?', 'class': 'form-control form-control-sm mb-1'}))

    class Meta:
        model = MB_HerramientasUtilProcesa
        fields = ['otra_herram']

    def clean(self):

        otra_herram = self.cleaned_data.get('otra_herram')
        
        if otra_herram:
            if len(otra_herram) > 800: 
                self._errors['otra_herram'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        return(self.cleaned_data)


# Modulo B pregunta 13
class EditMB_SeguridadInformForm(forms.ModelForm):

    otra_cual_s =  forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿cuál?', 'class': 'form-control form-control-sm mb-1'}))

    class Meta:
        model = MB_SeguridadInform
        fields = ['otra_cual_s']

    def clean(self):

        otra_cual_s = self.cleaned_data.get('otra_cual_s')
        
        if otra_cual_s:
            if len(otra_cual_s) > 800: 
                self._errors['otra_cual_s'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])

        return(self.cleaned_data)


# Modulo B pregunta 14
class EditMB_FrecuenciaAlmacenamientobdForm(forms.ModelForm):
    
    otra_alm_bd = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿cuál?', 'class': 'form-control form-control-sm mb-1'}))

    class Meta:
        model = MB_FrecuenciaAlmacenamientobd
        fields = ['otra_alm_bd']

    def clean(self):

        otra_alm_bd = self.cleaned_data.get('otra_alm_bd')
        
        if otra_alm_bd:
            if len(otra_alm_bd) > 500: 
                self._errors['otra_alm_bd'] = self.error_class([ 
                    'Máximo 500 caracteres requeridos'])

        return(self.cleaned_data)


    # Modulo B pregunta 15
class EditMB_CoberturaGeograficaRRAAForm(forms.ModelForm):

    cual_regio = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿Cuál (es)?', 'class': 'form-control form-control-sm mb-1'}))

    cual_depa = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿Cuál (es)?', 'class': 'form-control form-control-sm mb-1'}))

    cual_are_metrop = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿Cuál (es)?', 'class': 'form-control form-control-sm mb-1'}))

    cual_munic = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿Cuál (es)?', 'class': 'form-control form-control-sm mb-1'}))

    cual_otro = forms.CharField(required=False, widget=forms.TextInput(attrs={
                           'placeholder': '¿Cuál (es)?', 'class': 'form-control form-control-sm mb-1'}))

    class Meta:
        model = MB_CoberturaGeograficaRRAA
        fields = ['cual_regio', 'cual_depa', 'cual_are_metrop', 'cual_munic', 'cual_otro' ]

    def clean(self):

        cual_regio = self.cleaned_data.get('cual_regio')
        cual_depa = self.cleaned_data.get('cual_depa')
        cual_are_metrop = self.cleaned_data.get('cual_are_metrop')
        cual_munic = self.cleaned_data.get('cual_munic')
        cual_otro = self.cleaned_data.get('cual_otro')
        
        if cual_regio:
            if len(cual_regio) > 600: 
                self._errors['cual_regio'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])

        if cual_depa:
            if len(cual_depa) > 600: 
                self._errors['cual_depa'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])

        if cual_are_metrop:
            if len(cual_are_metrop) > 600: 
                self._errors['cual_are_metrop'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])
            
        if cual_munic:
            if len(cual_munic) > 1500: 
                self._errors['cual_munic'] = self.error_class([ 
                    'Máximo 1500 caracteres requeridos'])

        if cual_otro:
            if len(cual_otro) > 600: 
                self._errors['cual_otro'] = self.error_class([ 
                    'Máximo 600 caracteres requeridos'])


        return(self.cleaned_data)


# Modulo B pregunta 19

class EditMB_NoAccesoDatosForm(forms.ModelForm):

    otra_no_acceso = forms.CharField(required=False, widget=forms.TextInput(
        attrs={'placeholder': '¿Cuál (es)?', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = MB_NoAccesoDatos
        fields = ['otra_no_acceso']

    def clean(self):

        otra_no_acceso = self.cleaned_data.get('otra_no_acceso')
        
        if otra_no_acceso:
            if len(otra_no_acceso) > 800: 
                self._errors['otra_no_acceso'] = self.error_class([ 
                    'Máximo 800 caracteres requeridos'])


        return(self.cleaned_data)



## comentarios rraa

class CommentRRAAForm(forms.ModelForm):
    
    body = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Describa aquí las inconsistencias sobre la información diligenciada en el formulario', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = CommentRRAA
        fields = ['body']

    def clean(self):
            
        body = self.cleaned_data.get('body')
    
        if body:
            if len(body) > 1000: 
                self._errors['body'] = self.error_class([ 
                    'Máximo 1000 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data) 
                    

# formulario para actualización novedades

class NovedadActualizacionRRAAForm(forms.ModelForm):

    descrip_novedad = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Descripción de la novedad', 'class': 'form-control form-control-sm'}))

    novedad = forms.ModelChoiceField(required=False, label="tipo de novedad",  queryset=TipoNovedadRRAA.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))

    est_actualiz = forms.ModelChoiceField(required=False, label="estado de actualización",  queryset=EstadoActualizacionRRAA.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))

    obser_novedad = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Observaciones de la novedad', 'class': 'form-control form-control-sm'}))

    class Meta:
        model = NovedadActualizacionRRAA
        fields = ['descrip_novedad', 'novedad', 'est_actualiz', 'obser_novedad']


    def clean(self):
            
        descrip_novedad = self.cleaned_data.get('descrip_novedad')
        obser_novedad = self.cleaned_data.get('obser_novedad')
    
        if descrip_novedad:
            if len(descrip_novedad) > 1000: 
                self._errors['descrip_novedad'] = self.error_class([ 
                    'Máximo 1000 caracteres requeridos'])

        if obser_novedad:
            if len(obser_novedad) > 1000: 
                self._errors['obser_novedad'] = self.error_class([ 
                    'Máximo 1000 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)

## formulario para criticas

class CriticaRRAAForm(forms.ModelForm):

    observa_critica = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Observaciones de la crítica', 'class': 'form-control form-control-sm'}))

    estado_critica_ra = forms.ModelChoiceField(label="estado crítica", empty_label=None, queryset=EstadoCriticaRRAA.objects.all(
    ), widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))
    

    class Meta:
        model = CriticaRRAA
        fields = ['estado_critica_ra', 'observa_critica']

    def clean(self):
            
        observa_critica = self.cleaned_data.get('observa_critica')
    
        if observa_critica:
            if len(observa_critica) > 1000: 
                self._errors['observa_critica'] = self.error_class([ 
                    'Máximo 1000 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)


## crear Plan de fortalecimiento RRAA

class CreateFortalecimientoRRAAForm(forms.ModelForm):

    diagnostico_ra = forms.BooleanField(required=False, initial=False, widget=forms.CheckboxInput(
        attrs={"id": "id_diagnostico_ra", "data-toggle": "toggle", "data-size": "small", " data-on": "Si",  "data-off": "No", "data-onstyle": "success", "data-offstyle": "danger"}))

    year_diagnostico = forms.DateTimeField(
        input_formats=['%Y'],
        required=False,
        widget=forms.DateTimeInput(attrs={
            'class': 'form-control datetimepicker-input form-control-sm',
            'data-target': '##year_diagnostico'
        })
    )

    mod_sec_diagn = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Módulo o sección diagnosticado', 'class': 'form-control form-control-sm'}))

    plan_fort_aprob = forms.BooleanField(required=False, initial=False, widget=forms.CheckboxInput(
        attrs={"id": "id_plan_fort_aprob", "data-toggle": "toggle", "data-size": "small", " data-on": "Si",  "data-off": "No", "data-onstyle": "success", "data-offstyle": "danger"}))

    fecha_aprobacion = forms.DateTimeField(
        input_formats=['%d/%m/%Y'],
        required=False,
        widget=forms.DateTimeInput(attrs={
            'class': 'form-control datetimepicker-input form-control-sm',
            'data-target': '##fecha_aprobacion'
        })
    )

    seg_imple_plan = forms.ModelChoiceField(label="Seguimiento a la implementación del Plan de fortalecimiento", 
        empty_label=None, queryset=SeguimientoPlanFortalecimiento.objects.all(), 
        widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))

    fecha_inicio_plan = forms.DateTimeField(
        input_formats=['%d/%m/%Y'],
        required=False,
        widget=forms.DateTimeInput(attrs={
            'class': 'form-control datetimepicker-input form-control-sm',
            'data-target': '##fecha_inicio_plan'
        })
    )
    fecha_ultimo_seguimiento = forms.DateTimeField(
        input_formats=['%d/%m/%Y'],
        required=False,
        widget=forms.DateTimeInput(attrs={
            'class': 'form-control datetimepicker-input form-control-sm',
            'data-target': '##fecha_ultimo_seguimiento'
        })
    )
    fecha_finalizacion = forms.DateTimeField(
        input_formats=['%d/%m/%Y'],
        required=False,
        widget=forms.DateTimeInput(attrs={
            'class': 'form-control datetimepicker-input form-control-sm',
            'data-target': '##fecha_finalizacion'
        })
    )
   
    class Meta:
        model = FortalecimientoRRAA
        fields = [
            'diagnostico_ra',
            'year_diagnostico',
            'mod_sec_diagn',
            'plan_fort_aprob', 
            'fecha_aprobacion',
            'seg_imple_plan',
            'fecha_inicio_plan',
            'fecha_ultimo_seguimiento',
            'fecha_finalizacion',
        ]

    def clean(self):
            
        mod_sec_diagn = self.cleaned_data.get('mod_sec_diagn')
       
        if mod_sec_diagn:
            if len(mod_sec_diagn) > 1000: 
                self._errors['mod_sec_diagn'] = self.error_class([ 
                    'Máximo 1000 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)



## Editar Plan de fortalecimiento RRAA

class EditFortalecimientoRRAAForm(forms.ModelForm):

    diagnostico_ra = forms.BooleanField(required=False, initial=False, widget=forms.CheckboxInput(
        attrs={"id": "id_diagnostico_ra", "data-toggle": "toggle", "data-size": "small", " data-on": "Si",  "data-off": "No", "data-onstyle": "success", "data-offstyle": "danger"}))

    year_diagnostico = forms.DateTimeField(
        input_formats=['%Y'],
        required=False,
        widget=forms.DateTimeInput(attrs={
            'class': 'form-control datetimepicker-input form-control-sm',
            'data-target': '##year_diagnostico'
        })
    )

    mod_sec_diagn = forms.CharField(required=False, widget=forms.Textarea(
        attrs={'placeholder': 'Módulo o sección diagnosticado', 'class': 'form-control form-control-sm'}))

    plan_fort_aprob = forms.BooleanField(required=False, initial=False, widget=forms.CheckboxInput(
        attrs={"id": "id_plan_fort_aprob", "data-toggle": "toggle", "data-size": "small", " data-on": "Si",  "data-off": "No", "data-onstyle": "success", "data-offstyle": "danger"}))

    fecha_aprobacion = forms.DateTimeField(
        input_formats=['%d/%m/%Y'],
        required=False,
        widget=forms.DateTimeInput(attrs={
            'class': 'form-control datetimepicker-input form-control-sm',
            'data-target': '##fecha_aprobacion'
        })
    )

    seg_imple_plan = forms.ModelChoiceField(label="Seguimiento a la implementación del Plan de fortalecimiento", 
        empty_label=None, queryset=SeguimientoPlanFortalecimiento.objects.all(), 
        widget=forms.Select(attrs={'class': 'form-control form-control-sm'}))

    fecha_inicio_plan = forms.DateTimeField(
        input_formats=['%d/%m/%Y'],
        required=False,
        widget=forms.DateTimeInput(attrs={
            'class': 'form-control datetimepicker-input form-control-sm',
            'data-target': '##fecha_inicio_plan'
        })
    )
    fecha_ultimo_seguimiento = forms.DateTimeField(
        input_formats=['%d/%m/%Y'],
        required=False,
        widget=forms.DateTimeInput(attrs={
            'class': 'form-control datetimepicker-input form-control-sm',
            'data-target': '##fecha_ultimo_seguimiento'
        })
    )
    fecha_finalizacion = forms.DateTimeField(
        input_formats=['%d/%m/%Y'],
        required=False,
        widget=forms.DateTimeInput(attrs={
            'class': 'form-control datetimepicker-input form-control-sm',
            'data-target': '##fecha_finalizacion'
        })
    )
   
    class Meta:
        model = FortalecimientoRRAA
        fields = [
            'post_ra',
            'diagnostico_ra',
            'year_diagnostico',
            'mod_sec_diagn',
            'plan_fort_aprob', 
            'fecha_aprobacion',
            'seg_imple_plan',
            'fecha_inicio_plan',
            'fecha_ultimo_seguimiento',
            'fecha_finalizacion',
        ]

    def clean(self):
            
        mod_sec_diagn = self.cleaned_data.get('mod_sec_diagn')
       
        if mod_sec_diagn:
            if len(mod_sec_diagn) > 1000: 
                self._errors['mod_sec_diagn'] = self.error_class([ 
                    'Máximo 1000 caracteres requeridos'])

        # return any errors if found 
        return(self.cleaned_data)